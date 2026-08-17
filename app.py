import io
import json
import os
import re
from flask import Flask, request, jsonify, render_template, abort, send_file
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
import qrcode
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

load_dotenv()

app = Flask(__name__)

@app.template_filter("regex_search")
def regex_search_filter(value, pattern):
    m = re.search(pattern, str(value or ""))
    return m.group(1) if m and m.lastindex else (m.group(0) if m else "")

SHEET_ID        = os.getenv("SHEET_ID")
CREDS_PATH      = os.getenv("CREDENTIALS_PATH", "credentials.json")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID", "")


def _build_credentials() -> Credentials:
    """Load credentials from env JSON string (Vercel) or file path (local/Render)."""
    raw = os.getenv("GOOGLE_CREDENTIALS_JSON")
    if raw:
        info = json.loads(raw)
        return Credentials.from_service_account_info(info, scopes=SCOPES)
    return Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TAB_DISPLAY_NAMES = {
    "H1000-H1999":    "Hand Tools & Hardware",
    "E1000 - E1999":  "Electrical Hardware",
    "F1000-F1999":    "Fasteners",
    "W1000-W1999":    "PPE Equipment",
    "S1000-S1999":    "SOMA14",
    "S2000-S2999":    "SOMA50 for NSF",
    "S3000-S3999":    "SOMA50x12 for 2UR",
    "S4000-S4999":    "SOMA50x24 for 2UR",
    "M1000-M1999":    "Molds and Consumables",
    "D1000-D1999":    "SOMA5x5",
    "D2000-D2999":    "DARPA Press",
    "T1000-T1999":    "Aerogel Transport & Post-Processing",
    "P1000-P1999":    "Product and Mockups",
    "Y1000-Y1999":    "SOMA72",
    "R1000-R1999":    "Standards for Implementing Mold Features",
    "Z1000-Z1999":    "Full Line Integration",
}

TITLE_ROW_TABS    = {"W1000", "L1000"}
SUMMARY_TAB       = "Summary"
LOCATION_REF_TAB  = "Location Code Reference"

FIELD_ALIASES = {
    "part_id":        ["part id", "part id number", "id number", "part no", "part number", "id"],
    "part_name":      ["part name", "name", "description", "item name", "item", "component"],
    "category":       ["category", "type", "section", "group", "component type", "assembly", "system"],
    "location_code":  ["location code", "location", "loc code", "loc", "bin", "shelf"],
    "zone":           ["zone"],
    "qty":            ["qty", "quantity", "count", "on hand", "quantity on hand", "qty on hand", "stock qty", "stock"],
    "condition":      ["condition", "status", "state"],
    "reorder_thresh": ["reorder threshold", "reorder thresh", "min qty", "minimum", "threshold", "reorder level", "reorder point"],
    "reorder_flag":   ["reorder flag", "reorder", "flag", "needs reorder", "order needed"],
    "photo":          ["photo", "photo link", "photo or link", "image", "link", "url", "image link"],
    "qr_code":        ["qr code", "qr", "qr link", "qr url", "qr image"],
}

_ss_cache = None


def get_spreadsheet():
    global _ss_cache
    if _ss_cache is None:
        creds  = _build_credentials()
        client = gspread.authorize(creds)
        _ss_cache = client.open_by_key(SHEET_ID)
    return _ss_cache


def get_drive_service():
    return build("drive", "v3", credentials=_build_credentials())


def drive_file_id(url: str) -> str:
    m = re.search(r"(?:/d/|id=)([a-zA-Z0-9_-]{25,})", url or "")
    return m.group(1) if m else ""


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9 ]", "", s.strip().lower()).strip()


def detect_columns(headers: list[str]) -> dict[str, int]:
    col_map = {}
    for idx, h in enumerate(headers):
        n = norm(h)
        for field, aliases in FIELD_ALIASES.items():
            if field not in col_map and n in aliases:
                col_map[field] = idx
                break
    return col_map


def data_start_for(tab_name: str, rows: list[list[str]]) -> int:
    if tab_name in TITLE_ROW_TABS:
        return 2
    if rows:
        non_empty = sum(1 for c in rows[0] if c.strip())
        if non_empty <= 1:
            return 2
    return 1


def get_all_tabs() -> list[str]:
    """All worksheet titles except meta/reference tabs."""
    ss = get_spreadsheet()
    skip = {SUMMARY_TAB, LOCATION_REF_TAB}
    return [ws.title for ws in ss.worksheets() if ws.title not in skip]


def get_tab_meta() -> dict[str, dict]:
    ss         = get_spreadsheet()
    all_titles = [ws.title for ws in ss.worksheets()]
    meta       = {t: {"display_name": t, "description": ""} for t in all_titles}

    if SUMMARY_TAB in all_titles:
        try:
            rows = ss.worksheet(SUMMARY_TAB).get_all_values()
            for row in rows:
                for i, cell in enumerate(row):
                    if cell.strip() in all_titles and cell.strip() != SUMMARY_TAB:
                        tab  = cell.strip()
                        rest = [c.strip() for c in row[i+1:] if c.strip()]
                        if rest:              meta[tab]["display_name"] = rest[0]
                        if len(rest) > 1:     meta[tab]["description"]  = rest[1]
        except Exception:
            pass
    # Apply hardcoded display names (override sheet values).
    # Normalize by collapsing spaces around dashes so "E1000 - E1999" matches "E1000-E1999".
    def _norm(s): return s.replace(" - ", "-").replace("- ", "-").replace(" -", "-").strip()
    norm_to_real = {_norm(k): k for k in meta}
    for tab_key, name in TAB_DISPLAY_NAMES.items():
        real = norm_to_real.get(_norm(tab_key))
        if real:
            meta[real]["display_name"] = name

    return meta


def get_location_codes() -> list[str]:
    """Read all location codes from the Location Code Reference tab."""
    try:
        ss   = get_spreadsheet()
        rows = ss.worksheet(LOCATION_REF_TAB).get_all_values()
        if not rows:
            return []
        headers = [h.strip().lower() for h in rows[0]]
        try:
            col = next(i for i, h in enumerate(headers) if "location code" in h)
        except StopIteration:
            col = 0
        return [r[col].strip() for r in rows[1:] if col < len(r) and r[col].strip()]
    except Exception:
        return []


def save_location_code_if_new(code: str):
    """Append a custom location code to the reference tab if not already there."""
    if not code:
        return
    try:
        existing = get_location_codes()
        if code not in existing:
            ws = get_spreadsheet().worksheet(LOCATION_REF_TAB)
            ws.append_row([code], value_input_option="USER_ENTERED")
    except Exception:
        pass


def get_next_part_id(tab_name: str) -> str:
    """Detect the Part ID pattern in a tab and return the next suggested ID."""
    try:
        data = get_tab_items(tab_name)
        ids  = [item["fields"].get("part_id", "") for item in data["items"]
                if item["fields"].get("part_id")]
        if not ids:
            return ""

        parsed = []
        for pid in ids:
            m = re.match(r'^([A-Za-z]*)(\d+)$', pid.strip())
            if m:
                parsed.append((m.group(1).upper(), int(m.group(2))))

        if not parsed:
            return ""

        from collections import Counter
        dominant_prefix = Counter(p for p, _ in parsed).most_common(1)[0][0]
        nums = [n for p, n in parsed if p == dominant_prefix]
        return f"{dominant_prefix}{max(nums) + 1}"
    except Exception:
        return ""


def get_tab_items(tab_name: str) -> dict:
    ss    = get_spreadsheet()
    ws    = ss.worksheet(tab_name)
    rows  = ws.get_all_values()
    start = data_start_for(tab_name, rows)

    header_row = rows[start - 1] if start - 1 < len(rows) else []
    col_map    = detect_columns(header_row)

    def cellv(row, col):
        return row[col].strip() if col < len(row) else ""

    items = []
    for row_idx, row in enumerate(rows):
        if row_idx < start:
            continue
        fields = {field: (cellv(row, col_map[field]) if field in col_map else "")
                  for field in FIELD_ALIASES}
        identifier = fields.get("part_id") or fields.get("part_name")
        if not identifier:
            continue
        known_idxs = set(col_map.values())
        extra = [{"header": header_row[i], "value": cellv(row, i)}
                 for i in range(len(header_row))
                 if i not in known_idxs and header_row[i].strip()]
        items.append({
            "tab": tab_name, "sheet_row": row_idx + 1,
            "identifier": identifier, "fields": fields, "extra": extra,
        })

    return {"headers": header_row, "col_map": col_map, "items": items}


def find_item(identifier: str, tab_name: str) -> dict | None:
    data   = get_tab_items(tab_name)
    needle = identifier.strip().lower()
    for item in data["items"]:
        if item["identifier"].lower() == needle:
            return item
    return None


def compute_reorder_flag(qty_str: str, threshold_str: str) -> str:
    try:
        qty       = float(qty_str or "0")
        threshold = float(threshold_str)
        if threshold <= 0:
            return "N/A"
        return "Yes" if qty < threshold else "No"
    except (ValueError, TypeError):
        return "N/A"


def write_fields(tab: str, sheet_row: int, updates: dict[str, str], col_map: dict[str, int]):
    ws = get_spreadsheet().worksheet(tab)
    for field, value in updates.items():
        if field in col_map:
            ws.update_cell(sheet_row, col_map[field] + 1, value)


# ── Routes ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    all_tabs   = get_all_tabs()
    tab_meta   = get_tab_meta()
    active_tab = request.args.get("tab", all_tabs[0] if all_tabs else "")
    if active_tab not in all_tabs:
        active_tab = all_tabs[0] if all_tabs else ""
    data = get_tab_items(active_tab) if active_tab else {"items": [], "col_map": {}}
    return render_template("index.html",
        items=data["items"], col_map=data.get("col_map", {}),
        active_tab=active_tab, all_tabs=all_tabs, tab_meta=tab_meta)


@app.route("/item")
def item_page():
    identifier = request.args.get("id",  "").strip()
    tab        = request.args.get("tab", "").strip()
    if not identifier or not tab:
        abort(400, "Missing ?id= or ?tab= parameter")
    item = find_item(identifier, tab)
    if item is None:
        abort(404, f"'{identifier}' not found in {tab}")
    data           = get_tab_items(tab)
    col_map        = data["col_map"]
    location_codes = get_location_codes()
    return render_template("item.html",
        item=item, col_map=col_map, location_codes=location_codes,
        all_tabs=get_all_tabs(), tab_meta=get_tab_meta())


@app.route("/label/<tab>/<identifier>")
def label_page(tab, identifier):
    item = find_item(identifier, tab)
    if item is None:
        abort(404)
    data    = get_tab_items(tab)
    col_map = data["col_map"]
    return render_template("label.html", item=item, f=item["fields"], col_map=col_map)


@app.route("/label-image/<tab>/<path:identifier>.png")
def label_image(tab, identifier):
    """Generate a 300 DPI PNG of the label at exactly 62mm × height_mm."""
    from PIL import Image, ImageDraw, ImageFont

    item = find_item(identifier, tab)
    if item is None:
        abort(404)

    f           = item["fields"]
    DPI         = 300
    MM_PER_INCH = 25.4

    def mm2px(mm):
        return int(round(mm / MM_PER_INCH * DPI))

    def _f(key, default, lo, hi):
        try: return max(lo, min(hi, float(request.args.get(key, default))))
        except ValueError: return default

    height_mm    = _f("h",  22,  15, 60)
    qr_mm        = _f("qr", 18,   8, 50)
    font_name_mm = _f("fn",  4,   2, 12)
    font_loc_mm  = _f("fl",  4,   2, 10)
    font_id_mm   = _f("fi", 3.5,  2,  8)

    W = mm2px(62)
    H = mm2px(height_mm)

    # Generate QR
    base_url   = request.host_url.rstrip("/")
    target_url = f"{base_url}/item?id={identifier}&tab={tab}"
    qr_obj = qrcode.QRCode(version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10, border=2)
    qr_obj.add_data(target_url)
    qr_obj.make(fit=True)
    qr_img = qr_obj.make_image(fill_color="black", back_color="white").convert("RGB")
    qr_px  = mm2px(qr_mm)
    qr_img = qr_img.resize((qr_px, qr_px), Image.LANCZOS)

    img  = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    pad = mm2px(2)
    # Place QR centred vertically on left
    qr_y = (H - qr_px) // 2
    img.paste(qr_img, (pad, qr_y))

    # Text area
    text_x  = pad + qr_px + mm2px(2)
    text_w  = W - text_x - pad

    def best_font(size_mm):
        size_px = int(round(mm2px(size_mm)))
        font_paths = [
            "/System/Library/Fonts/Helvetica.ttc",                          # macOS
            "/System/Library/Fonts/Arial.ttf",                               # macOS alt
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",          # Debian/Ubuntu
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",  # CentOS/Amazon Linux
            "/usr/share/fonts/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
        ]
        for path in font_paths:
            try:
                return ImageFont.truetype(path, size_px)
            except Exception:
                continue
        # Pillow 10+ built-in scalable font
        try:
            return ImageFont.load_default(size=size_px)
        except TypeError:
            return ImageFont.load_default()

    def wrap_text(text, font, max_width):
        """Split text into lines that fit within max_width pixels."""
        words = text.split()
        wrapped = []
        current = ""
        for word in words:
            test = (current + " " + word).strip()
            bb = draw.textbbox((0, 0), test, font=font)
            if bb[2] - bb[0] <= max_width:
                current = test
            else:
                if current:
                    wrapped.append(current)
                current = word
        if current:
            wrapped.append(current)
        return wrapped if wrapped else [text]

    # Build render list: (text_line, font, color)
    render_lines = []
    if f.get("part_name"):
        font_name = best_font(font_name_mm)
        for ln in wrap_text(f["part_name"], font_name, text_w):
            render_lines.append((ln, font_name, "#000000"))
    if f.get("location_code"):
        font_loc = best_font(font_loc_mm)
        for ln in wrap_text(f["location_code"], font_loc, text_w):
            render_lines.append((ln, font_loc, "#000000"))
    font_id = best_font(font_id_mm)
    render_lines.append((identifier, font_id, "#666666"))

    gap = mm2px(1)
    line_heights = []
    for text, font, _ in render_lines:
        bb = draw.textbbox((0, 0), text, font=font)
        line_heights.append(bb[3] - bb[1])

    total_h   = sum(line_heights) + gap * (len(render_lines) - 1)
    current_y = (H - total_h) // 2

    for i, (text, font, color) in enumerate(render_lines):
        bb = draw.textbbox((0, 0), text, font=font)
        draw.text((text_x, current_y - bb[1]), text, fill=color, font=font)
        current_y += line_heights[i] + gap

    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(DPI, DPI))
    buf.seek(0)
    return send_file(buf, mimetype="image/png",
                     download_name=f"label-{identifier}.png")


@app.route("/dashboard")
def dashboard():
    all_tabs = get_all_tabs()
    tab_meta = get_tab_meta()
    return render_template("dashboard.html", all_tabs=all_tabs, tab_meta=tab_meta)


@app.route("/api/dashboard-stats")
def api_dashboard_stats():
    all_tabs      = get_all_tabs()
    tab_counts    = {}
    reorder_items = []
    missing_items = []
    needs_repair  = []
    total         = 0

    for tab in all_tabs:
        try:
            data  = get_tab_items(tab)
            items = data["items"]
            tab_counts[tab] = len(items)
            total += len(items)

            for item in items:
                f   = item["fields"]
                id_ = item["identifier"]

                if f.get("reorder_flag", "").strip().lower() == "yes":
                    reorder_items.append({"tab": tab, "id": id_, "name": f.get("part_name", "")})

                missing = []
                if not f.get("photo"):         missing.append("photo")
                if not f.get("location_code"): missing.append("location")
                if not f.get("qty"):           missing.append("qty")
                if missing:
                    missing_items.append({"tab": tab, "id": id_, "name": f.get("part_name", ""), "missing": missing})

                if f.get("condition", "").strip().lower() == "needs repair":
                    needs_repair.append({"tab": tab, "id": id_, "name": f.get("part_name", "")})
        except Exception:
            tab_counts[tab] = 0

    return jsonify({
        "total": total,
        "tab_counts": tab_counts,
        "reorder_items": reorder_items,
        "missing_items": missing_items,
        "needs_repair": needs_repair,
    })


@app.route("/add")
def add_page():
    all_tabs   = get_all_tabs()
    tab_meta   = get_tab_meta()
    active_tab = request.args.get("tab", all_tabs[0] if all_tabs else "")
    if active_tab not in all_tabs:
        active_tab = all_tabs[0] if all_tabs else ""
    data           = get_tab_items(active_tab) if active_tab else {"col_map": {}, "headers": []}
    location_codes = get_location_codes()
    next_part_id   = get_next_part_id(active_tab) if active_tab else ""
    return render_template("add.html",
        active_tab=active_tab, all_tabs=all_tabs, tab_meta=tab_meta,
        col_map=data.get("col_map", {}), headers=data.get("headers", []),
        location_codes=location_codes, next_part_id=next_part_id)


@app.route("/qr/<tab>/<path:identifier>.png")
def qr_image(tab, identifier):
    base       = request.host_url.rstrip("/")
    target_url = f"{base}/item?id={identifier}&tab={tab}"
    qr = qrcode.QRCode(version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10, border=2)
    qr.add_data(target_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return send_file(buf, mimetype="image/png")


@app.route("/photo/<file_id>")
def proxy_photo(file_id):
    """Stream a Drive image through Flask — no public sharing required."""
    try:
        service  = get_drive_service()
        meta     = service.files().get(fileId=file_id, fields="mimeType",
                       supportsAllDrives=True).execute()
        mimetype = meta.get("mimeType", "image/jpeg")
        content  = service.files().get_media(fileId=file_id).execute()
        return send_file(io.BytesIO(content), mimetype=mimetype)
    except Exception as exc:
        app.logger.warning("Photo proxy failed for %s: %s", file_id, exc)
        abort(404)


# ── API ──────────────────────────────────────────────────────────────────────

@app.route("/api/save", methods=["POST"])
def api_save():
    body      = request.get_json(force=True)
    tab       = body.get("tab", "").strip()
    sheet_row = int(body.get("sheet_row", 0))
    qty           = str(body.get("qty", "")).strip()
    condition     = body.get("condition", "").strip()
    threshold     = str(body.get("threshold", "")).strip()
    location_code = body.get("location_code", "").strip()

    if not tab or sheet_row <= 0:
        return jsonify({"ok": False, "error": "Invalid tab or row"}), 400
    if condition not in {"Good", "Fair", "Needs Repair", "N/A", ""}:
        return jsonify({"ok": False, "error": "Invalid condition"}), 400

    data    = get_tab_items(tab)
    col_map = data["col_map"]
    reorder_flag = compute_reorder_flag(qty, threshold) if threshold else None

    updates = {}
    if qty:            updates["qty"]           = qty
    if condition:      updates["condition"]      = condition
    if threshold:      updates["reorder_thresh"] = threshold
    if reorder_flag:   updates["reorder_flag"]   = reorder_flag
    if location_code:  updates["location_code"]  = location_code

    try:
        write_fields(tab, sheet_row, updates, col_map)
        return jsonify({"ok": True, "reorder_flag": reorder_flag})
    except Exception as exc:
        app.logger.exception("Sheet write failed")
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/save-link", methods=["POST"])
def api_save_link():
    body      = request.get_json(force=True)
    tab       = body.get("tab", "").strip()
    sheet_row = int(body.get("sheet_row", 0))
    link      = body.get("link", "").strip()
    if not tab or sheet_row <= 0:
        return jsonify({"ok": False, "error": "Invalid tab or row"}), 400
    data    = get_tab_items(tab)
    col_map = data["col_map"]
    try:
        write_fields(tab, sheet_row, {"photo": link}, col_map)
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/clear-photo", methods=["POST"])
def api_clear_photo():
    body      = request.get_json(force=True)
    tab       = body.get("tab", "").strip()
    sheet_row = int(body.get("sheet_row", 0))
    photo_url = body.get("photo_url", "").strip()
    if not tab or sheet_row <= 0:
        return jsonify({"ok": False, "error": "Invalid tab or row"}), 400

    fid = drive_file_id(photo_url)
    if fid:
        try:
            get_drive_service().files().delete(fileId=fid, supportsAllDrives=True).execute()
        except Exception:
            pass

    data    = get_tab_items(tab)
    col_map = data["col_map"]
    try:
        write_fields(tab, sheet_row, {"photo": ""}, col_map)
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/upload-photo", methods=["POST"])
def api_upload_photo():
    if not DRIVE_FOLDER_ID or DRIVE_FOLDER_ID == "paste_your_folder_id_here":
        return jsonify({"ok": False, "error": "Drive folder not configured"}), 400

    tab        = request.form.get("tab", "").strip()
    sheet_row  = int(request.form.get("sheet_row", 0))
    identifier = request.form.get("identifier", "").strip()
    if not tab or sheet_row <= 0 or not identifier:
        return jsonify({"ok": False, "error": "Missing required fields"}), 400
    if "photo" not in request.files or not request.files["photo"].filename:
        return jsonify({"ok": False, "error": "No file received"}), 400

    file     = request.files["photo"]
    ext      = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else "jpg"
    mimetype = file.content_type or "image/jpeg"
    file_buf = file.read()

    try:
        service = get_drive_service()

        # Find all existing files for this identifier (any extension)
        existing = service.files().list(
            q=f"name contains '{identifier}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false",
            fields="files(id,name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute().get("files", [])

        # Number the new file: E1000.png → E1000 (1).png → E1000 (2).png …
        count    = len(existing)
        filename = f"{identifier}.{ext}" if count == 0 else f"{identifier} ({count}).{ext}"

        media    = MediaIoBaseUpload(io.BytesIO(file_buf), mimetype=mimetype, resumable=False)
        uploaded = service.files().create(
            body={"name": filename, "parents": [DRIVE_FOLDER_ID]},
            media_body=media,
            fields="id,webViewLink",
            supportsAllDrives=True,
        ).execute()

        file_id  = uploaded["id"]
        view_url = uploaded.get("webViewLink", f"https://drive.google.com/file/d/{file_id}/view")

        # Write Drive URL to sheet
        data    = get_tab_items(tab)
        col_map = data["col_map"]
        write_fields(tab, sheet_row, {"photo": view_url}, col_map)

        return jsonify({"ok": True, "url": view_url,
                        "thumbnail": f"/photo/{file_id}", "file_id": file_id})
    except Exception as exc:
        app.logger.exception("Drive upload failed")
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/add-item", methods=["POST"])
def api_add_item():
    body   = request.get_json(force=True)
    tab    = body.get("tab", "").strip()
    fields = body.get("fields", {})

    if not tab:
        return jsonify({"ok": False, "error": "Missing tab"}), 400

    data       = get_tab_items(tab)
    col_map    = data["col_map"]
    headers    = data["headers"]
    num_cols   = len(headers)
    row        = [""] * num_cols

    for field_key, value in fields.items():
        if field_key in col_map:
            row[col_map[field_key]] = str(value).strip()

    # Auto-compute reorder flag
    qty       = fields.get("qty", "")
    threshold = fields.get("reorder_thresh", "")
    if "reorder_flag" in col_map and qty and threshold:
        row[col_map["reorder_flag"]] = compute_reorder_flag(str(qty), str(threshold))

    try:
        ws = get_spreadsheet().worksheet(tab)
        ws.append_row(row, value_input_option="USER_ENTERED")
        identifier = fields.get("part_id") or fields.get("part_name") or ""
        # Persist custom location codes back to the reference tab
        save_location_code_if_new(fields.get("location_code", ""))
        return jsonify({"ok": True, "identifier": identifier})
    except Exception as exc:
        app.logger.exception("Add item failed")
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/location-codes")
def api_location_codes():
    """Return location codes live from the reference tab."""
    return jsonify(get_location_codes())


@app.route("/api/test-drive")
def api_test_drive():
    result = {"folder_id": DRIVE_FOLDER_ID, "steps": []}
    try:
        service = get_drive_service()
        result["steps"].append("✓ Drive service created")
        about = service.about().get(fields="user").execute()
        result["service_account"] = about.get("user", {}).get("emailAddress", "unknown")
        result["steps"].append(f"✓ Authenticated as {result['service_account']}")
        folder = service.files().get(fileId=DRIVE_FOLDER_ID, fields="id,name",
                     supportsAllDrives=True).execute()
        result["steps"].append(f"✓ Folder found: '{folder.get('name')}'")
        result["ok"] = True
    except Exception as exc:
        result["steps"].append(f"✗ Error: {str(exc)}")
        result["ok"] = False
    return jsonify(result)


if __name__ == "__main__":
    app.run(debug=True)
