# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "flask>=3.0",
#   "curl-cffi>=0.7",
#   "openpyxl>=3.1",
# ]
# ///
"""Menu Dashboard — Keeta BD Team.

Paste a restaurant menu link (HungerStation today; Jahez adapter ready to
plug in) and get the full menu: categories, items, descriptions, prices,
calories and images — on screen + styled Excel export. Optional Google
Drive image upload per teammate's own Drive.

Local run:      uv run app.py serve          (or double-click start.bat)
Hosted run:     ACCESS_PASSWORD + SECRET_KEY env vars, gunicorn app:app
                (see Dockerfile) — teammates log in with the shared password.
CLI check:      uv run app.py test <url> [--export file.xlsx]
"""
from __future__ import annotations

import io
import json
import os
import re
import secrets
import sys
import threading
import webbrowser
from datetime import datetime
from urllib.parse import quote, urlparse

from curl_cffi import requests as cf
from flask import Flask, Response, jsonify, redirect, render_template_string, request as flask_request, session, url_for

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

# --- hosting config (env) -------------------------------------------------
ACCESS_PASSWORD = os.environ.get("ACCESS_PASSWORD", "").strip()   # set = password login enabled
HOST_BIND = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "8787"))

app = Flask(__name__)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    PERMANENT_SESSION_LIFETIME=60 * 60 * 24 * 30,  # 30 days
)


def _persistent_secret() -> str:
    """Session signing key: env first, else generated once and persisted."""
    env = os.environ.get("SECRET_KEY")
    if env:
        return env
    try:
        cfg = {}
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                cfg = json.load(f) or {}
        if cfg.get("secret_key"):
            return str(cfg["secret_key"])
        cfg["secret_key"] = secrets.token_hex(32)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return cfg["secret_key"]
    except Exception:
        return secrets.token_hex(32)


app.secret_key = _persistent_secret()


# --------------------------------------------------------------------------
# Auth gate (only active when ACCESS_PASSWORD is set, e.g. on the hosted copy)
# --------------------------------------------------------------------------

@app.before_request
def _gate():
    if not ACCESS_PASSWORD:
        return None
    if flask_request.path in ("/login", "/healthz"):
        return None
    if not session.get("team_ok"):
        return redirect("/login")
    return None


@app.get("/healthz")
def healthz():
    return jsonify({"ok": True})


@app.get("/login")
def login_page():
    if not ACCESS_PASSWORD:
        return redirect("/")
    return Response(LOGIN_PAGE, mimetype="text/html; charset=utf-8")


@app.post("/login")
def login_post():
    pwd = flask_request.form.get("password", "")
    if pwd and secrets.compare_digest(pwd, ACCESS_PASSWORD):
        session["team_ok"] = True
        session.permanent = True
        return redirect("/")
    return render_template_string(LOGIN_PAGE, wrong=1)


@app.get("/logout")
def logout():
    session.clear()
    return redirect("/login")


# --------------------------------------------------------------------------
# Fetching & parsing (multi-platform registry)
# --------------------------------------------------------------------------

def _normalize_url(url: str) -> str:
    url = (url or "").strip()
    if not url:
        raise ValueError("اكتب اللينك الأول")
    if not url.lower().startswith(("http://", "https://")):
        url = "https://" + url
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()
    base = host[4:] if host.startswith("www.") else host
    if base == "jahez.net":
        raise ValueError(
            "جاهز شغّالة قريب 👷 — ابعتلي لينك مشاركة حقيقي لمطعم من تطبيق جاهز "
            "(اللي بيبدأ بـ jahez.page.link) وأنا أوصّل المنصة في نفس اليوم")
    if base != "hungerstation.com":
        raise ValueError(
            "اللينك لازم يكون من موقع hungerstation.com — ولو ده لينك جاهز ابعتلي مثال حقيقي له")
    return url


def _fetch_through_cloudscraper(url: str) -> str | None:
    """Fallback: solve Cloudflare's JS challenge with cloudscraper.
    Slower than curl-cffi but survives tougher bot rules."""
    try:
        import cloudscraper
    except ImportError:
        return None
    try:
        s = cloudscraper.create_scraper(
            browser={"browser": "chrome", "platform": "windows", "mobile": False})
        r = s.get(url, timeout=90)
        if "__NEXT_DATA__" in (r.text or ""):
            return r.text
    except Exception:
        pass
    return None


def _looks_like_menu_html(text: str) -> bool:
    return bool(text) and ('id="__NEXT_DATA__"' in text or "__NEXT_DATA__" in text)


def fetch_menu(url: str) -> dict:
    """Fetch a HungerStation restaurant page and return a structured menu."""
    url = _normalize_url(url)
    r = cf.get(url, impersonate="chrome", timeout=90)
    if r.status_code in (404, 410):
        raise RuntimeError("الصفحة دي مش موجودة — المطعم ممكن يكون اتقفل من HungerStation أو اللينك غلط")
    text = r.text if r.status_code == 200 else ""
    if r.status_code != 200 or not _looks_like_menu_html(text):
        # Cloud provider IPs are often blocked by HungerStation's bot protection;
        # retry with a full Cloudflare challenge solver before giving up.
        alt = _fetch_through_cloudscraper(url)
        if _looks_like_menu_html(alt or ""):
            text = alt
        elif r.status_code != 200:
            raise RuntimeError(f"الموقع رجّع كود {r.status_code} — جرب تاني بعد شوية")
    low = text.lower()
    if "just a moment" in low or "attention required" in low:
        raise RuntimeError("الصفحة محمية بحماية Cloudflare — استنى دقيقة وجرب تاني")

    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', text, re.S)
    if not m:
        raise RuntimeError("مش لاقي بيانات المنيو في الصفحة دي — اتأكد إن اللينك لصفحة مطعم")
    data = json.loads(m.group(1))
    props = (data.get("props") or {}).get("pageProps") or {}
    vendor = props.get("vendor") or {}
    raw_menu = props.get("vendorMenu") or []
    if not raw_menu:
        raise RuntimeError("الصفحة دي مفيهاش منيو — اتأكد إن اللينك صح")

    categories = []
    flat: list[dict] = []
    for group in raw_menu:
        gname = ((group.get("name") or "")).strip()
        items = []
        for it in group.get("items") or []:
            price_raw = it.get("price")
            try:
                price = float(price_raw) if price_raw not in (None, "") else None
            except (TypeError, ValueError):
                price = None
            item = {
                "name": ((it.get("name") or "")).strip(),
                "description": ((it.get("description") or "")).strip(),
                "price": price,
                "calories": it.get("calories"),
                "category": gname,
                "image": ((it.get("image") or "")).strip(),
            }
            if item["name"]:
                items.append(item)
                flat.append(item)
        if items:
            categories.append({"name": gname, "items": items})

    if not categories:
        raise RuntimeError("المنيو فاضي في الصفحة دي")

    return {
        "url": url,
        "platform": "HungerStation",
        "restaurant": {
            "name": (vendor.get("chain_name") or "").strip(),
            "cuisines": vendor.get("kitchens") or [],
            "rating": vendor.get("average_rating") or "",
            "rating_count": vendor.get("rate_count") or 0,
            "status": vendor.get("status") or "",
        },
        "categories": categories,
        "total_items": len(flat),
    }


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------
# Column template (header, item key). Easy to remap later to match the
# Keeta onboarding team's exact template.
EXCEL_COLUMNS = [
    ("Category", "category"),
    ("Item Name", "name"),
    ("Description", "description"),
    ("Price (SR)", "price"),
    ("Calories", "calories"),
    ("Image Link", "image_link"),
]

BRAND = "FF7420"
LIGHT = "FFF1E8"


def build_workbook(menu: dict):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Menu"
    rest = menu.get("restaurant") or {}
    n_cols = len(EXCEL_COLUMNS)
    last_col = get_column_letter(n_cols)

    # -- meta header --
    ws.merge_cells(f"A1:{last_col}1")
    a1 = ws["A1"]
    a1.value = rest.get("name") or "Menu"
    a1.font = Font(bold=True, size=14, color="FFFFFF")
    a1.fill = PatternFill("solid", fgColor=BRAND)
    a1.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    cuisines = rest.get("cuisines") or []
    rating = rest.get("rating") or ""
    cats = menu.get("categories") or []
    meta2 = " • ".join(
        part for part in [
            "Platform: " + (menu.get("platform") or "HungerStation"),
            "Cuisines: " + ", ".join(cuisines) if cuisines else "",
            f"Rating: {rating} ({rest.get('rating_count', 0)})" if rating else "",
            f"Categories: {len(cats)}  •  Items: {menu.get('total_items', 0)}",
        ] if part
    )
    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"] = meta2

    ws.merge_cells(f"A3:{last_col}3")
    ws["A3"] = f"Source: {menu.get('url', '')}"
    ws.merge_cells(f"A4:{last_col}4")
    ws["A4"] = f"Extracted: {datetime.now():%Y-%m-%d %H:%M} — Menu Dashboard (Keeta BD Team)"

    # -- table header --
    header_row = 6
    for col, (title, _key) in enumerate(EXCEL_COLUMNS, start=1):
        c = ws.cell(row=header_row, column=col, value=title)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BRAND)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[header_row].height = 20

    thin = Side(style="thin", color="FFD9C6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # -- data rows --
    row = header_row + 1
    for group in cats:
        for it in group.get("items", []):
            for col, (_title, key) in enumerate(EXCEL_COLUMNS, start=1):
                val = it.get(key)
                if val is None:
                    val = ""
                cell = ws.cell(row=row, column=col, value=val)
                cell.border = border
                cell.alignment = Alignment(
                    vertical="top", wrap_text=(key == "description"))
                if key == "price" and isinstance(val, float):
                    cell.number_format = "0.00"
                if key == "calories" and not isinstance(val, int):
                    try:
                        cell.value = int(val)
                    except (TypeError, ValueError):
                        cell.value = val
                if key == "image_link" and val:
                    cell.hyperlink = val
                    cell.font = Font(color="1155CC", underline="single")
            row += 1

    widths = {"category": 26, "name": 42, "description": 64,
              "price": 12, "calories": 12, "image_link": 52}
    for col, (_t, key) in enumerate(EXCEL_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = widths.get(key, 20)

    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:{last_col}{row - 1}"
    return wb


# --------------------------------------------------------------------------
# UI — login (used only on the hosted copy)
# --------------------------------------------------------------------------

LOGIN_PAGE = r"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>دخول — داشبورد المنيوهات</title>
<style>
  :root { --brand:#ff7420; --brand-dark:#e45f0e; --border:#f0dcd0; --ink:#23272b; }
  * { box-sizing:border-box; }
  body { font-family:"Segoe UI",Tahoma,Arial,sans-serif; margin:0; min-height:100vh;
         display:flex; align-items:center; justify-content:center; background:#faf7f5; color:var(--ink); }
  .card { background:#fff; border:1px solid var(--border); border-radius:16px; padding:34px 30px;
          width:360px; box-shadow:0 6px 24px rgba(0,0,0,.06); text-align:center; }
  .logo { font-size:40px; }
  h1 { font-size:20px; margin:10px 0 4px; }
  p.sub { color:#8a8f98; font-size:13px; margin:0 0 20px; }
  input { width:100%; padding:13px 14px; font-size:15px; border:2px solid var(--border);
          border-radius:10px; outline:none; text-align:center; }
  input:focus { border-color:var(--brand); }
  button { width:100%; margin-top:12px; padding:13px; font-size:15px; font-weight:bold;
           color:#fff; background:var(--brand); border:none; border-radius:10px; cursor:pointer; }
  button:hover { background:var(--brand-dark); }
  .wrong { background:#fdecea; color:#b3261e; border:1px solid #f5c6c2; padding:10px;
           border-radius:10px; font-size:13px; margin-bottom:14px; }
</style>
</head>
<body>
  <div class="card">
    <div class="logo">&#127869;&#65039;</div>
    <h1>داشبورد المنيوهات — Keeta BD</h1>
    <p class="sub">منيوهات HungerStation و Excel جاهز — للتيم بس</p>
    {% if wrong %}<div class="wrong">الباسورد مش صح — جرب تاني</div>{% endif %}
    <form method="post" action="/login">
      <input type="password" name="password" placeholder="الباسورد بتاع التيم" autofocus>
      <button type="submit">دخول</button>
    </form>
  </div>
</body>
</html>
"""


# --------------------------------------------------------------------------
# UI — dashboard
# --------------------------------------------------------------------------

PAGE = r"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>داشبورد المنيوهات — Keeta BD Team</title>
<style>
  :root {
    --brand: #ff7420; --brand-dark: #e45f0e; --light: #fff1e8;
    --ink: #23272b; --muted: #8a8f98; --border: #f0dcd0; --green: #1e7b34;
  }
  * { box-sizing: border-box; }
  body { font-family: "Segoe UI", Tahoma, Arial, sans-serif; margin: 0; background: #faf7f5; color: var(--ink); }
  header { background: linear-gradient(135deg, var(--brand), #ff9752); color: #fff; padding: 22px 20px; text-align: center; }
  header h1 { margin: 0 0 6px; font-size: 24px; }
  header p { margin: 0; opacity: .92; font-size: 13.5px; }
  header .logout { color:#fff; opacity:.85; font-size:12px; text-decoration:none; }
  main { max-width: 1050px; margin: 0 auto; padding: 20px 16px 60px; }
  .input-row { display: flex; gap: 10px; margin-bottom: 8px; }
  .input-row input { flex: 1; padding: 13px 14px; font-size: 15px; border: 2px solid var(--border); border-radius: 10px; outline: none; direction: ltr; text-align: left; }
  .input-row input:focus { border-color: var(--brand); }
  .input-row button { padding: 13px 26px; font-size: 15px; font-weight: bold; color: #fff; background: var(--brand); border: none; border-radius: 10px; cursor: pointer; }
  .input-row button:hover { background: var(--brand-dark); }
  .input-row button:disabled { background: #c9c9c9; cursor: wait; }
  .hint { color: var(--muted); font-size: 12px; margin-bottom: 14px; }
  #status { margin: 12px 0; font-size: 14px; }
  .error { background: #fdecea; color: #b3261e; border: 1px solid #f5c6c2; padding: 12px 14px; border-radius: 10px; }
  .loading { color: var(--brand-dark); font-weight: bold; }
  .ok { background: #eef4ee; color: var(--green); border: 1px solid #cfe3cf; padding: 12px 14px; border-radius: 10px; }
  details.settings { background: #fff; border: 1px solid var(--border); border-radius: 12px; padding: 12px 16px; margin-bottom: 16px; }
  details.settings summary { cursor: pointer; font-weight: bold; color: var(--brand-dark); }
  details.settings .inner { margin-top: 12px; font-size: 13.5px; }
  details.settings ol { margin: 8px 0 10px; padding-inline-start: 22px; }
  details.settings li { margin-bottom: 4px; }
  pre#gasCode { direction: ltr; text-align: left; background: #23272b; color: #e8e8e8; padding: 14px; border-radius: 10px; overflow: auto; font-size: 12px; max-height: 300px; }
  .settings-row { display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap; }
  .settings-row input { flex: 1; min-width: 260px; padding: 10px 12px; font-size: 13px; border: 2px solid var(--border); border-radius: 10px; direction: ltr; text-align: left; outline: none; }
  .settings-row input:focus { border-color: var(--brand); }
  .btn-sm { padding: 10px 18px; font-size: 13px; font-weight: bold; border: none; border-radius: 10px; cursor: pointer; }
  .btn-sm.primary { background: var(--brand); color: #fff; }
  .btn-sm.plain { background: var(--light); color: var(--brand-dark); }
  .btn-sm:hover { opacity: .9; }
  .rest-card { background: #fff; border: 1px solid var(--border); border-radius: 14px; padding: 16px 20px; margin-bottom: 14px; box-shadow: 0 2px 8px rgba(0,0,0,.04); }
  .rest-card h2 { margin: 0 0 6px; font-size: 22px; }
  .rest-card .meta { color: var(--muted); font-size: 13px; }
  .stats { display: flex; gap: 12px; flex-wrap: wrap; margin: 14px 0; }
  .stat { flex: 1; min-width: 120px; background: var(--light); border-radius: 12px; padding: 10px 12px; text-align: center; }
  .stat .v { font-size: 19px; font-weight: bold; color: var(--brand-dark); }
  .stat .l { font-size: 12px; color: var(--muted); margin-top: 2px; }
  .toolbar { display: flex; gap: 10px; margin: 12px 0; flex-wrap: wrap; align-items: center; }
  .toolbar input { flex: 1; min-width: 200px; padding: 10px 12px; font-size: 14px; border: 2px solid var(--border); border-radius: 10px; outline: none; }
  .toolbar input:focus { border-color: var(--brand); }
  .toolbar .btn { text-decoration: none; padding: 10px 18px; border-radius: 10px; font-weight: bold; font-size: 13.5px; cursor: pointer; border: none; }
  .toolbar .btn.green { background: var(--green); color: #fff; }
  .toolbar .btn.green:hover { background: #156427; }
  .toolbar .btn.blue { background: #1a73e8; color: #fff; }
  .toolbar .btn.blue:hover { background: #1257c0; }
  .toolbar .btn:disabled { background: #c9c9c9; cursor: not-allowed; }
  .cat h3 { margin: 24px 0 10px; font-size: 17px; color: var(--brand-dark); border-inline-start: 4px solid var(--brand); padding-inline-start: 10px; }
  table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 10px; overflow: hidden; }
  th, td { padding: 8px 10px; font-size: 13.5px; border-bottom: 1px solid var(--border); text-align: start; vertical-align: top; }
  th { background: var(--light); color: var(--brand-dark); font-size: 13px; }
  td.price { font-weight: bold; white-space: nowrap; }
  .thumb { width: 58px; height: 58px; object-fit: cover; border-radius: 8px; background: #f3f3f3; }
  .drv-tag { display: inline-block; margin-top: 3px; font-size: 10.5px; background: #e8f0fe; color: #1a73e8; padding: 1px 8px; border-radius: 999px; font-weight: bold; }
  .empty { color: var(--muted); text-align: center; padding: 30px; }
  .badge { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 11.5px; background: #eef4ee; color: var(--green); margin-inline-start: 8px; font-weight: bold; }
  .badge.closed { background: #fdecea; color: #b3261e; }
</style>
</head>
<body>
<header>
  <h1>داشبورد المنيوهات</h1>
  <p>حط لينك المنيو (HungerStation) واسحب المنيو كامل زي ما فريق الـ Onboarding بيعمل — مع رفع الصور على Google Drive</p>
</header>
<main>
  <div class="input-row">
    <input id="url" type="text" spellcheck="false"
      placeholder="https://www.hungerstation.com/sa-en/restaurant/riyadh/riyadh/...">
    <button id="go">اسحب المنيو</button>
  </div>
  <div class="hint">مثال: https://www.hungerstation.com/sa-en/restaurant/riyadh/riyadh/68281 — لينكات جاهز هتشتغل قريب</div>

  <details class="settings" id="settings">
    <summary>&#9881; إعدادات صور Google Drive (بتتعمل مرة واحدة بس — كل واحد على درايفه)</summary>
    <div class="inner">
      <p>عشان الداشبورد يرفع صور المنيو على <b>الدرايف بتاعك أنت</b> ويرجّعلك لينك Google Drive لكل صورة، محتاج تعمل الإعداد ده مرة واحدة بس (٥ دقايق) — بيتم تخزينه في متصفحك انت:</p>
      <ol>
        <li>افتح <b>script.google.com</b> بحساب Google بتاعك واعمل <b>مشروع جديد</b></li>
        <li>امسح أي كود موجود، والصق الكود اللي تحت (زرار "انسخ الكود")</li>
        <li>من فوق اضغط <b>Deploy &gt; New deployment</b> واختار النوع <b>Web app</b></li>
        <li>الإعدادات: <b>Execute as: Me</b> + <b>Who has access: Anyone with a Google account</b></li>
        <li>اضغط Deploy وسيبك من التحذير: اختار حسابك &gt; Advanced &gt; Go to ... (unsafe) &gt; Allow</li>
        <li>هيظهرلك رابط بيخلص بـ <b>/exec</b> — انسخه وحطه تحت واضغط "حفظ"</li>
      </ol>
      <pre id="gasCode">// ==== Menu Dashboard Uploader ====
// Deploy as: Web app | Execute as: Me | Access: Anyone with a Google account

function doGet() {
  return HtmlService.createHtmlOutput('Uploader is running - &#1578;&#1605;&#1575;&#1605;');
}

function doPost(e) {
  var out = { ok: true, links: {}, errors: [] };
  try {
    var data = JSON.parse(e.parameter.payload);
    var folderName = data.folder || 'Menu Dashboard Images';
    var it = DriveApp.getFoldersByName(folderName);
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
    (data.images || []).forEach(function (img) {
      try {
        var name = img.name.replace(/[\\\/:*?"&lt;&gt;|]+/g, '_');
        var ex = folder.getFilesByName(name);
        var file;
        if (ex.hasNext()) {
          file = ex.next();
        } else {
          var blob = UrlFetchApp.fetch(img.url).getBlob().setName(name);
          file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        }
        out.links[img.id] = 'https://drive.google.com/uc?export=view&amp;id=' + file.getId();
      } catch (err) {
        out.errors.push(img.id + ': ' + err);
      }
    });
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  var t = JSON.stringify(out);
  var html = '&lt;!DOCTYPE html&gt;&lt;html&gt;&lt;body style="font-family:Segoe UI,Tahoma,sans-serif;direction:rtl;text-align:center;padding:40px"&gt;&lt;h3&gt;&#1578;&#1605; &#1606;&#1602;&#1604; &#1575;&#1604;&#1589;&#1608;&#1585;&lt;/h3&gt;&lt;p&gt;&#1578;&#1602;&#1583;&#1585; &#1578;&#1602;&#1601;&#1604; &#1575;&#1604;&#1589;&#1601;&#1581;&#1577; &#1583;&#1610;&lt;/p&gt;&lt;script&gt;window.opener.postMessage(' + t + ", '*');window.close();&lt;\/script&gt;&lt;/body&gt;&lt;/html&gt;";
  return HtmlService.createHtmlOutput(html).setTitle('Menu Dashboard Uploader');
}</pre>
      <div class="settings-row">
        <button class="btn-sm plain" onclick="copyCode()">انسخ الكود</button>
        <button class="btn-sm plain" onclick="window.open($('gas-url').value || 'https://script.google.com','_blank')">افتح script.google.com</button>
      </div>
      <div class="settings-row">
        <input id="gas-url" type="text" spellcheck="false" placeholder="https://script.google.com/macros/s/......./exec">
        <button class="btn-sm primary" onclick="saveSettings()">حفظ</button>
        <button class="btn-sm plain" onclick="testScript()">اختبر الرابط</button>
      </div>
      <div id="settings-status" style="margin-top:8px;font-size:13px"></div>
    </div>
  </details>

  <div id="status"></div>
  <div id="results"></div>
</main>
<script>
  const $ = (id) => document.getElementById(id);
  let data = null, link = null, settingsUrl = '';
  let driveLinks = null; // {flatIndex: driveUrl} after upload

  function esc(s) {
    return String(s).replace(/[&<>"']/g, c =>
      ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }
  function safeName(s) { return String(s).replace(/[\\\/:*?"<>|]+/g, '_').slice(0, 80); }
  function pad(n) { return String(n).padStart(3, '0'); }
  function setStatus(cls, txt) { const st = $('status'); st.className = cls; st.textContent = txt; }

  // ---------- settings (localStorage — per teammate browser) ----------
  settingsUrl = localStorage.getItem('gas_url') || '';
  if (settingsUrl) $('gas-url').value = settingsUrl;
  function saveSettings() {
    settingsUrl = $('gas-url').value.trim();
    if (settingsUrl && !settingsUrl.startsWith('https://script.google.com/')) {
      $('settings-status').innerHTML = '<span class="error" style="display:inline-block;padding:6px 12px">الرابط لازم يبدأ بـ https://script.google.com/</span>';
      return;
    }
    localStorage.setItem('gas_url', settingsUrl);
    $('settings-status').innerHTML = '<span class="ok" style="display:inline-block;padding:6px 12px">تم الحفظ &#10003; (في المتصفح بتاعك)</span>';
  }
  function testScript() {
    const url = $('gas-url').value.trim();
    if (!url.startsWith('https://script.google.com/')) {
      $('settings-status').innerHTML = '<span class="error" style="display:inline-block;padding:6px 12px">الرابط لازم يبدأ بـ https://script.google.com/</span>';
      return;
    }
    window.open(url, '_blank');
    $('settings-status').innerHTML = '<span style="color:var(--muted)">لو فتحت صفحة مكتوب فيها "Uploader is running" يبقى تمام. لو طلب تسجيل دخول: عادي سجل وعدّل إعدادات النشر لـ "Anyone with a Google account".</span>';
  }
  function copyCode() {
    navigator.clipboard.writeText($('gasCode').textContent).then(() => {
      $('settings-status').innerHTML = '<span class="ok" style="display:inline-block;padding:6px 12px">الكود اتنسخ — الصقه في script.google.com</span>';
    });
  }

  // ---------- extract ----------
  async function extract() {
    const url = $('url').value.trim();
    setStatus('loading', 'بيسحب المنيو، استنى ثانية...');
    $('results').innerHTML = '';
    $('go').disabled = true;
    try {
      const res = await fetch('/api/extract', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({url})
      });
      if (res.status === 401) { location.href = '/login'; return; }
      const j = await res.json();
      if (j.error) throw new Error(j.error);
      data = j; link = url; driveLinks = null;
      let i = 0;
      data.categories.forEach(g => g.items.forEach(it => { it._i = i++; }));
      setStatus('', '');
      render(data, '');
    } catch (e) {
      setStatus('error', 'مشكلة: ' + e.message);
    } finally {
      $('go').disabled = false;
    }
  }

  // ---------- render ----------
  function render(menu, q) {
    const r = menu.restaurant;
    const flat = [];
    menu.categories.forEach(g => g.items.forEach(it => flat.push(it)));
    const withImg = flat.filter(it => it.image).length;
    const prices = flat.map(it => it.price).filter(p => typeof p === 'number');
    const avg = prices.length ? (prices.reduce((a, b) => a + b, 0) / prices.length) : 0;
    const maxP = prices.length ? Math.max(...prices) : 0;
    const open = (r.status || '').toUpperCase() === 'OPEN';
    const drvCount = driveLinks ? Object.keys(driveLinks).length : 0;

    let html = `
      <div class="rest-card">
        <h2>${esc(r.name)}
          <span class="badge ${open ? '' : 'closed'}">${open ? 'مفتوح' : 'مغلق حالياً'}</span>
        </h2>
        <div class="meta">${esc((r.cuisines || []).join(' • '))}${r.rating ? ' — تقييم ' + esc(r.rating) + ' (' + esc(r.rating_count) + ')' : ''}</div>
      </div>
      <div class="stats">
        <div class="stat"><div class="v">${menu.categories.length}</div><div class="l">قسم</div></div>
        <div class="stat"><div class="v">${menu.total_items}</div><div class="l">صنف</div></div>
        <div class="stat"><div class="v">${withImg}</div><div class="l">صنف له صورة</div></div>
        <div class="stat"><div class="v">${avg.toFixed(1)}</div><div class="l">متوسط السعر (SR)</div></div>
        <div class="stat"><div class="v">${maxP.toFixed(1)}</div><div class="l">أغلى صنف (SR)</div></div>
      </div>
      <div class="toolbar">
        <input id="search" type="text" placeholder="دوّر على صنف... (مثلاً: شاورما)" value="${esc(q)}">
        <button class="btn green" id="export-plain" onclick="location.href='/api/export?url=' + encodeURIComponent(link)">تحميل Excel (لينكات HungerStation)</button>
        <button class="btn blue" id="export-drive" ${settingsUrl ? '' : 'disabled title="ظبط إعدادات Google Drive الأول"'}
          onclick="exportDrive()">${drvCount ? `تحميل Excel (${drvCount} صورة على Drive)` : 'تحميل Excel (صور Google Drive)'}</button>
      </div>`;

    if (!settingsUrl) html += `<div class="hint">&#9888; زرار صور Google Drive محتاج تظبط السكريبت مرة واحدة من <a href="#" onclick="$('settings').open=true;return false;">الإعدادات فوق</a>. لينكات HungerStation المباشرة شغالة حالاً.</div>`;

    const ql = q.trim().toLowerCase();
    let shown = 0;
    menu.categories.forEach(g => {
      const items = g.items.filter(it => !ql ||
        it.name.toLowerCase().includes(ql) ||
        (it.description || '').toLowerCase().includes(ql) ||
        g.name.toLowerCase().includes(ql));
      if (!items.length) return;
      shown += items.length;
      html += `<div class="cat"><h3>${esc(g.name)} <small>(${items.length})</small></h3><table>
        <tr><th>صورة</th><th>الصنف</th><th>الوصف</th><th>السعر (SR)</th><th>السعرات</th></tr>`;
      items.forEach(it => {
        const drv = driveLinks ? driveLinks[String(it._i)] : null;
        const imgCell = it.image
          ? `<a href="${esc(drv || it.image)}" target="_blank"><img class="thumb" src="${esc(it.image)}" loading="lazy" alt=""></a>${drv ? '<br><span class="drv-tag">Drive</span>' : ''}`
          : '—';
        html += `<tr>
          <td>${imgCell}</td>
          <td><b>${esc(it.name)}</b></td>
          <td>${esc(it.description || '—')}</td>
          <td class="price">${it.price != null ? esc(it.price.toFixed ? it.price.toFixed(2) : it.price) : '—'}</td>
          <td>${it.calories != null && it.calories !== '' ? esc(it.calories) : '—'}</td>
        </tr>`;
      });
      html += `</table></div>`;
    });

    if (!shown) html += `<div class="empty">مفيش نتايج مطابقة للبحث</div>`;
    $('results').innerHTML = html;

    const s = $('search');
    s.addEventListener('input', () => render(menu, s.value));
    if (q) { const c = s.value.length; s.focus(); s.setSelectionRange(c, c); }
  }

  // ---------- Google Drive upload (from the user's own browser session) ----------
  function postBatch(batch, folder) {
    return new Promise((resolve) => {
      const t = setTimeout(() => resolve({ok: false, error: 'السكريبت اتأخر أكتر من 8 دقايق — جرب تاني'}), 480000);
      const listener = (ev) => {
        window.removeEventListener('message', listener);
        clearTimeout(t);
        const d = ev.data;
        resolve(d && typeof d === 'object' ? d : {ok: false, error: 'رد غير معروف من السكريبت'});
      };
      window.addEventListener('message', listener);
      const w = window.open('', 'hsDriveWin', 'width=520,height=440');
      if (!w) { clearTimeout(t); window.removeEventListener('message', listener); resolve({ok: false, error: 'المتصفح حجب النافذة المنبثقة — اسمح بالنوافذ المنبثقة للصفحة دي و جرب تاني'}); return; }
      const f = document.createElement('form');
      f.action = settingsUrl; f.method = 'POST'; f.target = 'hsDriveWin';
      const inp = document.createElement('input'); inp.type = 'hidden'; inp.name = 'payload';
      inp.value = JSON.stringify({folder: folder, images: batch});
      f.appendChild(inp); document.body.appendChild(f); f.submit();
      setTimeout(() => f.remove(), 8000);
    });
  }

  async function exportDrive() {
    if (!settingsUrl) { $('settings').open = true; return; }
    if (!data) return;
    const flat = [];
    data.categories.forEach(g => g.items.forEach(it => flat.push(it)));
    if (!driveLinks) {
      const imgs = flat.filter(it => it.image)
        .map(it => ({id: String(it._i), url: it.image, name: pad(it._i) + ' - ' + safeName(it.name) + '.jpg'}));
      if (!imgs.length) { setStatus('error', 'المنيو ده مفيهوش صور أصلاً'); return; }
      const folder = 'HungerStation - ' + (safeName(data.restaurant.name) || 'Images');
      const CHUNK = 60; let done = 0;
      driveLinks = {};
      $('export-drive').disabled = true;
      for (let s = 0; s < imgs.length; s += CHUNK) {
        const batch = imgs.slice(s, s + CHUNK);
        setStatus('loading', 'بيرفع الصور على Google Drive... ' + done + ' / ' + imgs.length + ' — سيب النافذة الصغيرة تفتح وتقفل لوحدها');
        const res = await postBatch(batch, folder);
        if (!res.ok) {
          driveLinks = null;
          setStatus('error', 'حصل خطأ في الرفع: ' + (res.error || 'غير معروف') + ' — اتأكد من رابط السكريبت ومن إعدادات النشر (Anyone with a Google account)');
          return;
        }
        Object.assign(driveLinks, res.links || {});
        done += batch.length;
      }
      const errs = flat.filter(it => it.image && !driveLinks[String(it._i)]).length;
      setStatus('ok', 'تم رفع ' + Object.keys(driveLinks).length + ' صورة على Google Drive' + (errs ? ' — (' + errs + ' صورة فشلت، هيتحط لينك HungerStation مكانها)' : '') + ' — جاري تحميل ملف Excel...');
    } else {
      setStatus('loading', 'بيجهز ملف Excel...');
    }
    const res = await fetch('/api/export_links', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({menu: data, links: driveLinks})
    });
    if (res.status === 401) { location.href = '/login'; return; }
    if (!res.ok) { const t = await res.text(); setStatus('error', 'مشكلة في توليد الملف: ' + t); return; }
    const blob = await res.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = safeName(data.restaurant.name || 'menu') + ' - Menu (Google Drive).xlsx';
    a.click(); URL.revokeObjectURL(a.href);
    setStatus('ok', 'اتحمّل ملف Excel بلينكات Google Drive &#10003;');
    render(data, ($('search') || {}).value || '');
  }

  $('go').addEventListener('click', extract);
  $('url').addEventListener('keydown', e => { if (e.key === 'Enter') extract(); });
</script>
</body>
</html>
"""


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------

@app.get("/")
def index():
    return Response(PAGE, mimetype="text/html; charset=utf-8")


@app.post("/api/extract")
def api_extract():
    payload = flask_request.get_json(silent=True) or {}
    url = payload.get("url", "")
    try:
        menu = fetch_menu(url)
        return jsonify(menu)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"حصل خطأ أثناء السحب: {e}"}), 502


@app.get("/api/export")
def api_export():
    """Fast export: Image Link column = original HungerStation CDN URLs."""
    url = flask_request.args.get("url", "")
    try:
        menu = fetch_menu(url)
    except ValueError as e:
        return Response(str(e), status=400, mimetype="text/plain; charset=utf-8")
    except Exception as e:  # noqa: BLE001
        return Response(f"خطأ أثناء السحب: {e}", status=502, mimetype="text/plain; charset=utf-8")
    for g in menu["categories"]:
        for it in g["items"]:
            it["image_link"] = it.get("image") or ""
    wb = build_workbook(menu)
    buf = io.BytesIO()
    wb.save(buf)
    fname = re.sub(r'[\\/:*?"<>|]+', "_", (menu["restaurant"]["name"] or "menu")).strip() or "menu"
    fname = f"{fname} - Menu.xlsx"
    resp = Response(
        buf.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(fname)}"
    return resp


@app.post("/api/export_links")
def api_export_links():
    """Build Excel from the menu the browser already extracted + Drive links."""
    payload = flask_request.get_json(silent=True) or {}
    menu = payload.get("menu") or {}
    links = payload.get("links") or {}
    cats = menu.get("categories") or []
    if not isinstance(cats, list) or not cats:
        return jsonify({"error": "بيانات المنيو ناقصة — اسحب المنيو تاني"}), 400
    total = 0
    for g in cats:
        for it in g.get("items", []):
            it["image_link"] = links.get(str(total)) or (it.get("image") or "")
            total += 1
    menu["total_items"] = menu.get("total_items") or total
    wb = build_workbook(menu)
    buf = io.BytesIO()
    wb.save(buf)
    name = ((menu.get("restaurant") or {}).get("name") or "menu").strip() or "menu"
    fname = re.sub(r'[\\/:*?"<>|]+', "_", name) + " - Menu (Google Drive).xlsx"
    resp = Response(
        buf.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(fname)}"
    return resp


# --------------------------------------------------------------------------
# Entrypoint
# --------------------------------------------------------------------------

def cli_test(url: str, export: str | None = None) -> None:
    menu = fetch_menu(url)
    rest = menu["restaurant"]
    print("Restaurant :", rest["name"])
    print("Cuisines    :", ", ".join(rest["cuisines"]))
    print("Rating      :", rest["rating"], f"({rest['rating_count']} ratings)")
    print("Status      :", rest["status"])
    print("Categories  :", len(menu["categories"]))
    print("Total items :", menu["total_items"])
    with_img = sum(1 for g in menu["categories"] for it in g["items"] if it.get("image"))
    print("Items w/img :", with_img)
    for g in menu["categories"]:
        print(f"   - {g['name']}: {len(g['items'])} items")
    sample = menu["categories"][0]["items"][0]
    print("Sample item :", sample["name"], "| price:", sample["price"], "| cal:", sample["calories"])
    if export:
        for g in menu["categories"]:
            for it in g["items"]:
                it["image_link"] = it.get("image") or ""
        wb = build_workbook(menu)
        wb.save(export)
        print("Excel saved :", export)


def main() -> None:
    args = sys.argv[1:]
    if args and args[0] == "test":
        if len(args) < 2:
            print("usage: uv run app.py test <url> [--export file.xlsx]")
            sys.exit(1)
        export = args[args.index("--export") + 1] if "--export" in args else None
        cli_test(args[1], export)
        return

    print(f"Dashboard running at http://{HOST_BIND}:{PORT} ... (Ctrl+C to stop)")
    if HOST_BIND in ("127.0.0.1", "localhost"):
        threading.Timer(1.5, lambda: webbrowser.open(f"http://{HOST_BIND}:{PORT}")).start()
    app.run(host=HOST_BIND, port=PORT, debug=False)


if __name__ == "__main__":
    main()
