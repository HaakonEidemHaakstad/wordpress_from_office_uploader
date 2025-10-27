import os
import sys
import json
import base64
import re
import traceback
import mimetypes
from pathlib import Path
from urllib.parse import urljoin

import chardet
import requests
from bs4 import BeautifulSoup

# --- GUI (PySide6) ---
from PySide6.QtCore import Qt, Signal, QObject, QThread, QFileInfo, QMetaObject
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QComboBox,
    QCheckBox,
    QTextEdit,
    QProgressBar,
    QFileDialog,
    QFrame,
    QMessageBox,
    QFileIconProvider,
    QToolButton,
    QStyle,
    QSizePolicy,
)

# --- Windows crypto (DPAPI) ---
try:
    import win32crypt  # from pywin32

    HAS_DPAPI = True
except Exception:
    HAS_DPAPI = False

# --- COM automation ---
try:
    import win32com.client as win32

    HAS_COM = True
except Exception:
    HAS_COM = False


APP_TITLE = "Word/Excel → WordPress Page Updater"

ALLOWED_EXTS = {".doc", ".docx", ".xls", ".xlsx", ".htm", ".html"}

HERE = Path(getattr(sys, "_MEIPASS", Path.cwd())).resolve()  # PyInstaller-safe
DEFAULTS_PATH = (Path(sys.argv[0]).resolve().parent / "defaults.txt").resolve()


def _pt_to_px(css_text: str, ratio: float = 1.3333) -> str:
    """
    Convert all 'pt' units in a CSS string to 'px' (approx. 1pt = 1.3333px).
    """

    def repl(match):
        val = float(match.group(1))
        px = round(val * ratio, 2)
        return f"{px}px"

    return re.sub(r"([\d.]+)\s*pt", repl, css_text)


# ------------------ Helpers: DPAPI encrypt/decrypt ------------------


def protect_cleartext(text: str) -> str:
    """
    Protects text using Windows DPAPI; returns base64 of protected bytes.
    Falls back to plain base64 (obfuscation) if DPAPI unavailable.
    """
    if text is None:
        return ""
    data = text.encode("utf-8")
    if HAS_DPAPI:
        blob = win32crypt.CryptProtectData(data, None, None, None, None, 0)
        return base64.b64encode(blob).decode("ascii")
    # fallback (not real encryption)
    return "B64:" + base64.b64encode(data).decode("ascii")


def unprotect_to_cleartext(protected: str) -> str:
    if not protected:
        return ""
    try:
        if HAS_DPAPI:
            raw = base64.b64decode(protected)
            out = win32crypt.CryptUnprotectData(raw, None, None, None, 0)[1]
            return out.decode("utf-8")
        # fallback
        if protected.startswith("B64:"):
            return base64.b64decode(protected[4:]).decode("utf-8")
        return protected
    except Exception:
        return ""


def load_defaults():
    if not DEFAULTS_PATH.exists():
        return {}
    try:
        data = json.loads(DEFAULTS_PATH.read_text(encoding="utf-8"))
        # decrypt sensitive fields
        for k in ("application_password",):
            if k in data:
                data[k] = unprotect_to_cleartext(data[k])
        return data
    except Exception:
        return {}


def save_defaults(url, username, application_password):
    payload = {
        "url": url.strip(),
        "username": username.strip(),
        "application_password": protect_cleartext(application_password.strip()),
    }
    DEFAULTS_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")


# ------------------ Status bus ------------------


class Bus(QObject):
    log = Signal(str)
    step = Signal(int, int)  # current, total
    pages_ready = Signal(list)  # list of (id, title)
    done = Signal(bool, str)  # success, message


bus = Bus()


# ------------------ WordPress client ------------------


class WPClient:
    def __init__(self, base_url, username, app_password):
        self.base_url = base_url.rstrip("/") + "/"
        self.username = username
        self.app_password = app_password
        self.session = requests.Session()
        self.session.auth = (username, app_password)
        self.session.headers.update({"Accept": "application/json"})
        self.temp_backup_path = Path.cwd() / "temp_backup.html"
        self.last_page_id = None


    def _url(self, path):
        if path.startswith("http"):
            return path
        return urljoin(self.base_url, path.lstrip("/"))

    def list_pages(self, per_page=100):
        pages = []
        page = 1
        while True:
            r = self.session.get(
                self._url(f"/wp-json/wp/v2/pages"),
                params={"per_page": per_page, "page": page, "_fields": "id,title"},
            )
            r.raise_for_status()
            batch = r.json()
            if not batch:
                break
            for item in batch:
                title = (item.get("title") or {}).get("rendered", "")
                pages.append((item["id"], title))
            if len(batch) < per_page:
                break
            page += 1
        # sort by title asc, keep ID
        pages.sort(key=lambda t: (t[1] or "").lower())
        return pages

    def update_page_content(self, page_id, html_fragment):
        payload = {"content": html_fragment}
        r = self.session.post(
            self._url(f"/wp-json/wp/v2/pages/{page_id}"), json=payload
        )
        r.raise_for_status()
        return r.json()


# ------------------ Office converters ------------------


def _append_inline_style(el, css: str):
    st = el.get("style") or ""
    if st and not st.strip().endswith(";"):
        st = st.strip() + ";"
    el["style"] = (st + css).strip(";")


def clamp_hairline_borders(container):
    for el in container.find_all(True):
        st = el.get("style") or ""
        st2 = re.sub(
            r"\bborder([^:]*):\s*0\.5pt\s+solid\s+([^;]+);",
            r"border\1: 1px solid \2;",
            st,
            flags=re.I,
        )
        if st2 != st:
            el["style"] = st2


FONT_MAP = {
    "calibri": "Calibri, 'Segoe UI', Arial, Helvetica, sans-serif",
    "cambria": "Cambria, 'Times New Roman', Times, serif",
    "times new roman": "'Times New Roman', Times, serif",
    "arial": "Arial, Helvetica, sans-serif",
}


def map_font_families(container):
    for el in container.find_all(True):
        st = el.get("style") or ""
        m = re.search(r"font-family\s*:\s*([^;]+);", st, flags=re.I)
        if not m:
            continue
        fam_raw = m.group(1)
        fam_norm = fam_raw.strip().strip("'\"").lower()
        for k, repl in FONT_MAP.items():
            if k in fam_norm:
                st = re.sub(
                    r"font-family\s*:\s*[^;]+;", f"font-family: {repl};", st, flags=re.I
                )
                el["style"] = st
                break


def convert_word_to_html(src_path: Path, workdir: Path) -> Path:
    """
    Use Word COM automation to save as 'Filtered HTML' (wdFormatFilteredHTML = 10).
    Returns the output .htm path.
    """
    if not HAS_COM:
        raise RuntimeError("pywin32 / win32com is not available.")
    import pythoncom
    pythoncom.CoInitialize()
    try:
        bus.log.emit("Launching Word for conversion…")
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(str(src_path))

            # Hints to keep fidelity sensible
            try:
                doc.WebOptions.Encoding = 65001  # UTF-8
                doc.WebOptions.RelyOnCSS = True
                doc.WebOptions.OptimizeForBrowser = True
                doc.WebOptions.BrowserLevel = 4
                doc.WebOptions.AllowPNG = True
            except Exception:
                pass

            out_path = workdir / (src_path.stem + ".htm")
            # 10 = wdFormatFilteredHTML
            doc.SaveAs2(str(out_path), FileFormat=10)
            doc.Close(False)
        finally:
            word.Quit()
    finally:
        pythoncom.CoUninitialize()
    return out_path


def convert_excel_to_html(src_path: Path, workdir: Path) -> Path:
    """
    Use Excel COM automation to publish/save workbook as HTML.
    Returns the output .htm path.
    """
    if not HAS_COM:
        raise RuntimeError("pywin32 / win32com is not available.")
    import pythoncom
    pythoncom.CoInitialize()
    try:
        bus.log.emit("Launching Excel for conversion…")
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(str(src_path))
            out_path = workdir / (src_path.stem + ".htm")
            try:
                # Publish active sheet first (cleaner HTML)
                bus.log.emit("Excel: trying PublishObjects (active sheet)…")
                wb.PublishObjects.Add(
                    SourceType=1,                # xlSourceSheet
                    Filename=str(out_path),
                    Sheet=wb.ActiveSheet.Name,
                    HtmlType=0,                  # standard
                ).Publish(True)
            except Exception:
                bus.log.emit("PublishObjects failed; falling back to SaveAs (HTML)…")
                wb.SaveAs(str(out_path), FileFormat=44)  # 44 = xlHtml
            finally:
                wb.Close(False)
        finally:
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()
    return out_path



# ------------------ HTML cleaning & inlining ------------------

CSS_STRIP_RULES = [
    r"mso-[^:]+:[^;]+;?",
    r"-ms-[^:]+:[^;]+;?",
]

MSO_CLASS_RE = re.compile(r"\bmso\w*\b", flags=re.I)


def read_file_text(path: Path) -> str:
    """
    Reads text file with automatic encoding detection.
    Tries UTF-8 first; falls back to detected encoding (latin-1, cp1252, etc.)
    so Scandinavian/European characters survive.
    """
    try:
        # Try UTF-8 first (most HTML exports actually are)
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        raw = path.read_bytes()
        guess = chardet.detect(raw)
        enc = guess.get("encoding") or "latin-1"
        try:
            return raw.decode(enc)
        except Exception:
            # Final fallback
            return raw.decode("latin-1", errors="replace")


def data_uri_for(path: Path) -> str:
    mime, _ = mimetypes.guess_type(path.name)
    if not mime:
        mime = "application/octet-stream"
    b = path.read_bytes()
    return f"data:{mime};base64,{base64.b64encode(b).decode('ascii')}"


def ensure_default_line_height(container_soup, default_value: str = "1"):
    """
    Ensure a default line-height on common text elements if none is present.
    Does NOT overwrite existing line-height applied via CSS or inline styles.
    """
    selectors = ["p", "li", "div", "span", "h1", "h2", "h3", "h4", "h5", "h6"]
    for sel in selectors:
        for el in container_soup.select(sel):
            st = el.get("style") or ""
            if "line-height" in st.lower():
                continue  # respect existing line-height
            # add inline line-height
            if st and not st.strip().endswith(";"):
                st = st.strip() + ";"
            el["style"] = (st + f"line-height: {default_value};").strip(";")


def inline_assets_and_clean(html_path: Path) -> str:
    """
    Excel/Word -> WP cleaner that preserves Excel formatting:
      - Detect Excel frameset (mal_*.htm) and load the real sheet (sheet001.htm)
      - Keep <style>, inline external CSS from *_files/*.css (rewrite url(...) to data URIs)
      - Inline <img> sources
      - Light whitespace normalization in <p>/<li>
      - Wrap in flush-left, fixed-width container
    """
    # --- read helper should already detect encoding (windows-1252 common for Excel) ---
    html_text = read_file_text(html_path)

    # Prefer lxml, fall back gracefully
    try:
        soup = BeautifulSoup(html_text, "lxml")
    except Exception:
        soup = BeautifulSoup(html_text, "html.parser")

    root_dir = html_path.parent
    assets_dir = root_dir / (html_path.stem + "_files")  # default guess

    # ---------- If this is an Excel frameset, open the real sheet (e.g., *_files/sheet001.htm) ----------
    content_soup = None
    frameset = soup.find("frameset")
    if frameset:
        # Prefer first <frame src="...sheetXXX.htm">
        frame = frameset.find("frame", attrs={"src": True})
        candidate = None
        if frame:
            candidate = frame.get("src")
        # If not found, try Excel XML hint: <x:WorksheetSource HRef="*_files/sheet001.htm"/>
        if not candidate:
            try:
                xml_text = str(soup)
                m = re.search(r'HRef="([^"]+sheet\d+\.htm)"', xml_text, flags=re.I)
                if m:
                    candidate = m.group(1)
            except Exception:
                pass

        if candidate:
            sheet_path = (root_dir / candidate)
            if not sheet_path.exists():
                # also try under *_files by filename only
                sheet_path2 = (root_dir / (html_path.stem + "_files") / Path(candidate).name)
                if sheet_path2.exists():
                    sheet_path = sheet_path2
            if sheet_path.exists():
                # This becomes our main content; set assets_dir to that sheet's folder
                assets_dir = sheet_path.parent
                sheet_text = read_file_text(sheet_path)
                try:
                    content_soup = BeautifulSoup(sheet_text, "lxml")
                except Exception:
                    content_soup = BeautifulSoup(sheet_text, "html.parser")

    # If not a frameset or failed to resolve, use the original soup as content
    if content_soup is None:
        content_soup = soup

    # ---------- remove things WP won't need ----------
    for tag in content_soup.find_all(["script", "xml"]):
        tag.decompose()
    # Do NOT remove <style> — Excel depends on this

    # ---------- get just the body content ----------
    body = content_soup.body or content_soup
    # Collect any inline <style> found in the sheet (we’ll preserve them)
    inline_styles = [st for st in content_soup.find_all("style")]

    # Extract body children into a content div
    content = content_soup.new_tag("div")
    for child in list(body.children):
        content.append(child.extract())

    # ---------- path resolver ----------
    def resolve_local(src: str) -> Path | None:
        if not src:
            return None
        p = (assets_dir / src) if assets_dir.exists() else None
        # First try relative to assets_dir (Excel sheet lives there)
        if p and p.exists():
            return p
        # Then try relative to the original HTML path
        p2 = (html_path.parent / src)
        if p2.exists():
            return p2
        # Then try by filename inside assets_dir
        if assets_dir.exists():
            p3 = assets_dir / Path(src).name
            if p3.exists():
                return p3
        # And finally same dir as original
        p4 = html_path.parent / Path(src).name
        if p4.exists():
            return p4
        return None

    # ---------- inline <img> sources ----------
    for img in content.find_all("img"):
        src = img.get("src")
        local = resolve_local(src)
        if local:
            try:
                img["src"] = data_uri_for(local)
            except Exception:
                pass

    # ---------- inline external CSS: <link rel="stylesheet"> + any *.css in assets_dir ----------
    css_bundle = []

    # 1) <link rel="stylesheet" href="...">
    for ln in list(content_soup.find_all("link")):
        rel = (ln.get("rel") or [])
        rel = [r.lower() for r in rel] if isinstance(rel, list) else [str(rel).lower()]
        if "stylesheet" in rel or ln.get("rel", "") == "stylesheet":
            href = ln.get("href")
            local = resolve_local(href)
            if local and local.suffix.lower() == ".css":
                try:
                    css_text = local.read_text(encoding="utf-8", errors="ignore")
                    css_bundle.append((local, css_text))
                except Exception:
                    pass
            ln.decompose()  # remove link; we’re inlining

    # 2) Any *.css under assets_dir (Excel’s stylesheet.css lives here)
    if assets_dir.exists():
        for css_file in sorted(assets_dir.glob("*.css")):
            try:
                css_text = css_file.read_text(encoding="utf-8", errors="ignore")
                css_bundle.append((css_file, css_text))
            except Exception:
                pass

    # Rewrite url(...) inside bundled CSS to data: URIs
    def _rewrite_css_urls(css_text: str, base_dir: Path) -> str:
        def repl(m):
            raw = m.group(1).strip().strip('\'"')
            # Resolve relative to CSS file first
            local = (base_dir / raw)
            if local.exists():
                try:
                    return f"url({data_uri_for(local)})"
                except Exception:
                    return f"url({raw})"
            # Try by filename inside assets_dir
            if assets_dir.exists():
                by_name = assets_dir / Path(raw).name
                if by_name.exists():
                    try:
                        return f"url({data_uri_for(by_name)})"
                    except Exception:
                        return f"url({raw})"
            return f"url({raw})"
        return re.sub(r"url\(([^)]+)\)", repl, css_text, flags=re.I)

    bundled_css_text = ""
    for css_file, css_text in css_bundle:
        bundled_css_text += "\n" + _rewrite_css_urls(css_text, css_file.parent)

    # ---------- whitespace normalization (light) ----------
    # (keep table structure; just remove spacerun and collapse inside p/li)
    for sp in list(content.find_all("span")):
        st = (sp.get("style") or "").lower()
        if "mso-spacerun:yes" in st:
            txt = (sp.get_text() or " ").replace("\xa0", " ")
            txt = re.sub(r" {2,}", " ", txt)
            sp.replace_with(txt)

    for tag in list(content.find_all(["p", "li"])):
        for text_node in list(tag.find_all(string=True)):
            parent_name = getattr(text_node.parent, "name", "").lower()
            if parent_name in ("code", "pre", "style"):
                continue
            s = str(text_node).replace("\xa0", " ")
            s = re.sub(r"[ \t\r\n]+", " ", s)
            text_node.replace_with(s)
        for span in list(tag.find_all("span")):
            if not span.contents or (span.get_text(strip=True) == ""):
                span.decompose()

    # ---------- wrapper (flush-left) ----------
    wrapper = content_soup.new_tag("div", attrs={"class": "wp-office-fixed"})
    wrapper["style"] = (
        "margin:0 !important;"
        "padding:0;"
        "width:min(100%, 900px);"
        "background:transparent;"
        "overflow:visible;"
    )

    # inject bundled CSS first (sheet CSS + stylesheet.css)
    if bundled_css_text.strip():
        style_tag = content_soup.new_tag("style")
        style_tag.string = bundled_css_text
        wrapper.append(style_tag)

    # preserve any inline <style> tags from the sheet itself (insert before content)
    for st in inline_styles:
        # detach from old soup and attach into wrapper
        try:
            st.extract()
            wrapper.append(st)
        except Exception:
            pass

    wrapper.append(content)

    # ---------- cleanup wrapper noise ----------
    frag = str(wrapper)
    frag = re.sub(r"<!--\[if.*?endif\]-->", "", frag, flags=re.I | re.S)
    frag = re.sub(r"<\s*(html|head|body)[^>]*>", "", frag, flags=re.I)
    frag = re.sub(r"</\s*(html|head|body)\s*>", "", frag, flags=re.I)

    return frag



# ------------------ Worker thread ------------------


class UpdateWorker(QThread):
    def __init__(self, task, *args, **kwargs):
        super().__init__()
        self.task = task
        self.args = args
        self.kwargs = kwargs
        self.success = False
        self.message = ""

    def run(self):
        try:
            self.task(*self.args, **self.kwargs)
            self.success = True
            self.message = f"✓ Success."
        except Exception as e:
            self.success = False
            self.message = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"


# ------------------ GUI ------------------


class DropFrame(QFrame):
    fileDropped = Signal(str)

    def __init__(self):
        super().__init__()
        self.setFrameShape(QFrame.NoFrame)
        self.setAcceptDrops(True)
        self.setMinimumHeight(180)
        self.setCursor(Qt.PointingHandCursor)  # 👈 visual hint it’s clickable

        pal = self.palette()
        pal.setColor(QPalette.Window, pal.color(QPalette.Base))
        self.setPalette(pal)
        self.setAutoFillBackground(True)

        # Optional: subtle “clickable”/hover styling
        self.setStyleSheet(
            "DropFrame { border: 2px dashed #999; padding: 10px; border-radius: 8px; background-color: transparent; }"
        )

        # If you already added icons earlier, keep those; otherwise keep your original layout
        self.icon_provider = QFileIconProvider()

        lay = QVBoxLayout(self)
        lay.setSpacing(2)
        lay.setContentsMargins(8, 10, 8, 10)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(96, 96)
        self.icon_label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)

        self.label = QLabel(
            "Drag & drop or click to choose a file (Excel/Word/HTML) to upload."
        )
        self.label.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.label.setWordWrap(False)
        self.label.setMinimumWidth(320)

        # 🔹 Make it as small as its text allows vertically
        self.label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.label.setFixedHeight(
            self.label.sizeHint().height() + 2
        )  # +2 = tiny padding

        # Optional: small margin tweak
        self.label.setContentsMargins(0, 0, 0, 0)

        lay.addWidget(self.icon_label, alignment=Qt.AlignHCenter)
        lay.addWidget(self.label, alignment=Qt.AlignHCenter)

        self._set_default_preview()

    # --- existing helpers (unchanged) ---
    def _set_default_preview(self):
        generic_icon = QFileIconProvider().icon(QFileIconProvider.File)
        self.icon_label.setPixmap(generic_icon.pixmap(96, 96))
        self.label.setText(
            "Drag & drop or click to choose a file (Excel/Word/HTML) to upload."
        )

    def set_file_preview(self, path: str):
        info = QFileInfo(path)
        icon = QFileIconProvider().icon(info)
        pm = icon.pixmap(96, 96)
        self.icon_label.setPixmap(pm)
        self.label.setText(info.fileName())

    # --- NEW: click opens file dialog ---
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            filters = (
                "All Supported (*.doc *.docx *.xls *.xlsx *.htm *.html);;"
                "Word Documents (*.doc *.docx);;"
                "Excel Workbooks (*.xls *.xlsx);;"
                "HTML Files (*.htm *.html);;"
                "All Files (*.*)"
            )
            path, _ = QFileDialog.getOpenFileName(
                self, "Choose a file", str(Path.home()), filters
            )
            if path:
                self.set_file_preview(path)
                self.fileDropped.emit(path)
        # pass to base so selection/drag states still behave normally
        super().mousePressEvent(event)

    # --- DnD stays the same ---
    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls():
            for u in e.mimeData().urls():
                if Path(u.toLocalFile()).suffix.lower() in ALLOWED_EXTS:
                    e.acceptProposedAction()
                    return
        e.ignore()

    def dropEvent(self, e: QDropEvent):
        paths = [Path(u.toLocalFile()) for u in e.mimeData().urls()]
        for p in paths:
            if p.suffix.lower() in ALLOWED_EXTS:
                self.set_file_preview(str(p))  # keep preview in sync
                self.fileDropped.emit(str(p))
                break


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(560, 780)

        central = QWidget()
        self.setCentralWidget(central)
        v = QVBoxLayout(central)

        # --- Drop area ---
        self.drop = DropFrame()
        v.addWidget(self.drop)

        # spacing between sections
        v.addSpacing(20)

        # --- Login section header ---
        login_label = QLabel("<b>Login:</b>")
        login_label.setStyleSheet(
            "font-weight: bold; font-size: 13px; margin-top: 2px;"
        )
        v.addWidget(login_label)

        # --- Login inputs with help buttons (to the right) ---
        self.url, row_url = self._make_input_with_help(
            "URL (e.g., https://example.com)",
            "Site URL",
            "Your WordPress site base URL, e.g. https://example.com . "
            "Include https:// and no trailing /wp-admin.",
        )
        v.addLayout(row_url)

        self.username, row_user = self._make_input_with_help(
            "Username with administrator privileges",
            "Username",
            "WordPress user that owns an Application Password with permission to edit pages.",
        )
        v.addLayout(row_user)

        self.app_password, row_pass = self._make_input_with_help(
            "Application password",
            "Application Password",
            "Create this in WordPress: Profile → Application Passwords. "
            "Paste it here exactly as shown (with spaces).",
            password=True,
        )
        v.addLayout(row_pass)

        # Login row (button + remember)
        row_login = QHBoxLayout()
        self.btn_login = QPushButton("Log in")
        self.chk_remember = QCheckBox("Remember log in information.")
        row_login.addWidget(self.btn_login)
        row_login.addWidget(self.chk_remember)
        v.addLayout(row_login)

        # spacing between sections
        v.addSpacing(20)

        # --- Page chooser header ---
        page_label = QLabel("<b>Select Page to Update:</b>")
        page_label.setStyleSheet("font-weight: bold; font-size: 13px; margin-top: 2px;")
        v.addWidget(page_label)

        # Page chooser + help button
        self.page_combo = QComboBox()
        self.page_combo.setEditable(False)
        self.page_combo.setInsertPolicy(QComboBox.NoInsert)
        self.page_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.page_combo.setToolTip(
            "Choose the WordPress page whose content will be replaced with the cleaned HTML."
        )

        row_page = self._wrap_with_help(
            self.page_combo,
            "Select Page",
            "Choose the WordPress page to update. The program will replace that page’s content "
            "with the content in the selected Word/Excel/HTML file.",
        )

        v.addLayout(row_page)

        # spacing between sections
        v.addSpacing(20)

        # --- Update / confirm row ---
        update_label = QLabel("<b>Upload / Update:</b>")
        update_label.setStyleSheet(
            "font-weight: bold; font-size: 13px; margin-top: 2px;"
        )
        v.addWidget(update_label)

        row2 = QHBoxLayout()
        self.btn_update = QPushButton("Update")
        self.btn_update.setEnabled(False)
        self.chk_confirm = QCheckBox("I have selected the correct page.")
        row2.addWidget(self.btn_update)
        row2.addWidget(self.chk_confirm)
        v.addLayout(row2)

        self.btn_undo = QPushButton("Undo / Regret")
        self.btn_undo.setEnabled(False)
        self.btn_undo.clicked.connect(self.on_undo_clicked)
        # fix
        v.addWidget(self.btn_undo)


        # Progress (used only for Update task)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        v.addWidget(self.progress)

        # Log console (color-coded)
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setMinimumHeight(160)
        self.console.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #cccccc;
                font-family: Consolas, monospace;
                font-size: 12px;
                border: 1px solid #444;
            }
        """)
        v.addWidget(self.console)

        # State
        self.current_file: Path | None = None
        self.wp: WPClient | None = None

        # Wire up
        self.drop.fileDropped.connect(self.on_file_dropped)
        self.btn_login.clicked.connect(self.on_login_clicked)
        self.chk_confirm.stateChanged.connect(self.on_confirm_changed)
        self.btn_update.clicked.connect(self.on_update_clicked)


        bus.log.connect(self.log)
        bus.step.connect(self.on_step)
        bus.pages_ready.connect(self.populate_pages)
        bus.done.connect(self.on_done)

        # Load defaults if available
        self.load_defaults_to_fields()

        self.temp_backup_path = Path.cwd() / "temp_backup.html"
        self.last_page_id = None

    # ---------- small helpers ----------

    def _wrap_with_help(self, widget, help_title: str, help_text: str):
        """
        Wrap an input widget with a small '?' help button on the right.
        Returns the QHBoxLayout so you can add it to the main layout.
        """
        btn = QToolButton()
        btn.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxQuestion))
        btn.setToolTip(help_text)
        btn.setFixedSize(22, 22)
        btn.setStyleSheet(
            "QToolButton { border: none; padding: 0; } QToolButton:hover { color: #0078d7; }"
        )
        btn.clicked.connect(
            lambda: QMessageBox.information(self, help_title, help_text)
        )

        row = QHBoxLayout()
        row.addWidget(widget)
        row.addWidget(btn)
        return row

    def _make_input_with_help(
        self, placeholder: str, help_title: str, help_text: str, password: bool = False
    ):
        """
        Returns (line_edit, layout) where the layout contains the line edit + small help button.
        """
        edit = QLineEdit()
        edit.setPlaceholderText(placeholder)
        if password:
            edit.setEchoMode(QLineEdit.Password)

        btn = QToolButton()
        btn.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxQuestion))
        btn.setToolTip(help_text)
        btn.setFixedSize(22, 22)
        btn.setStyleSheet(
            "QToolButton { border: none; padding: 0; } QToolButton:hover { color: #0078d7; }"
        )
        btn.clicked.connect(
            lambda: QMessageBox.information(self, help_title, help_text)
        )

        row = QHBoxLayout()
        row.addWidget(edit)
        row.addWidget(btn)
        return edit, row

    def log(self, msg: str):
        """
        Append color-coded messages to the console QTextEdit.
        """
        lower = msg.lower()
        if any(k in lower for k in ("error", "failed", "exception", "traceback")):
            color = "#ff4d4d"  # red
        elif any(k in lower for k in ("success", "updated", "done", "✓")):
            color = "#00cc66"  # green
        elif any(
            k in lower
            for k in (
                "connecting",
                "converting",
                "cleaning",
                "updating",
                "loading",
                "launching",
            )
        ):
            color = "#3399ff"  # blue
        elif any(k in lower for k in ("warning", "note", "saved", "info")):
            color = "#ffaa00"  # amber
        else:
            color = "#cccccc"  # light gray default

        html_line = f'<span style="color:{color};">{msg}</span>'
        self.console.append(html_line)
        self.console.ensureCursorVisible()
        QApplication.processEvents()

    def on_step(self, cur, total):
        val = int(100 * cur / max(1, total))
        self.progress.setValue(val)
        QApplication.processEvents()

    def on_done(self, ok, message):
        # Ensure the bar completes visually at end of task
        self.progress.setValue(100)
        self.log(message)

    def on_confirm_changed(self, state):
        self.btn_update.setEnabled(
            bool(state) and self.page_combo.count() > 0 and bool(self.current_file)
        )

    def populate_pages(self, pages):
        self.page_combo.clear()
        for pid, title in pages:
            self.page_combo.addItem(title or f"(untitled {pid})", pid)
        self.log(f"Loaded {len(pages)} pages.")
        self.on_confirm_changed(self.chk_confirm.checkState())

    def on_file_dropped(self, path_str: str):
        p = Path(path_str)
        self.current_file = p
        self.log(f"Selected file: {p}")
        self.on_confirm_changed(self.chk_confirm.checkState())

    def load_defaults_to_fields(self):
        data = load_defaults()
        if data:
            self.url.setText(data.get("url", ""))
            self.username.setText(data.get("username", ""))
            self.app_password.setText(data.get("application_password", ""))
            self.chk_remember.setChecked(True)
            self.log("Loaded defaults.txt")

    # ---------- button handlers ----------

    def perform_upload(self, html_fragment: str):
        """
        Upload the given HTML fragment to the selected WordPress page.
        - Backup current page content to temp_backup.html before overwriting
        - POST (or PUT on 405) the new content
        - Enable Undo / Regret on success
        """
        # Basic guards
        if self.page_combo.count() == 0 or self.page_combo.currentIndex() < 0:
            QMessageBox.warning(self, "Upload", "Please select a page to update.")
            return
        site = self.url.text().strip().rstrip("/")
        user = self.username.text().strip()
        app_pass = self.app_password.text().strip()


        try:
            page_id = int(self.page_combo.currentData())
        except Exception:
            QMessageBox.warning(self, "Upload", "Could not read selected page ID.")
            return

        page_title = self.page_combo.currentText() or f"ID {page_id}"
        api_base = f"{site}/wp-json/wp/v2/pages"
        auth = (user, app_pass)

        # Progress + log (if present in your UI)
        if hasattr(self, "progress"): self.progress.setValue(60)
        if hasattr(self, "log"): self.log(f"Preparing to update: “{page_title}” (ID {page_id})")

        # -------- Backup current content --------
        if hasattr(self, "log"): self.log("Fetching current page content for backup...")
        try:
            get_url = f"{api_base}/{page_id}?context=edit"
            r = requests.get(get_url, auth=auth, timeout=20)
            if r.status_code in (401, 403):
                # fallback to rendered if raw not allowed
                get_url = f"{api_base}/{page_id}"
                r = requests.get(get_url, auth=auth, timeout=20)
            r.raise_for_status()
            data = r.json()
            current_html = (
                (data.get("content") or {}).get("raw")
                or (data.get("content") or {}).get("rendered")
                or ""
            )
            self.temp_backup_path.write_text(current_html, encoding="utf-8")
            self.last_page_id = page_id
            if hasattr(self, "btn_undo"): self.btn_undo.setEnabled(True)
            if hasattr(self, "log"): self.log(f"Backup saved: {self.temp_backup_path}")
        except Exception as e:
            if hasattr(self, "log"): self.log(f"Warning: backup failed ({e}). Proceeding with upload.")
            if hasattr(self, "btn_undo"): self.btn_undo.setEnabled(self.temp_backup_path.exists())

        if hasattr(self, "progress"): self.progress.setValue(83)
        if hasattr(self, "log"):
            size_kb = len(html_fragment.encode("utf-8")) / 1024.0
            self.log(f"Uploading new content ({size_kb:.1f} KB) to WordPress...")

        # -------- Upload new content --------
        payload = {"content": html_fragment}
        update_url = f"{api_base}/{page_id}"
        try:
            resp = requests.post(update_url, auth=auth, json=payload, timeout=30)
            if resp.status_code == 405:  # some sites require PUT
                resp = requests.put(update_url, auth=auth, json=payload, timeout=30)
            resp.raise_for_status()

            if hasattr(self, "progress"): self.progress.setValue(100)
            if hasattr(self, "log"): self.log(f"Upload complete: “{page_title}” updated.")
            if hasattr(self, "chk_confirm"):
                try: self.chk_confirm.setChecked(False)
                except Exception: pass
            if hasattr(self, "btn_undo") and self.temp_backup_path.exists():
                self.btn_undo.setEnabled(True)
        except Exception as e:
            if hasattr(self, "log"): self.log(f"ERROR: Upload failed: {e}")
            QMessageBox.critical(self, "Upload failed", str(e))
            if hasattr(self, "progress"): self.progress.setValue(0)
            return


    def on_login_clicked(self):
        url = self.url.text().strip()
        username = self.username.text().strip()
        apppw = self.app_password.text().strip()
        if not url or not username or not apppw:
            QMessageBox.warning(
                self, "Missing", "Please fill URL, username, and application password."
            )
            return

        def task():
            # No progress bar usage during login
            bus.log.emit("Connecting to WordPress…")
            self.wp = WPClient(url, username, apppw)
            pages = self.wp.list_pages()
            bus.pages_ready.emit(pages)
            if self.chk_remember.isChecked():
                save_defaults(url, username, apppw)
                bus.log.emit(f"Saved defaults to {DEFAULTS_PATH}")

        worker = UpdateWorker(task)
        worker.finished.connect(lambda: bus.done.emit(worker.success, worker.message))
        worker.start()

        # reset progress bar visually (idle during login)
        self.progress.setValue(0)

    def on_update_clicked(self):
        """
        Handle the Update button click:
        - Checks confirmation checkbox
        - Converts DOC/DOCX/XLS/XLSX to HTML (via COM) or uses HTML directly
        - Cleans/inlines with inline_assets_and_clean
        - Uploads via perform_upload (which also backs up for Undo)
        """
        if hasattr(self, "chk_confirm") and not self.chk_confirm.isChecked():
            QMessageBox.warning(
                self,
                "Confirm Update",
                "Please check the box confirming you've selected the correct page before uploading."
            )
            return

        if not getattr(self, "current_file", None):
            QMessageBox.warning(self, "No File", "Please drag and drop a file to upload.")
            return

        try:
            src = Path(self.current_file)
            if not src.exists():
                QMessageBox.warning(self, "No File", f"File not found: {src}")
                return

            # Progress/log helpers
            if hasattr(self, "progress"): self.progress.setValue(10)
            if hasattr(self, "log"): self.log(f"Input: {src.name}")

            # Prepare work dir for exports
            workdir = Path.cwd() / "_tmp_export"
            workdir.mkdir(exist_ok=True)

            # Decide conversion
            ext = src.suffix.lower()
            if ext in {".htm", ".html"}:
                html_path = src
                if hasattr(self, "log"): self.log("Using provided HTML file.")
                if hasattr(self, "progress"): self.progress.setValue(25)
            elif ext in {".doc", ".docx"}:
                if hasattr(self, "log"): self.log("Converting Word document to filtered HTML…")
                html_path = convert_word_to_html(src, workdir)
                if hasattr(self, "progress"): self.progress.setValue(30)
            elif ext in {".xls", ".xlsx"}:
                if hasattr(self, "log"): self.log("Converting Excel workbook to HTML…")
                html_path = convert_excel_to_html(src, workdir)
                if hasattr(self, "progress"): self.progress.setValue(30)
            else:
                raise RuntimeError(f"Unsupported file type: {ext}. Expected HTML, Word, or Excel.")

            # Clean & inline
            if hasattr(self, "log"): self.log("Inlining assets and cleaning for WordPress…")
            if hasattr(self, "progress"): self.progress.setValue(55)
            fragment = inline_assets_and_clean(html_path)
            if not isinstance(fragment, str) or not fragment.strip():
                raise RuntimeError("Cleaning produced an empty fragment.")

            # Upload (with backup inside)
            if hasattr(self, "log"): self.log("Uploading…")
            if hasattr(self, "progress"): self.progress.setValue(70)
            self.perform_upload(fragment)

            # Uncheck confirm after success (defensive)
            try:
                if hasattr(self, "chk_confirm"):
                    self.chk_confirm.setChecked(False)
            except Exception:
                pass

        except Exception as e:
            QMessageBox.critical(self, "Upload Error", f"Upload failed:\n{e}")
            if hasattr(self, "log"):
                self.log(f"Upload error: {e}")
            if hasattr(self, "progress"): self.progress.setValue(0)


    def on_undo_clicked(self):
        if not getattr(self, "temp_backup_path", None) or not self.temp_backup_path.exists():
            QMessageBox.warning(self, "Undo / Regret", "No backup found to restore.")
            return
        if not getattr(self, "last_page_id", None):
            QMessageBox.warning(self, "Undo / Regret", "No recorded page to restore.")
            return

        if QMessageBox.question(
            self, "Undo / Regret",
            "Do you want to restore the previous version of this page?",
            QMessageBox.Yes | QMessageBox.No
        ) != QMessageBox.Yes:
            return

        try:
            site = self.url.text().strip().rstrip("/")
            auth = (self.username.text().strip(), self.app_password.text().strip())

            url = f"{site}/wp-json/wp/v2/pages/{self.last_page_id}"
            html_to_restore = self.temp_backup_path.read_text(encoding="utf-8")
            payload = {"content": html_to_restore}

            r = requests.post(url, auth=auth, json=payload, timeout=30)
            if r.status_code == 405:
                r = requests.put(url, auth=auth, json=payload, timeout=30)
            r.raise_for_status()

            QMessageBox.information(self, "Undo / Regret", "Page restored to previous content.")
            if hasattr(self, "log"): self.log("Undo complete — previous version restored successfully.")
            if hasattr(self, "btn_undo"): self.btn_undo.setEnabled(False)
        except Exception as e:
            QMessageBox.critical(self, "Undo / Regret", f"Undo failed:\n{e}")
            if hasattr(self, "log"): self.log(f"Undo failed: {e}")


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
