"""
main_mac_ready.py
-----------------
Cross-platform helpers for your Word/Excel → HTML uploader app with:
- macOS Keychain support via `keyring` for storing the WordPress application password.
- LibreOffice (`soffice`) fallback for .docx/.xlsx conversion on non-Windows systems.
- Direct handling of .htm/.html inputs saved from Word/Excel (with asset-folder detection, inlining, and cleanup).
- User warning dialog if an .html is dropped without its companion assets folder.

Integration:
- Import or copy these functions into your existing app.
- Replace your old converters with `convert_word_to_html()` / `convert_excel_to_html()`.
- Route files via `process_file_for_upload(file_path, workdir, parent_window)`.
- Keep your upload-to-WordPress code as-is, using the returned single-file HTML path.

Dependencies:
    pip install PySide6 requests beautifulsoup4 chardet keyring

macOS prerequisites for .docx/.xlsx direct conversion:
    brew install --cask libreoffice

Windows keeps using COM (Word/Excel) and DPAPI automatically.
"""

from __future__ import annotations

import base64
import mimetypes
import platform
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, List

# ---------- Optional GUI (for warning dialog) ----------
try:
    from PySide6.QtWidgets import QMessageBox
    HAS_QT = True
except Exception:
    HAS_QT = False

# ---------- Secure storage: DPAPI (Windows) and Keychain (macOS/Linux) ----------
try:
    import win32crypt  # Windows DPAPI
    HAS_DPAPI = platform.system() == "Windows"
except Exception:
    HAS_DPAPI = False

try:
    import keyring  # Cross-platform keychain
    HAS_KEYRING = True
except Exception:
    HAS_KEYRING = False

APP_KEYRING_SERVICE = "WP-Office-Uploader"


def protect_cleartext(text: Optional[str]) -> str:
    """
    Protects an application password or token for storage.
    - On Windows: wraps with DPAPI (CryptProtectData) and returns "DPAPI:<b64>"
    - Where keyring is available: stores secret in OS keychain and returns "KEYRING:<key>"
    - Else: last-resort base64 obfuscation "B64:<b64>"
    """
    if text is None:
        return ""
    if HAS_DPAPI:
        blob = win32crypt.CryptProtectData(text.encode("utf-8"), None, None, None, None, 0)
        return "DPAPI:" + base64.b64encode(blob).decode("ascii")
    if HAS_KEYRING:
        key = "application_password"
        keyring.set_password(APP_KEYRING_SERVICE, key, text)
        return "KEYRING:" + key
    return "B64:" + base64.b64encode(text.encode("utf-8")).decode("ascii")


def unprotect_to_cleartext(protected: Optional[str]) -> str:
    """
    Reverse of protect_cleartext(). Returns the cleartext or "" if not available.
    """
    if not protected:
        return ""
    try:
        if protected.startswith("DPAPI:") and HAS_DPAPI:
            raw = base64.b64decode(protected[len("DPAPI:"):])
            out = win32crypt.CryptUnprotectData(raw, None, None, None, 0)[1]
            return out.decode("utf-8")
        if protected.startswith("KEYRING:") and HAS_KEYRING:
            key = protected.split(":", 1)[1]
            return keyring.get_password(APP_KEYRING_SERVICE, key) or ""
        if protected.startswith("B64:"):
            return base64.b64decode(protected[4:]).decode("utf-8")
    except Exception:
        pass
    return ""


# ---------- Conversion: Windows COM vs LibreOffice CLI ----------
try:
    import win32com.client as win32com  # type: ignore
    HAS_COM = platform.system() == "Windows"
except Exception:
    HAS_COM = False


def _find_soffice() -> Optional[str]:
    """
    Locate LibreOffice CLI (soffice). On macOS, try the app bundle default path.
    """
    p = shutil.which("soffice")
    if p:
        return p
    mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if Path(mac_path).exists():
        return mac_path
    return None


def _convert_with_soffice(src_path: Path, workdir: Path) -> Path:
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice 'soffice' not found. Install LibreOffice or add 'soffice' to PATH.")
    workdir.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, "--headless", "--convert-to", "html", "--outdir", str(workdir), str(src_path)]
    subprocess.run(cmd, check=True)
    out = workdir / (src_path.stem + ".html")
    if not out.exists():
        alt = workdir / (src_path.stem + ".htm")
        if alt.exists():
            out = alt
    if not out.exists():
        candidates = list(workdir.glob(f"{src_path.stem}*.htm*"))
        if not candidates:
            raise RuntimeError("LibreOffice conversion did not produce an HTML file.")
        out = candidates[0]
    return out


def _convert_word_with_com(src_path: Path, workdir: Path) -> Path:
    """
    Windows-only: use Word COM to save as filtered HTML.
    """
    if not HAS_COM:
        raise RuntimeError("COM is not available on this platform.")
    word = win32com.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(src_path))
        out_path = workdir / (src_path.stem + ".html")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        # 10 = wdFormatFilteredHTML (reduces MSO gunk)
        doc.SaveAs(str(out_path), FileFormat=10)
        doc.Close(False)
    finally:
        word.Quit()
    return out_path


def _convert_excel_with_com(src_path: Path, workdir: Path) -> Path:
    """
    Windows-only: use Excel COM to save as filtered HTML.
    """
    if not HAS_COM:
        raise RuntimeError("COM is not available on this platform.")
    excel = win32com.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(str(src_path))
        out_path = workdir / (src_path.stem + ".html")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        # 45 = xlHtml (web page)
        wb.SaveAs(str(out_path), FileFormat=45)
        wb.Close(False)
    finally:
        excel.Quit()
    return out_path


def convert_word_to_html(src_path: Path, workdir: Path) -> Path:
    """
    Cross-platform Word → HTML
    """
    if platform.system() == "Windows" and HAS_COM:
        return _convert_word_with_com(src_path, workdir)
    return _convert_with_soffice(src_path, workdir)


def convert_excel_to_html(src_path: Path, workdir: Path) -> Path:
    """
    Cross-platform Excel → HTML
    """
    if platform.system() == "Windows" and HAS_COM:
        return _convert_excel_with_com(src_path, workdir)
    return _convert_with_soffice(src_path, workdir)


# ---------- HTML handling for files manually saved from Word/Excel (.htm/.html) ----------
from bs4 import BeautifulSoup, Comment  # type: ignore

def _read_text_guess_encoding(p: Path) -> str:
    """Robustly read HTML/CSS with a best-effort encoding guess."""
    raw = p.read_bytes()
    enc = "utf-8"
    try:
        import chardet  # type: ignore
        guess = chardet.detect(raw) or {}
        if guess.get("encoding"):
            enc = guess["encoding"]
    except Exception:
        pass
    return raw.decode(enc, errors="ignore")


def _detect_companion_dirs(html_path: Path) -> List[Path]:
    """
    Find likely asset folders produced by Word/Excel/LibreOffice next to the HTML.
    Returns a list ordered by likelihood.
    """
    parent = html_path.parent
    stem = html_path.stem
    candidates = [
        parent / f"{stem}_files",
        parent / f"{stem}.files",
        parent / f"{stem}-files",
        parent / f"{stem}.html_files",
        parent / f"{stem}-Dateien",
        parent / f"{stem}.fld",
        parent / f"{stem}_Dateien",
    ]
    return [c for c in candidates if c.exists() and c.is_dir()]


def _read_binary_as_data_uri(file_path: Path) -> str:
    mime, _ = mimetypes.guess_type(str(file_path))
    if not mime:
        ext = file_path.suffix.lower()
        if ext in (".png", ".apng"):
            mime = "image/png"
        elif ext in (".jpg", ".jpeg", ".jfif"):
            mime = "image/jpeg"
        elif ext == ".gif":
            mime = "image/gif"
        elif ext == ".bmp":
            mime = "image/bmp"
        elif ext == ".svg":
            mime = "image/svg+xml"
        elif ext == ".webp":
            mime = "image/webp"
        elif ext == ".css":
            mime = "text/css"
        else:
            mime = "application/octet-stream"
    try:
        b = file_path.read_bytes()
    except Exception:
        return ""
    b64 = base64.b64encode(b).decode("ascii")
    if mime == "text/css":
        return f"data:text/css;base64,{b64}"
    return f"data:{mime};base64,{b64}"


def _resolve_local(src: str, html_dir: Path, companion_dirs: List[Path]) -> Optional[Path]:
    """
    Resolve a relative href/src against companion dirs or HTML directory.
    Ignore absolute URLs (http/https/data/mailto/#).
    """
    if not src or re.match(r"^(data:|https?:|mailto:|#)", src, re.I):
        return None
    src_clean = src.replace("\\", "/").lstrip("./")
    for d in companion_dirs:
        candidate = d / src_clean
        if candidate.exists():
            return candidate
    candidate = html_dir / src_clean
    if candidate.exists():
        return candidate
    name = Path(src_clean).name
    for d in companion_dirs:
        hit = list(d.rglob(name))
        if hit:
            return hit[0]
    return None


def _inline_assets_and_clean(html_text: str, html_path: Path) -> str:
    """
    Inline linked CSS and IMG resources; do light cleanup for WordPress/Astra.
    """
    soup = BeautifulSoup(html_text, "html.parser")
    html_dir = html_path.parent
    companions = _detect_companion_dirs(html_path)

    # Inline <link rel="stylesheet">
    for link in list(soup.find_all("link")):
        rels = [r.lower() for r in link.get("rel", [])]
        href = link.get("href")
        if "stylesheet" in rels and href:
            local = _resolve_local(href, html_dir, companions)
            if local:
                css_text = _read_text_guess_encoding(local)
                style_tag = soup.new_tag("style")
                css_text = re.sub(r"/\*.*?\*/", "", css_text, flags=re.S)  # strip comments
                style_tag.string = css_text
                link.replace_with(style_tag)
            else:
                link.decompose()

    # Inline <img src="...">
    for img in list(soup.find_all("img")):
        src = img.get("src")
        local = _resolve_local(src, html_dir, companions)
        if local:
            data_uri = _read_binary_as_data_uri(local)
            if data_uri:
                img["src"] = data_uri

    # Word VML images: <v:imagedata src="...">
    for v in soup.find_all(lambda tag: tag.name and ":" in tag.name and tag.name.endswith("imagedata")):
        src = v.get("src")
        local = _resolve_local(src, html_dir, companions)
        if local:
            data_uri = _read_binary_as_data_uri(local)
            if data_uri:
                v["src"] = data_uri

    # Remove conditional comments and MSO cruft
    for c in soup.find_all(string=lambda text: isinstance(text, Comment)):
        c.extract()
    raw = str(soup)
    raw = re.sub(r"<!--\[if.*?<!\[endif\]-->", "", raw, flags=re.S | re.I)
    raw = re.sub(r'\sclass="[^"]*\bmso-[^"]*"', "", raw)
    raw = re.sub(r'\sstyle="[^"]*mso-[^"]*"', "", raw)
    raw = re.sub(r'\sxmlns(:\w+)?="urn:schemas-microsoft-com:[^"]*"', "", raw)
    raw = re.sub(r'\sxmlns(:\w+)?="vml:[^"]*"', "", raw)

    soup2 = BeautifulSoup(raw, "html.parser")
    for s in soup2.find_all("script"):
        s.decompose()
    for sp in list(soup2.find_all("span")):
        if not sp.get_text(strip=True) and not sp.attrs:
            sp.decompose()

    head = soup2.head or soup2.new_tag("head")
    if not soup2.head:
        if soup2.html:
            soup2.html.insert(0, head)
        else:
            soup2.insert(0, head)
    if not head.find("meta", attrs={"charset": True}):
        meta = soup2.new_tag("meta")
        meta.attrs["charset"] = "utf-8"
        head.insert(0, meta)

    return str(soup2)


def handle_html_input(src_path: Path, workdir: Optional[Path] = None) -> Path:
    """
    Given a .htm/.html saved from Word/Excel (with a companion assets folder),
    inline all local assets and lightly clean for WordPress/Astra.
    Returns the path to a single, ready-to-upload HTML file.
    """
    if src_path.suffix.lower() not in (".htm", ".html"):
        raise ValueError(f"handle_html_input expects .htm/.html, got: {src_path}")

    html_text = _read_text_guess_encoding(src_path)
    cleaned = _inline_assets_and_clean(html_text, src_path)

    out_dir = Path(workdir) if workdir else Path(tempfile.mkdtemp(prefix="wp_inline_"))
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / (src_path.stem + "_inlined.html")
    out_path.write_text(cleaned, encoding="utf-8")
    return out_path


def handle_html_input_with_warning(src_path: Path, parent_window=None, workdir: Optional[Path] = None) -> Path:
    """
    Like handle_html_input(), but warns user if no companion assets folder is found.
    """
    companions = _detect_companion_dirs(src_path)
    if not companions:
        warning = (
            "The selected HTML file does not appear to have a companion folder "
            "(e.g. 'filename_files' or 'filename.fld') containing images and styles.\n\n"
            "If you saved this from Word or Excel, please make sure to keep the HTML file "
            "and its companion folder together in the same directory.\n\n"
            "You can still proceed, but images and layout styles may be missing."
        )
        if HAS_QT and parent_window is not None:
            msg = QMessageBox(parent_window)
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Missing assets folder")
            msg.setText(warning)
            msg.addButton("Proceed", QMessageBox.AcceptRole)
            msg.addButton("Cancel", QMessageBox.RejectRole)
            if msg.exec() != QMessageBox.AcceptRole:
                raise RuntimeError("User cancelled upload due to missing assets folder.")
        else:
            # Console fallback
            print("[WARNING] " + warning.replace("\n", " "))

    return handle_html_input(src_path, workdir)


# ---------- Entry point used by your pipeline ----------
def process_file_for_upload(file_path: Path, workdir: Path, parent_window=None) -> Path:
    """
    Unified entry point:
    - .htm/.html  → handle_html_input_with_warning → returns single-file HTML
    - .docx       → convert_word_to_html (COM on Windows, soffice elsewhere)
    - .xlsx       → convert_excel_to_html (COM on Windows, soffice elsewhere)
    Returns path to a single HTML file ready for upload to WordPress.
    """
    ext = file_path.suffix.lower()
    if ext in (".html", ".htm"):
        return handle_html_input_with_warning(file_path, parent_window, workdir)
    elif ext == ".docx":
        return convert_word_to_html(file_path, workdir)
    elif ext == ".xlsx":
        return convert_excel_to_html(file_path, workdir)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
