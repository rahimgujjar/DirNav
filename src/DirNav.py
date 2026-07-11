#!/usr/bin/env python3
"""
DirNav.py

Features:
- Solves the "File vs Folder" native dialog limitation by asking for INTENT first.
- If you choose "Folder", it launches the Native Folder Picker.
- If you choose "Zip File", it launches the Native File Picker.
- No "dummy file" selection required.
- Maintains the exact Native Windows 11/10 look for the browsing experience.
- Recursively scans folders and inspects .zip files.
- Outputs .txt, .json, and .mmd (Mermaid) graphs.

Required:
    pip install PyQt5
Optional:
    pip install python-magic-bin    # for MIME detection
    pip install pywin32             # for resolving .lnk on disk
"""
import io
import os
import sys
import json
import zipfile
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

from PyQt5 import QtWidgets, QtCore, QtGui

# ----------------- PDF Setup -----------------
try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.pagesizes import letter, landscape
    from xml.sax.saxutils import escape
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ----------------- Optional features -----------------
try:
    import magic
    MAGIC_AVAILABLE = True
    _magic_mime = magic.Magic(mime=True)
except Exception:
    MAGIC_AVAILABLE = False
    _magic_mime = None

try:
    import win32com.client
    WIN32_AVAILABLE = True
except Exception:
    WIN32_AVAILABLE = False

# ----------------- Image Resolution Setup -----------------
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.tiff'}

# ----------------- Configuration -----------------
def auto_workers():
    cores = os.cpu_count() or 4
    return min(64, max(1, cores * 2))

MAX_WORKERS = auto_workers()
MERMAID_MAX_DEPTH = 4
MERMAID_MAX_NODES = 2000

# ----------------- Utilities -----------------
def timestamp_str():
    return datetime.now().strftime("%d-%m-%Y_%I-%M-%S_%p")

def safe_label(name, max_len=80):
    s = str(name).replace('"', '\\"').replace('\n', ' ').strip()
    return s if len(s) <= max_len else s[:max_len-3] + "..."

def detect_mime_from_path(path):
    if not MAGIC_AVAILABLE:
        return None
    try:
        with open(path, 'rb') as f:
            head = f.read(2048)
        return _magic_mime.from_buffer(head)
    except Exception:
        return None

def detect_mime_from_zip_entry(zip_obj, entry_name):
    if not MAGIC_AVAILABLE:
        return None
    try:
        with zip_obj.open(entry_name) as ef:
            head = ef.read(2048)
        return _magic_mime.from_buffer(head)
    except Exception:
        return None

def resolve_lnk_on_disk(path):
    if not WIN32_AVAILABLE:
        return None
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(str(path))
        return shortcut.Targetpath or shortcut.TargetPath or None
    except Exception:
        return None

def get_resolution_from_path(path):
    if not PIL_AVAILABLE: return None
    try:
        with Image.open(path) as img:
            return f"{img.width}x{img.height}"
    except Exception:
        return None

def get_resolution_from_zip(zip_obj, entry_name):
    if not PIL_AVAILABLE: return None
    try:
        with zip_obj.open(entry_name) as ef:
            # Load bytes into memory to let Pillow read it directly from the ZIP
            with Image.open(io.BytesIO(ef.read())) as img:
                return f"{img.width}x{img.height}"
    except Exception:
        return None

def calculate_folder_sizes(node):
    """Recursively calculates and assigns size_bytes and size_mb for directories."""
    total_bytes = 0

    # If it's a file or zip-file, return its size
    if node.get("type") in ("file", "zip-file"):
        return node.get("size_bytes", 0)

    # If it's a directory or zip-root, sum the children
    for child in node.get("children", []):
        total_bytes += calculate_folder_sizes(child)

    # If a file has an inspected ZIP attached to it, add that to the total size too
    if "zip_contents" in node and isinstance(node["zip_contents"], dict):
        if "structured" in node["zip_contents"]:
            # Note: We don't add this to total_bytes to avoid double counting the zip file itself,
            # but we still want to calculate the internal folder sizes of the zip contents.
            calculate_folder_sizes(node["zip_contents"]["structured"])

    # Assign sizes to the current folder node
    node["size_bytes"] = total_bytes
    node["size_mb"] = round(total_bytes / (1024 * 1024), 4)
    return total_bytes

# ----------------- ZIP inspection -----------------
def inspect_zip_list_structured(zip_path, max_depth=None):
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            raw = z.namelist()
            root = {"name": Path(zip_path).name, "type": "zip-root", "children": []}
            for entry in raw:
                parts = [p for p in entry.split('/') if p != ""]

                # --- NEW CODE FOR DEPTH LIMIT IN ZIP ---
                entry_depth = len(parts) - (0 if entry.endswith('/') else 1)
                if max_depth is not None and entry_depth > max_depth:
                    continue
                # ---------------------------------------

                node = root
                for i, part in enumerate(parts):
                    is_last = (i == len(parts) - 1)
                    if not is_last:
                        found = next((c for c in node["children"] if c["name"] == part and c.get("type") == "dir"), None)
                        if not found:
                            found = {"name": part, "type": "dir", "children": []}
                            node["children"].append(found)
                        node = found
                    else:
                        if entry.endswith('/'):
                            found = next((c for c in node["children"] if c["name"] == part and c.get("type") == "dir"), None)
                            if not found:
                                found = {"name": part, "type": "dir", "children": []}
                                node["children"].append(found)
                        else:
                            try:
                                info = z.getinfo(entry)
                                size_b = info.file_size
                                size_mb = round(size_b / (1024 * 1024), 4)
                            except KeyError:
                                size_b, size_mb = 0, 0.0

                            file_node = {
                                "name": part,
                                "type": "zip-file",
                                "zip_path": entry,
                                "size_bytes": size_b,
                                "size_mb": size_mb
                            }

                            # --- NEW CODE FOR ZIP RESOLUTION ---
                            ext = Path(part).suffix.lower()
                            if ext in IMAGE_EXTS:
                                res = get_resolution_from_zip(z, entry)
                                if res:
                                    file_node["resolution"] = res
                            # -----------------------------------

                            if MAGIC_AVAILABLE:
                                try:
                                    mime = detect_mime_from_zip_entry(z, entry)
                                    if mime:
                                        file_node["mime"] = mime
                                except Exception:
                                    pass
                            node["children"].append(file_node)
            return {"structured": root}
    except Exception as e:
        return {"error": str(e)}

# ----------------- Filesystem scanning -----------------
def scan_folder_collect(root_path, max_depth=None, progress_callback=None):
    root = Path(root_path)
    tree = {"name": root.name, "path": Path(root).as_posix(), "type": "dir", "children": []}
    nodes = {str(root): tree}
    zip_tasks = []

    for dirpath, dirs, files in os.walk(root):
        # --- NEW CODE FOR LOCAL FOLDER DEPTH LIMIT ---
        current_depth = len(Path(dirpath).relative_to(root).parts)
        if max_depth is not None and current_depth >= max_depth:
            dirs[:] = []  # Prevent digging deeper into subdirectories
        # ---------------------------------------------

        dirnode = nodes.get(str(Path(dirpath)))
        if dirnode is None:
            try:
                rel = Path(dirpath).relative_to(root)
                node = tree
                for p in rel.parts:
                    found = next((c for c in node["children"] if c["name"]==p and c["type"]=="dir"), None)
                    if not found:
                        found = {"name": p, "path": (Path(node["path"])/p).as_posix(), "type": "dir", "children": []}
                        node["children"].append(found)
                    node = found
                dirnode = node
                nodes[str(Path(dirpath))] = dirnode
            except ValueError:
                continue

        for d in dirs:
            child = {"name": d, "path": (Path(dirpath)/d).as_posix(), "type": "dir", "children": []}
            dirnode["children"].append(child)
            nodes[str(Path(dirpath)/d)] = child

        for f in files:
            fp = Path(dirpath)/f
            try:
                size_b = fp.stat().st_size
                size_mb = round(size_b / (1024 * 1024), 4)
            except OSError:
                size_b, size_mb = 0, 0.0

            file_node = {
                "name": f,
                "path": Path(fp).as_posix(),
                "type": "file",
                "size_bytes": size_b,
                "size_mb": size_mb
            }

            # --- NEW CODE FOR FOLDER RESOLUTION ---
            ext = fp.suffix.lower()
            if ext in IMAGE_EXTS:
                res = get_resolution_from_path(fp)
                if res:
                    file_node["resolution"] = res
            # --------------------------------------

            if MAGIC_AVAILABLE:
                try:
                    mime = detect_mime_from_path(fp)
                    if mime:
                        file_node["mime"] = mime
                except Exception:
                    pass
            if f.lower().endswith(".lnk"):
                if WIN32_AVAILABLE:
                    try:
                        target = resolve_lnk_on_disk(fp)
                        if target:
                            file_node["lnk_target"] = target
                    except Exception:
                        file_node["lnk_target_error"] = "resolve_failed"
                else:
                    file_node["lnk_note"] = "pywin32 not installed"
            if f.lower().endswith((".zip", ".apk", ".jar", ".cbz")):
                file_node["zip_action"] = "inspect"
                zip_tasks.append((str(fp), file_node))
            dirnode["children"].append(file_node)

        if progress_callback:
            progress_callback(f"Scanned: {dirpath}")

    # Inspect ZIPs in parallel
    if zip_tasks:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            future_map = {}
            for path, node_ref in zip_tasks:
                future = ex.submit(inspect_zip_list_structured, path, max_depth) # Apply same depth to internal zips
                future_map[future] = (path, node_ref)
            for fut in as_completed(future_map):
                path, node_ref = future_map[fut]
                try:
                    res = fut.result()
                    node_ref["zip_contents"] = res
                except Exception as e:
                    node_ref["zip_contents"] = {"error": str(e)}
                if progress_callback:
                    progress_callback(f"Processed ZIP: {node_ref.get('path','')}")
    return tree

# ----------------- Mermaid generation -----------------
def mermaid_lines_limited(node, max_depth=MERMAID_MAX_DEPTH, max_nodes=MERMAID_MAX_NODES):
    lines = []
    counter = {"n": 0, "limit": max_nodes}
    def _sub(nod, nid, depth):
        if counter["n"] >= counter["limit"]:
            return [], nid
        lines_local = []
        myid = f"n{nid}"
        label = safe_label(nod.get("name",""))
        lines_local.append(f'{myid}["{label}"]')
        counter["n"] += 1
        cid = nid + 1
        if depth >= max_depth:
            return lines_local, cid

        children = nod.get("children", [])
        if nod.get("type") == "file" and nod.get("zip_contents") and isinstance(nod["zip_contents"], dict):
            z = nod["zip_contents"]
            if "structured" in z:
                children = z["structured"].get("children", [])

        for c in children:
            if counter["n"] >= counter["limit"]:
                break
            lines_local.append(f'{myid} --> n{cid}')
            sub_lines, cid = _sub(c, cid, depth+1)
            lines_local.extend(sub_lines)
        return lines_local, cid
    top_lines, _ = _sub(node, 0, 0)
    return top_lines

# ----------------- Worker thread -----------------
class ScanThread(QtCore.QThread):
    finished_signal = QtCore.pyqtSignal(object, object)
    progress_signal = QtCore.pyqtSignal(str)
    error_signal = QtCore.pyqtSignal(str)

    def __init__(self, target, max_depth=None):
        super().__init__()
        self.target = target
        self.max_depth = max_depth

    def run(self):
        try:
            p = Path(self.target)
            if p.is_file() and p.suffix.lower() == ".zip":
                tree = {"name": p.name, "path": Path(p).as_posix(), "type": "file"}
                self.progress_signal.emit(f"Inspecting ZIP: {p}")
                tree["zip_contents"] = inspect_zip_list_structured(str(p), self.max_depth)
                out_root = p.parent
            else:
                tree = scan_folder_collect(self.target, self.max_depth, progress_callback=self.progress_signal.emit)
                out_root = Path(self.target)

            # Calculate all directory sizes recursively before emitting
            calculate_folder_sizes(tree)

            self.finished_signal.emit(tree, out_root)
        except Exception as e:
            self.error_signal.emit(str(e))

# ----------------- Mode Selection Dialog -----------------
class ModeSelector(QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Select Scan Mode")
        self.resize(350, 200) # Increased height slightly for the new input field

        layout = QtWidgets.QVBoxLayout(self)

        lbl = QtWidgets.QLabel("What would you like to select?")
        lbl.setAlignment(QtCore.Qt.AlignCenter)
        lbl.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(lbl)

        # --- NEW: Depth Input UI ---
        self.depth_input = QtWidgets.QLineEdit()
        self.depth_input.setPlaceholderText("Max Depth (Leave blank for infinite)")
        self.depth_input.setAlignment(QtCore.Qt.AlignCenter)
        self.depth_input.setStyleSheet("font-size: 12px; padding: 5px; margin-bottom: 10px;")
        layout.addWidget(self.depth_input)
        # ---------------------------

        btn_layout = QtWidgets.QHBoxLayout()

        # Button 1: Folder
        self.btn_folder = QtWidgets.QPushButton("Select a Folder")
        self.btn_folder.setMinimumHeight(50)
        self.btn_folder.setStyleSheet("font-size: 12px;")

        # Button 2: Zip File
        self.btn_zip = QtWidgets.QPushButton("Select a Zip File")
        self.btn_zip.setMinimumHeight(50)
        self.btn_zip.setStyleSheet("font-size: 12px;")

        btn_layout.addWidget(self.btn_folder)
        btn_layout.addWidget(self.btn_zip)

        layout.addLayout(btn_layout)

        self.selection_type = None  # 'folder' or 'zip'
        self.max_depth = None       # NEW

        self.btn_folder.clicked.connect(self.select_folder)
        self.btn_zip.clicked.connect(self.select_zip)

    def _parse_depth(self):
        txt = self.depth_input.text().strip()
        if txt.isdigit():
            self.max_depth = int(txt)
        else:
            self.max_depth = None

    def select_folder(self):
        self._parse_depth()
        self.selection_type = 'folder'
        self.accept()

    def select_zip(self):
        self._parse_depth()
        self.selection_type = 'zip'
        self.accept()

# ----------------- Save Format Dialog -----------------
class SaveFormatDialog(QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Select Output Format")
        # --- FIXED: Increased height so the OK button is visible ---
        self.resize(350, 250)

        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)

        layout = QtWidgets.QVBoxLayout(self)

        lbl = QtWidgets.QLabel("Scan complete! Select output format:")
        lbl.setStyleSheet("font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(lbl)

        self.radio_json = QtWidgets.QRadioButton("JSON only (Default)")
        self.radio_json.setChecked(True)
        self.radio_txt = QtWidgets.QRadioButton("Text only")
        self.radio_pdf = QtWidgets.QRadioButton("PDF only")
        self.radio_all = QtWidgets.QRadioButton("All Formats (JSON, TXT, PDF)")

        layout.addWidget(self.radio_json)
        layout.addWidget(self.radio_txt)
        layout.addWidget(self.radio_pdf)
        layout.addWidget(self.radio_all)

        # --- The OK button ---
        self.btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok)
        self.btn_box.accepted.connect(self.accept)
        layout.addWidget(self.btn_box)

    def get_format(self):
        if self.radio_txt.isChecked(): return "txt"
        if self.radio_pdf.isChecked(): return "pdf"
        if self.radio_all.isChecked(): return "all"
        return "json"

    # Map Enter, Return, and Space to instantly accept the default
    def keyPressEvent(self, event):
        if event.key() in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
            self.accept()
        else:
            super().keyPressEvent(event)

# ----------------- Main flow -----------------
def main():
    app = QtWidgets.QApplication(sys.argv)

    # Step 1: Ask user for intent using a clean dialog
    mode_selector = ModeSelector()
    if mode_selector.exec_() != QtWidgets.QDialog.Accepted:
        sys.exit(0)

    target_path = None
    start_dir = str(Path.home())

    # Step 2: Open the CORRECT Native Dialog based on intent
    if mode_selector.selection_type == 'folder':
        # Native Folder Picker (Image 1 style for folders)
        target_path = QtWidgets.QFileDialog.getExistingDirectory(
            None,
            "Select Folder to Scan",
            start_dir,
            QtWidgets.QFileDialog.ShowDirsOnly  # Forces native folder picker
        )
    elif mode_selector.selection_type == 'zip':
        # Native File Picker (Image 1 style for files)
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(
            None,
            "Select Zip File to Inspect",
            start_dir,
            "Zip Files (*.zip);;All Files (*)"
        )
        target_path = file_name

    if not target_path:
        sys.exit(0)

    # Step 3: Run the scan (now passing max_depth)
    thread = ScanThread(target_path, mode_selector.max_depth)

    def on_progress(msg):
        print(msg)

    def on_finished(tree, out_root):
        try:
            save_dialog = SaveFormatDialog()
            save_dialog.exec_()
            selected_format = save_dialog.get_format()

            base_name = Path(tree["path"]).stem if tree.get("type")=="file" else tree["name"]
            ts = timestamp_str()
            json_name = f"{base_name}_{ts}.json"
            txt_name = f"{base_name}_{ts}.txt"

            dl = QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.DownloadLocation)
            out_root = Path(dl) if dl else (Path.home() / "Downloads")
            out_root.mkdir(parents=True, exist_ok=True)

            saved_files = []

            # Write JSON if selected
            if selected_format in ("json", "both", "all"):
                with open(out_root / json_name, "w", encoding="utf-8") as f:
                    json.dump(tree, f, indent=2, ensure_ascii=False)
                saved_files.append(json_name)

            # Write TXT if selected
            if selected_format in ("txt", "both", "all"):
                with open(out_root / txt_name, "w", encoding="utf-8") as f:
                    json.dump(tree, f, indent=2, ensure_ascii=False)
                saved_files.append(txt_name)

            # --- FIXED: Smarter PDF Generation ---
            if selected_format in ("pdf", "all"):
                if REPORTLAB_AVAILABLE:
                    pdf_name = f"{base_name}_{ts}.pdf"
                    pdf_path = out_root / pdf_name

                    # 1. Convert dict to string
                    json_str = json.dumps(tree, indent=2, ensure_ascii=True)

                    # 2. Setup document in Landscape (horizontal) for wider paths
                    doc = SimpleDocTemplate(str(pdf_path), pagesize=landscape(letter),
                                            leftMargin=15, rightMargin=15, topMargin=15, bottomMargin=15)
                    styles = getSampleStyleSheet()

                    # 3. Create a custom style that forces word wrapping smoothly
                    custom_style = ParagraphStyle(
                        'CustomJSON',
                        parent=styles['Normal'],
                        fontName='Courier',
                        fontSize=8,
                        leading=10,
                        wordWrap='CJK' # This allows long paths to wrap seamlessly
                    )

                    # 4. Escape XML characters and lock spacing so indents don't collapse
                    safe_text = escape(json_str).replace('\n', '<br/>').replace('  ', '&nbsp;&nbsp;')

                    # 5. Build the PDF
                    flowables = [Paragraph(safe_text, custom_style)]
                    doc.build(flowables)
                    saved_files.append(pdf_name)
                else:
                    saved_files.append("[PDF FAILED: reportlab not installed]")

            info = f"Created files in:\n{out_root}\n\n" + "\n".join(saved_files)

            msg = QtWidgets.QMessageBox()
            msg.setWindowFlags(msg.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
            msg.setWindowTitle("Completed")
            msg.setText(info)
            msg.exec_()

        except Exception as e:
            err = QtWidgets.QMessageBox()
            err.setWindowFlags(err.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
            err.setWindowTitle("Error writing outputs")
            err.setText(str(e))
            err.exec_()
        finally:
            QtCore.QCoreApplication.quit()

    def on_error(msg):
        QtWidgets.QMessageBox.critical(None, "Scan error", msg)
        QtCore.QCoreApplication.quit()

    thread.progress_signal.connect(on_progress)
    thread.finished_signal.connect(on_finished)
    thread.error_signal.connect(on_error)
    thread.start()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()