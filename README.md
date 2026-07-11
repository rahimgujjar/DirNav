# DirNav: Advanced File System Reconnaissance & Mapper

## 📑 Executive Summary
DirNav is a high-performance, multi-threaded directory mapping utility designed for deep file system reconnaissance. Engineered in Python with a PyQt5 native interface, it bypasses standard OS limitations to provide recursive, structural mapping of complex nested directories and compressed archives.

> [!TIP]
> **View the Raw Source Code:**
> You can bypass the cloning process and directly view or download the main execution engine here:
> **[`DirNav.py`](https://github.com/rahimgujjar/DirNav/blob/main/src/DirNav.py)**.

Designed for system administrators, forensic analysts, and developers, DirNav seamlessly inspects physical disks and virtual archives (ZIP, APK, JAR, CBZ) without requiring explicit extraction, generating comprehensive, portable data blueprints.

---

## 🚀 Core Capabilities

* **Asynchronous Archive Inspection:** Utilizes `ThreadPoolExecutor` to recursively scan `.zip`, `.apk`, and `.jar` files in parallel. It maps internal archive structures and calculates file allocations entirely in memory.
* **Native OS Integration:** Leverages PyQt5 to maintain native Windows 10/11 UX for file and folder intent selection, eliminating clunky custom dialogs.
* **Deep Metadata Resolution:**
  * **MIME Detection:** Integrates `python-magic` for exact file signature validation, ignoring easily spoofed file extensions.
  * **Media Telemetry:** Extracts exact image resolutions (via `Pillow`) directly from physical disks and embedded archive binaries.
  * **Shortcut Resolution:** Utilizes `pywin32` COM interfaces to resolve `.lnk` target paths dynamically.
* **Scalable Telemetry Export:** Dynamically compiles mapped system structures into heavily nested JSON, plaintext, or dynamically wrapped PDF ledgers (via `reportlab`).

---

## 🛠️ Technical Architecture

### Tech Stack
* **Language:** Python 3.x
* **GUI Framework:** PyQt5
* **Concurrency:** `concurrent.futures.ThreadPoolExecutor`
* **Optional Dependencies:** `python-magic-bin`, `pywin32`, `Pillow`, `reportlab`

### The "Smart Intent" Pipeline
Standard OS dialogs struggle with selecting either a "File" or a "Folder" dynamically. DirNav solves this by prompting the user for an **Intent** (Directory Mapping vs. Archive Inspection) prior to invoking the native OS picker. This allows the application to cleanly adapt its recursive depth limits and multi-threading allocation strategies based on the target vector.

---

## 📥 Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/rahimgujjar/DirNav.git
cd DirNav

```

### 2. Install Core Dependencies

To utilize the full feature set (MIME mapping, PDF generation, and LNK resolution), install the requirements:

```bash
pip install -r requirements.txt

```

### 3. Execution

Run the GUI application (the `.pyw` extension ensures it runs without a persistent background console on Windows environments):

```bash
python src/DirNav.py

```

---

## 🛡️ Privacy & Execution Note

DirNav operates entirely offline and requires zero elevated network permissions. All directory mapping, MIME evaluation, and PDF generation occur locally within the host machine's memory boundaries.
