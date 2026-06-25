"""
frontend/build.py
掃描 模組/ 資料夾，以 CP950 讀取所有 VBA 檔案內容，
產生 frontend/data.json 與 frontend/app-data.js。

使用方式：
    python frontend/build.py
"""

import base64
import json
import os
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
MODULES_DIR = os.path.join(REPO_ROOT, "模組")
OUTPUT_JSON = os.path.join(SCRIPT_DIR, "data.json")
OUTPUT_JS = os.path.join(SCRIPT_DIR, "app-data.js")

VBA_EXTENSIONS = {".bas", ".cls", ".frm"}


def read_file_cp950(filepath):
    try:
        with open(filepath, "r", encoding="cp950", errors="replace") as f:
            return f.read()
    except Exception as e:
        return f"[無法讀取: {e}]"


def read_file_bytes_b64(filepath):
    """讀取原始 bytes 並轉成 base64，供前端原始編碼下載使用。"""
    try:
        with open(filepath, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    except Exception:
        return ""


def scan_modules(modules_dir):
    folders = []

    # 根目錄下的 .bas 等檔案（非子資料夾）
    root_files = []
    for entry in sorted(os.listdir(modules_dir)):
        full_path = os.path.join(modules_dir, entry)
        if os.path.isfile(full_path):
            ext = os.path.splitext(entry)[1].lower()
            if ext in VBA_EXTENSIONS:
                rel_path = os.path.relpath(full_path, REPO_ROOT).replace("\\", "/")
                root_files.append({
                    "name": entry,
                    "path": rel_path,
                    "content": read_file_cp950(full_path),
                    "content_b64": read_file_bytes_b64(full_path),
                })

    if root_files:
        folders.append({
            "folder": "（根目錄）",
            "path": os.path.relpath(modules_dir, REPO_ROOT).replace("\\", "/"),
            "files": root_files,
        })

    # 子資料夾
    for entry in sorted(os.listdir(modules_dir)):
        full_path = os.path.join(modules_dir, entry)
        if os.path.isdir(full_path):
            files = []
            for fname in sorted(os.listdir(full_path)):
                fpath = os.path.join(full_path, fname)
                if os.path.isfile(fpath):
                    ext = os.path.splitext(fname)[1].lower()
                    if ext in VBA_EXTENSIONS:
                        rel_path = os.path.relpath(fpath, REPO_ROOT).replace("\\", "/")
                        files.append({
                            "name": fname,
                            "path": rel_path,
                            "content": read_file_cp950(fpath),
                            "content_b64": read_file_bytes_b64(fpath),
                        })
            if files:
                folders.append({
                    "folder": entry,
                    "path": os.path.relpath(full_path, REPO_ROOT).replace("\\", "/"),
                    "files": files,
                })

    return folders


def build():
    if not os.path.isdir(MODULES_DIR):
        print(f"[錯誤] 找不到模組資料夾：{MODULES_DIR}", file=sys.stderr)
        sys.exit(1)

    print(f"掃描：{MODULES_DIR}")
    folders = scan_modules(MODULES_DIR)

    total_files = sum(len(f["files"]) for f in folders)
    total_folders = len(folders)

    data = {
        "generated_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "total_folders": total_folders,
        "total_files": total_files,
        "modules": folders,
    }

    # 輸出 data.json
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[OK] 輸出：{OUTPUT_JSON}（{total_folders} 資料夾，{total_files} 檔案）")

    # 輸出 app-data.js（供 file:// 本地開啟使用，避免 CORS）
    js_content = "// 由 build.py 自動產生，請勿手動修改\n"
    js_content += "window.VBA_DATA = " + json.dumps(data, ensure_ascii=False) + ";\n"
    with open(OUTPUT_JS, "w", encoding="utf-8") as f:
        f.write(js_content)
    print(f"[OK] 輸出：{OUTPUT_JS}")
    print("完成！請用瀏覽器開啟 frontend/index.html")


if __name__ == "__main__":
    build()
