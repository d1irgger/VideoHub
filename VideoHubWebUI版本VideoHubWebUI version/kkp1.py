import os
import cv2
import webbrowser
import math
import re
import json
import http.server
import socketserver
import urllib.parse
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import tkinter as tk
from tkinter import filedialog
import win32com.client

# ========== 工具函数 ==========
def clean_string(raw_str):
    return str(raw_str).replace('\n', '').replace('\t', '').strip()

def get_windows_title(file_path):
    if os.name != "nt":
        return ""
    shell = folder = file_item = None
    try:
        norm_path = os.path.normpath(file_path)
        folder_path = os.path.dirname(norm_path)
        file_name = os.path.basename(norm_path)
        shell = win32com.client.DispatchEx("Shell.Application")
        folder = shell.NameSpace(folder_path)
        if folder is None:
            return ""
        file_item = folder.ParseName(file_name)
        if file_item is not None:
            try:
                return clean_string(folder.GetDetailsOf(file_item, 21))
            except Exception:
                pass
    except Exception:
        pass
    finally:
        for obj in (file_item, folder, shell):
            if obj is not None:
                try:
                    del obj
                except Exception:
                    pass
    return ""

def extract_url(text):
    if not text:
        return None
    text = text.replace('\n', ' ').replace('\t', ' ').strip()
    m = re.search(r'https?://\S+', text)
    return m.group(0) if m else None

def natural_key(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', text)]

def format_size(size_bytes):
    if size_bytes <= 0:
        return "0 B"
    units = ("B", "KB", "MB", "GB", "TB")
    idx = min(int(math.log(size_bytes, 1024)), len(units) - 1)
    return f"{size_bytes / (1024 ** idx):.2f} {units[idx]}"

def format_duration(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

# ========== 配置与数据类 ==========
class Config:
    VIDEO_EXT = (".mp4", ".mov", ".mkv", ".avi", ".flv", ".webm", ".wmv", ".mpg", ".mpeg")
    PARSE_WORKERS = min(8, (os.cpu_count() or 4) + 2)

class VideoItem:
    __slots__ = ("name", "path", "resolution", "duration_str", "size_str", "mtime", "title", "jump_url")
    def __init__(self, name, path, resolution, duration_str, size_str, mtime, title, jump_url):
        self.name = name
        self.path = path
        self.resolution = resolution
        self.duration_str = duration_str
        self.size_str = size_str
        self.mtime = mtime
        self.title = title
        self.jump_url = jump_url

# ========== 视频解析 ==========
class VideoParser:
    @staticmethod
    def parse_video(file_path):
        width = height = 0
        duration = 0.0
        try:
            cap = cv2.VideoCapture(file_path)
            if cap.isOpened():
                width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
                height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
                fps = cap.get(cv2.CAP_PROP_FPS) or 0
                total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0
                if fps > 0 and total_frames > 0:
                    duration = total_frames / fps
            cap.release()
        except Exception:
            pass

        size_bytes = 0
        try:
            size_bytes = os.path.getsize(file_path)
        except Exception:
            pass

        mtime = ""
        try:
            mtime = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass

        title = get_windows_title(file_path)
        jump_url = extract_url(title)
        resolution = f"{width}×{height}" if width and height else "分辨率未知"

        return VideoItem(
            name=os.path.basename(file_path),
            path=file_path,
            resolution=resolution,
            duration_str=format_duration(duration),
            size_str=format_size(size_bytes),
            mtime=mtime,
            title=title,
            jump_url=jump_url
        )

# ========== 扫描器 ==========
class VideoScanner:
    def __init__(self, scan_subfolder=True):
        self.scan_subfolder = scan_subfolder

    def scan(self, root_dir):
        result = []
        if self.scan_subfolder:
            for dirpath, _, files in os.walk(root_dir):
                for f in files:
                    if f.lower().endswith(Config.VIDEO_EXT):
                        result.append(os.path.join(dirpath, f))
        else:
            try:
                for f in os.listdir(root_dir):
                    fp = os.path.join(root_dir, f)
                    if os.path.isfile(fp) and f.lower().endswith(Config.VIDEO_EXT):
                        result.append(fp)
            except Exception:
                pass
        return result

# ========== 全局状态 ==========
scanner = VideoScanner(scan_subfolder=True)
use_browser = False
video_items = []
filtered_items = []
current_folder = ""

# ========== 多线程 HTTP 服务 ==========
class ThreadingHTTPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    allow_reuse_address = True
    daemon_threads = True

class MyHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if path == "/":
            self.send_html(self._index_html())
        elif path == "/api/files":
            self._api_files(query)
        elif path == "/api/play":
            self._api_play(query)
        elif path == "/api/export":
            self._api_export()
        elif path == "/api/config":
            self._api_config(query)
        elif path == "/api/select_folder":
            self._api_select_folder()
        else:
            self.send_error(404)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/scan":
            content_len = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_len)
            try:
                data = json.loads(post_data.decode("utf-8"))
                folder = data.get("folder", "").strip()
            except Exception:
                folder = ""
            self._api_scan(folder)
        else:
            self.send_error(404)

    def send_json(self, data):
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def send_html(self, content):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(content.encode("utf-8"))

    def send_file(self, filepath, content_type="application/octet-stream", disposition="attachment"):
        try:
            with open(filepath, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            filename = os.path.basename(filepath)
            encoded_filename = urllib.parse.quote(filename, safe="")
            ascii_filename = filename.encode("ascii", "ignore").decode("ascii") or "download"
            self.send_header(
                "Content-Disposition",
                f"{disposition}; filename*=UTF-8''{encoded_filename}; filename=\"{ascii_filename}\""
            )
            self.end_headers()
            self.wfile.write(data)
        except Exception as e:
            self.send_error(500, str(e))

    def _api_select_folder(self):
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            folder_selected = filedialog.askdirectory(title="选择视频文件夹")
            root.destroy()
            self.send_json({"folder": folder_selected or ""})
        except Exception as e:
            self.send_json({"error": str(e)})

    def _api_scan(self, folder):
        global video_items, filtered_items, current_folder
        if not folder or not os.path.isdir(folder):
            self.send_json({"error": "无效的文件夹路径"})
            return

        current_folder = folder
        try:
            paths = scanner.scan(folder)
            items = []

            with ThreadPoolExecutor(max_workers=Config.PARSE_WORKERS) as executor:
                future_to_path = {executor.submit(VideoParser.parse_video, p): p for p in paths}
                for future in as_completed(future_to_path):
                    try:
                        items.append(future.result())
                    except Exception as e:
                        print(f"解析失败 {future_to_path[future]}: {e}")

            items.sort(key=lambda x: natural_key(x.name))
            video_items = items
            filtered_items = items.copy()
            self.send_json({"success": True, "count": len(items)})
        except Exception as e:
            self.send_json({"error": f"扫描出错: {str(e)}"})

    def _api_files(self, query):
        global filtered_items
        try:
            page = int(query.get("page", ["1"])[0])
            per_page = int(query.get("per_page", ["7"])[0])
        except Exception:
            page, per_page = 1, 7

        keyword = query.get("keyword", [""])[0].strip()
        if keyword:
            kw_lower = keyword.lower()
            items = [item for item in video_items if kw_lower in item.name.lower()]
        else:
            items = video_items

        filtered_items = items
        total = len(items)
        total_pages = max(1, (total + per_page - 1) // per_page)
        page = max(1, min(page, total_pages))
        start = (page - 1) * per_page
        end = min(start + per_page, total)
        page_items = items[start:end]

        data = [{
            "name": it.name,
            "path": it.path,
            "resolution": it.resolution,
            "duration": it.duration_str,
            "size": it.size_str,
            "mtime": it.mtime,
            "title": it.title or "未命名",
            "jump_url": it.jump_url or "",
        } for it in page_items]

        self.send_json({
            "total": total,
            "page": page,
            "total_pages": total_pages,
            "items": data
        })

    def _api_play(self, query):
        path = query.get("path", [""])[0]
        if not path or not os.path.exists(path):
            self.send_error(404, "文件不存在")
            return
        mode = query.get("mode", ["play"])[0]

        if mode == "stream":
            ext = os.path.splitext(path)[1].lower()
            if ext in (".mp4", ".webm", ".ogg"):
                content_type = f"video/{ext[1:]}"
            else:
                content_type = "application/octet-stream"
            self.send_file(path, content_type=content_type, disposition="inline")
            return

        if use_browser:
            self.send_html(self._player_html(path))
        else:
            try:
                webbrowser.open(path)
            except Exception as e:
                self.send_error(500, f"无法启动浏览器: {str(e)}")
                return
            self.send_html("""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>正在启动</title></head>
<body style="background:#1c1c1e;color:#fff;display:flex;justify-content:center;align-items:center;height:100vh;font-family:-apple-system,BlinkMacSystemFont,sans-serif;">
    <div style="text-align:center;">
        <h2 style="font-weight:500;">正在使用浏览器打开…</h2>
        <p style="opacity:0.6;font-size:14px;">浏览器启动后此窗口将自动关闭</p>
        <script>setTimeout(function(){window.close();},1500);</script>
    </div>
</body>
</html>""")

    def _api_export(self):
        if not video_items:
            self.send_json({"error": "请先扫描视频文件夹"})
            return
        log_path = os.path.join(os.getcwd(), "log.txt")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                for item in video_items:
                    link = item.jump_url if item.jump_url else "无链接"
                    f.write(f"{item.name}\t{link}\n")
            self.send_file(log_path, content_type="text/plain")
        except Exception as e:
            self.send_json({"error": str(e)})

    def _api_config(self, query):
        global use_browser, scanner
        if "use_browser" in query:
            use_browser = query["use_browser"][0].lower() == "true"
        if "scan_subfolder" in query:
            scanner.scan_subfolder = query["scan_subfolder"][0].lower() == "true"
        self.send_json({
            "use_browser": use_browser,
            "scan_subfolder": scanner.scan_subfolder
        })

    def _player_html(self, file_path):
        safe_path = urllib.parse.quote(file_path, safe="")
        return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>播放器 - {os.path.basename(file_path)}</title>
    <style>
        body {{ background: #1c1c1e; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "Segoe UI", Arial, sans-serif; }}
        .container {{ background: rgba(255,255,255,0.08); backdrop-filter: blur(20px); border-radius: 24px; padding: 28px; box-shadow: 0 20px 60px rgba(0,0,0,0.55); max-width: 90vw; max-height: 90vh; text-align: center; }}
        video {{ width: 100%; max-height: 70vh; border-radius: 14px; background: #000; }}
        .info {{ color: #f5f5f7; margin-top: 14px; font-size: 15px; opacity: 0.85; font-weight: 500; }}
        .download-link {{ display: inline-block; margin-top: 12px; color: #0a84ff; text-decoration: none; font-weight: 500; padding: 8px 18px; border-radius: 20px; background: rgba(255,255,255,0.1); transition: background 0.2s cubic-bezier(0.25, 0.1, 0.25, 1); }}
        .download-link:hover {{ background: rgba(255,255,255,0.18); }}
        .error-msg {{ color: #ff6961; margin-top: 18px; font-size: 14px; }}
    </style>
</head>
<body>
<div class="container">
    <video controls autoplay>
        <source src="/api/play?path={safe_path}&mode=stream" type="video/mp4">
        <source src="/api/play?path={safe_path}&mode=stream" type="video/webm">
        <source src="/api/play?path={safe_path}&mode=stream" type="video/ogg">
        <div class="error-msg">
            当前浏览器不支持此视频格式，或文件无法解码。<br>
            请尝试 <a href="/api/play?path={safe_path}" download style="color:#0a84ff;">下载</a> 后本地播放。
        </div>
    </video>
    <div class="info">{os.path.basename(file_path)}</div>
    <a href="/api/play?path={safe_path}" download class="download-link">下载文件</a>
    <a href="/" style="color: #98989d; margin-left: 16px; text-decoration: none; font-size: 14px;">返回</a>
</div>
</body>
</html>"""

    def _index_html(self):
        return """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>视频搜索</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html, body { height: 100%; overflow: hidden; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Segoe UI", "Helvetica Neue", Arial, sans-serif;
            background: linear-gradient(160deg, #e8ecf1 0%, #f2f4f7 45%, #eef1f5 100%);
            display: flex; justify-content: center; align-items: center; padding: 18px;
            -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale;
        }
        .glass {
            background: rgba(255,255,255,0.68);
            backdrop-filter: blur(32px) saturate(180%);
            -webkit-backdrop-filter: blur(32px) saturate(180%);
            border-radius: 26px;
            box-shadow: 0 18px 50px rgba(0,0,0,0.07), 0 4px 14px rgba(0,0,0,0.03), inset 0 1px 0 rgba(255,255,255,0.85);
            border: 1px solid rgba(255,255,255,0.55);
            width: 100%; max-width: 1120px; height: 100%; max-height: 860px;
            padding: 22px 30px 18px; display: flex; flex-direction: column; overflow: hidden;
        }
        .header { display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; margin-bottom: 12px; flex-wrap: wrap; }
        .header h1 { font-size: 24px; font-weight: 600; letter-spacing: -0.4px; color: #1d1d1f; margin: 0; }
        .header .version { font-size: 12.5px; color: #8e8e93; font-weight: 500; letter-spacing: 0.15px; }
        .toolbar { display: flex; flex-wrap: wrap; gap: 9px; align-items: center; flex-shrink: 0; margin-bottom: 12px; }
        .toolbar input, .toolbar button {
            font-family: inherit; font-size: 14px; border: none; outline: none; border-radius: 11px; padding: 8px 15px;
            background: rgba(255,255,255,0.78); backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
            box-shadow: 0 1px 4px rgba(0,0,0,0.03);
            transition: transform 0.18s cubic-bezier(0.25, 0.1, 0.25, 1), background 0.18s cubic-bezier(0.25, 0.1, 0.25, 1), box-shadow 0.18s cubic-bezier(0.25, 0.1, 0.25, 1), border-color 0.18s cubic-bezier(0.25, 0.1, 0.25, 1);
        }
        .toolbar input { flex: 1 1 170px; min-width: 120px; border: 1px solid rgba(0,0,0,0.04); font-weight: 400; color: #1d1d1f; }
        .toolbar input::placeholder { color: #8e8e93; }
        .toolbar input:focus { background: rgba(255,255,255,0.96); border-color: #007aff; box-shadow: 0 0 0 3.5px rgba(0,122,255,0.15); }
        .toolbar button { background: #007aff; color: #ffffff; font-weight: 520; padding: 8px 16px; cursor: pointer; box-shadow: 0 2px 10px rgba(0,122,255,0.22); border: 1px solid rgba(255,255,255,0.2); letter-spacing: 0.1px; }
        .toolbar button:hover { background: #0071e3; transform: translateY(-0.5px); box-shadow: 0 4px 14px rgba(0,122,255,0.28); }
        .toolbar button:active { transform: scale(0.97); box-shadow: 0 1px 6px rgba(0,122,255,0.2); }
        .toolbar button.secondary { background: rgba(255,255,255,0.72); color: #1d1d1f; box-shadow: 0 1px 4px rgba(0,0,0,0.03); border: 1px solid rgba(0,0,0,0.05); }
        .toolbar button.secondary:hover { background: rgba(255,255,255,0.92); box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
        .toolbar .switch-group { display: flex; gap: 18px; align-items: center; margin-left: auto; font-size: 13px; color: #1d1d1f; }
        .switch-group label { display: flex; align-items: center; gap: 7px; cursor: pointer; user-select: none; }
        .switch-group label .switch-label { font-size: 13px; font-weight: 450; color: #1d1d1f; letter-spacing: -0.1px; }
        .switch-group .toggle { position: relative; width: 42px; height: 25px; flex-shrink: 0; background: #e5e5ea; border-radius: 12.5px; transition: background 0.28s cubic-bezier(0.25, 0.1, 0.25, 1); box-shadow: inset 0 1px 2px rgba(0,0,0,0.1); }
        .switch-group .toggle.active { background: #34c759; }
        .switch-group .toggle .knob { position: absolute; top: 2px; left: 2px; width: 21px; height: 21px; background: #ffffff; border-radius: 50%; box-shadow: 0 1px 4px rgba(0,0,0,0.16), 0 1px 1px rgba(0,0,0,0.08); transition: transform 0.28s cubic-bezier(0.34, 1.45, 0.64, 1); }
        .switch-group .toggle.active .knob { transform: translateX(17px); }
        .switch-group input[type="checkbox"] { display: none; }
        .list-container { flex: 1; display: flex; flex-direction: column; gap: 6px; min-height: 0; margin-top: 2px; justify-content: flex-start; }
        .file-item {
            background: rgba(255,255,255,0.58); backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
            border-radius: 13px; padding: 8px 15px; border: 1px solid rgba(255,255,255,0.4); box-shadow: 0 1px 6px rgba(0,0,0,0.02);
            transition: transform 0.2s cubic-bezier(0.25, 0.1, 0.25, 1), background 0.2s cubic-bezier(0.25, 0.1, 0.25, 1), box-shadow 0.2s cubic-bezier(0.25, 0.1, 0.25, 1);
            display: flex; flex-direction: column; justify-content: center; flex: 0 0 auto; will-change: transform;
        }
        .file-item:hover { background: rgba(255,255,255,0.82); box-shadow: 0 4px 14px rgba(0,0,0,0.045); transform: translateY(-1px); }
        .file-item .top { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 5px; }
        .file-item .top .name { font-size: 14.5px; font-weight: 520; color: #1d1d1f; word-break: break-word; letter-spacing: -0.2px; }
        .file-item .top .actions { display: flex; gap: 6px; }
        .file-item .top .actions button {
            padding: 4px 12px; font-size: 12px; border-radius: 16px; border: none; background: #007aff; color: #ffffff; cursor: pointer;
            transition: transform 0.16s cubic-bezier(0.25, 0.1, 0.25, 1), background 0.16s cubic-bezier(0.25, 0.1, 0.25, 1);
            font-weight: 510; letter-spacing: 0.05px;
        }
        .file-item .top .actions button:hover { background: #0071e3; transform: scale(1.03); }
        .file-item .top .actions button:active { transform: scale(0.97); }
        .file-item .top .actions button.web { background: #34c759; }
        .file-item .top .actions button.web:hover { background: #30b753; }
        .file-item .meta { margin-top: 2px; display: flex; flex-wrap: wrap; gap: 7px 14px; font-size: 11.5px; color: #6e6e73; letter-spacing: -0.1px; }
        .file-item .meta .title { color: #1d1d1f; opacity: 0.75; }
        .pagination-wrapper {
            display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; margin-top: 14px; padding-top: 12px;
            border-top: 1px solid rgba(0,0,0,0.045); flex-wrap: wrap; gap: 8px; min-height: 48px;
        }
        .pagination-wrapper .left { font-size: 12.5px; color: #8e8e93; flex: 0 0 auto; font-weight: 450; }
        .pagination-wrapper .center { flex: 1 1 auto; display: flex; justify-content: center; }
        .pagination-wrapper .right { font-size: 12px; color: #8e8e93; flex: 0 0 auto; }
        .pagination-wrapper .right a { color: #007aff; text-decoration: none; font-weight: 500; transition: opacity 0.18s cubic-bezier(0.25, 0.1, 0.25, 1); }
        .pagination-wrapper .right a:hover { opacity: 0.75; }
        .pagination { display: flex; justify-content: center; align-items: center; gap: 5px; flex-wrap: wrap; }
        .pagination button {
            font-family: inherit; cursor: pointer; user-select: none;
            transition: transform 0.16s cubic-bezier(0.25, 0.1, 0.25, 1), background 0.16s cubic-bezier(0.25, 0.1, 0.25, 1), box-shadow 0.16s cubic-bezier(0.25, 0.1, 0.25, 1), border-color 0.16s cubic-bezier(0.25, 0.1, 0.25, 1);
            will-change: transform;
        }
        .pagination button:disabled { opacity: 0.36; cursor: default; }
        .pagination button:hover:not(:disabled) { transform: translateY(-0.5px); }
        .pagination button:active:not(:disabled) { transform: scale(0.96); }
        .empty { text-align: center; color: #8e8e93; font-size: 15.5px; margin: auto; font-weight: 450; }
        .empty.scanning { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; padding: 20px; }
        .spinner { width: 36px; height: 36px; border: 2.8px solid rgba(0,0,0,0.07); border-top-color: #007aff; border-radius: 50%; animation: spin 0.7s linear infinite; margin: 0 auto; }
        @keyframes spin { to { transform: rotate(360deg); } }
        @media (max-width: 700px) {
            .glass { padding: 15px; border-radius: 20px; }
            .header h1 { font-size: 21px; }
            .toolbar { flex-direction: column; align-items: stretch; gap: 8px; }
            .toolbar .switch-group { margin-left: 0; justify-content: flex-start; flex-wrap: wrap; gap: 12px; }
            .pagination-wrapper { flex-direction: column; align-items: center; gap: 7px; }
        }
    </style>
</head>
<body>
<div class="glass" id="app">
    <div class="header">
        <h1>视频搜索</h1>
        <span class="version">版本 8.4.G</span>
    </div>
    <div class="toolbar">
        <button id="scanBtn">浏览</button>
        <input type="text" id="searchInput" placeholder="搜索视频" />
        <button class="secondary" id="exportBtn">导出链接</button>
        <div class="switch-group">
            <label>
                <span class="switch-label">包含子文件夹</span>
                <span class="toggle" id="subfolderToggle"><span class="knob"></span></span>
                <input type="checkbox" id="subfolderSwitch">
            </label>
            <label>
                <span class="switch-label">浏览器播放</span>
                <span class="toggle" id="browserToggle"><span class="knob"></span></span>
                <input type="checkbox" id="browserSwitch">
            </label>
        </div>
    </div>
    <div class="list-container" id="listContainer">
        <div class="empty">选择文件夹以开始</div>
    </div>
    <div class="pagination-wrapper">
        <span class="left">共 <span id="totalCount">0</span> 个项目</span>
        <div class="center"><div id="pagination"></div></div>
        <span class="right">由 <a href="//t.me/timharrys" target="_blank">D1r3ctor</a> 设计</span>
    </div>
</div>
<script>
    let currentPage = 1, totalPages = 1, totalItems = 0, scanFolder = '';
    const perPage = 7;
    const listContainer = document.getElementById('listContainer');
    const pagination = document.getElementById('pagination');
    const totalCount = document.getElementById('totalCount');
    const searchInput = document.getElementById('searchInput');

    function setupToggle(checkboxId, toggleId) {
        const checkbox = document.getElementById(checkboxId);
        const toggle = document.getElementById(toggleId);
        if (!checkbox || !toggle) return;
        if (checkbox.checked) toggle.classList.add('active');
        checkbox.addEventListener('change', function() {
            if (this.checked) toggle.classList.add('active');
            else toggle.classList.remove('active');
            this.dispatchEvent(new Event('change'));
        });
    }
    setupToggle('subfolderSwitch', 'subfolderToggle');
    setupToggle('browserSwitch', 'browserToggle');

    function loadFiles(page) {
        if (!scanFolder) {
            listContainer.innerHTML = '<div class="empty">请先选择文件夹</div>';
            pagination.innerHTML = '';
            return;
        }
        const keyword = searchInput.value.trim();
        const url = `/api/files?page=${page}&per_page=${perPage}&keyword=${encodeURIComponent(keyword)}`;
        fetch(url).then(r => r.json()).then(data => {
            if (data.error) {
                listContainer.innerHTML = `<div class="empty">${data.error}</div>`;
                return;
            }
            totalItems = data.total;
            totalPages = data.total_pages;
            currentPage = data.page;
            renderItems(data.items);
            renderPagination();
            totalCount.textContent = totalItems;
        }).catch(e => {
            listContainer.innerHTML = '<div class="empty">加载失败: ' + e.message + '</div>';
        });
    }

    function renderItems(items) {
        if (!items || items.length === 0) {
            listContainer.innerHTML = '<div class="empty">未找到视频</div>';
            return;
        }
        const containerHeight = listContainer.clientHeight;
        const gap = 6;
        const totalGap = (7 - 1) * gap;
        const cardHeight = Math.max(48, (containerHeight - totalGap) / 7);
        let html = '';
        items.forEach(item => {
            html += `
                <div class="file-item" data-path="${escapeHtml(item.path)}" data-jump="${escapeHtml(item.jump_url)}" data-name="${escapeHtml(item.name)}" style="height: ${cardHeight}px;">
                    <div class="top">
                        <span class="name">${escapeHtml(item.name)}</span>
                        <div class="actions">
                            <button class="play-local">播放</button>
                            <button class="play-web web">在浏览器中打开</button>
                        </div>
                    </div>
                    <div class="meta">
                        <span>${item.resolution} · ${item.duration} · ${item.size} · ${item.mtime}</span>
                        <span class="title">${escapeHtml(item.title)}</span>
                    </div>
                </div>`;
        });
        listContainer.innerHTML = html;
    }

    function renderPagination() {
        if (totalPages <= 1) {
            pagination.innerHTML = '';
            return;
        }
        const baseStyle = 'min-width:38px;height:34px;margin:0 1px;padding:0 11px;border-radius:9px;border:1.5px solid #d1d1d6;background:#fff;color:#1d1d1f;font-size:13.5px;font-weight:500;cursor:pointer;box-shadow:0 1px 2px rgba(0,0,0,0.05);display:inline-flex;align-items:center;justify-content:center;';
        const activeStyle = 'min-width:38px;height:34px;margin:0 1px;padding:0 11px;border-radius:9px;border:1.5px solid #007aff;background:#007aff;color:#fff;font-size:13.5px;font-weight:600;cursor:pointer;box-shadow:0 2px 7px rgba(0,122,255,0.25);display:inline-flex;align-items:center;justify-content:center;';
        const navStyle = 'min-width:38px;height:34px;margin:0 1px;padding:0 10px;border-radius:9px;border:1.5px solid #d1d1d6;background:#fff;color:#007aff;font-size:14.5px;font-weight:500;cursor:pointer;box-shadow:0 1px 2px rgba(0,0,0,0.05);display:inline-flex;align-items:center;justify-content:center;';
        const disabledStyle = 'opacity:0.36;cursor:default;';
        let html = '';
        html += `<button class="page-btn" data-page="1" ${currentPage===1?'disabled':''} style="${navStyle}${currentPage===1?disabledStyle:''}">«</button>`;
        html += `<button class="page-btn" data-page="${currentPage-1}" ${currentPage===1?'disabled':''} style="${navStyle}${currentPage===1?disabledStyle:''}">‹</button>`;

        const pages = [];
        if (totalPages <= 7) {
            for (let i = 1; i <= totalPages; i++) pages.push(i);
        } else {
            pages.push(1);
            let start = Math.max(2, currentPage - 1);
            let end = Math.min(totalPages - 1, currentPage + 1);
            if (currentPage <= 3) { start = 2; end = 4; }
            else if (currentPage >= totalPages - 2) { start = totalPages - 3; end = totalPages - 1; }
            if (start > 2) pages.push('...');
            for (let i = start; i <= end; i++) pages.push(i);
            if (end < totalPages - 1) pages.push('...');
            pages.push(totalPages);
        }

        pages.forEach(p => {
            if (p === '...') {
                html += `<span style="min-width:26px;height:34px;display:inline-flex;align-items:center;justify-content:center;color:#8e8e93;font-size:13.5px;user-select:none;">…</span>`;
            } else if (p === currentPage) {
                html += `<button class="page-num active" data-page="${p}" style="${activeStyle}">${p}</button>`;
            } else {
                html += `<button class="page-num" data-page="${p}" style="${baseStyle}">${p}</button>`;
            }
        });

        html += `<button class="page-btn" data-page="${currentPage+1}" ${currentPage===totalPages?'disabled':''} style="${navStyle}${currentPage===totalPages?disabledStyle:''}">›</button>`;
        html += `<button class="page-btn" data-page="${totalPages}" ${currentPage===totalPages?'disabled':''} style="${navStyle}${currentPage===totalPages?disabledStyle:''}">»</button>`;
        pagination.innerHTML = html;
    }

    document.addEventListener('click', function(e) {
        const target = e.target.closest('button');
        if (!target) return;
        if (target.classList.contains('page-btn') || target.classList.contains('page-num')) {
            const page = parseInt(target.dataset.page);
            if (page && page !== currentPage && page >= 1 && page <= totalPages) {
                currentPage = page;
                loadFiles(page);
            }
            return;
        }
        if (target.classList.contains('play-local')) {
            const item = target.closest('.file-item');
            if (item) {
                const path = item.dataset.path;
                if (path) window.open(`/api/play?path=${encodeURIComponent(path)}`, '_blank');
            }
            return;
        }
        if (target.classList.contains('play-web')) {
            const item = target.closest('.file-item');
            if (item) {
                const jump = item.dataset.jump;
                const name = item.dataset.name;
                if (jump && jump.startsWith('http')) window.open(jump, '_blank');
                else window.open(`https://www.bing.com/s?wd=${encodeURIComponent(name.replace(/\\s+/g, '+'))}`, '_blank');
            }
            return;
        }
    });

    document.getElementById('scanBtn').addEventListener('click', function() {
        fetch('/api/select_folder').then(r => r.json()).then(data => {
            if (data.error) { alert('无法打开文件夹选择器: ' + data.error); return; }
            const folder = data.folder;
            if (!folder) return;
            scanFolder = folder;
            listContainer.innerHTML = `<div class="empty scanning"><div class="spinner"></div><div style="margin-top: 11px; color: #8e8e93; font-size: 14.5px; font-weight: 450;">正在扫描…</div></div>`;
            return fetch('/api/scan', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({folder: folder}) });
        }).then(response => response ? response.json() : null).then(scanData => {
            if (!scanData) return;
            if (scanData.error) {
                alert('扫描失败: ' + scanData.error);
                listContainer.innerHTML = '<div class="empty">扫描出错，请重试</div>';
            } else {
                currentPage = 1;
                loadFiles(1);
            }
        }).catch(e => {
            alert('请求失败: ' + e.message);
            listContainer.innerHTML = '<div class="empty">请求失败</div>';
        });
    });

    let searchTimer = null;
    searchInput.addEventListener('input', function() {
        clearTimeout(searchTimer);
        searchTimer = setTimeout(() => { currentPage = 1; loadFiles(1); }, 280);
    });

    document.getElementById('exportBtn').addEventListener('click', function() {
        if (!scanFolder) { alert('请先选择文件夹'); return; }
        window.open('/api/export', '_blank');
    });

    document.getElementById('subfolderSwitch').addEventListener('change', function() {
        const sub = this.checked;
        const browser = document.getElementById('browserSwitch').checked;
        fetch(`/api/config?scan_subfolder=${sub}&use_browser=${browser}`);
    });
    document.getElementById('browserSwitch').addEventListener('change', function() {
        const browser = this.checked;
        const sub = document.getElementById('subfolderSwitch').checked;
        fetch(`/api/config?scan_subfolder=${sub}&use_browser=${browser}`);
    });

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    let resizeTimer = null;
    window.addEventListener('resize', function() {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(() => { if (scanFolder) loadFiles(currentPage); }, 180);
    });

    fetch('/api/config').then(r => r.json()).then(data => {
        document.getElementById('subfolderSwitch').checked = data.scan_subfolder;
        document.getElementById('browserSwitch').checked = data.use_browser;
        if (data.scan_subfolder) document.getElementById('subfolderToggle').classList.add('active');
        if (data.use_browser) document.getElementById('browserToggle').classList.add('active');
    });
</script>
</body>
</html>"""

# ========== 启动服务器 ==========
def start_server(port=8000):
    actual_port = port
    while True:
        try:
            with ThreadingHTTPServer(("", actual_port), MyHandler) as httpd:
                print(f"服务已启动 → http://localhost:{actual_port}")
                webbrowser.open(f"http://localhost:{actual_port}")
                httpd.serve_forever()
                break
        except OSError as e:
            if e.errno == 10048:
                actual_port += 1
                print(f"端口 {actual_port-1} 被占用，尝试 {actual_port}")
            else:
                raise

if __name__ == "__main__":
    start_server()