"""
飞书云文档工具
- 文件上传：上传各类文件到飞书云空间
- 多维表格：自动识别列标题，支持手动填写、Excel/CSV 导入、文件上传回填
"""

import csv
import json
import os
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import requests

# ============================================================
# 配置管理
# ============================================================

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(config: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ============================================================
# 飞书 API 封装
# ============================================================

BASE_URL = "https://open.feishu.cn/open-apis"

EXTENSION_TO_TYPE = {
    ".doc": "doc", ".docx": "doc", ".pdf": "pdf",
    ".xls": "xls", ".xlsx": "xls",
    ".ppt": "ppt", ".pptx": "ppt",
    ".txt": "doc", ".csv": "xls", ".md": "doc",
    ".png": "image", ".jpg": "image", ".jpeg": "image",
    ".gif": "image", ".bmp": "image", ".webp": "image", ".svg": "image",
    ".mp4": "video", ".avi": "video", ".mov": "video",
    ".mkv": "video", ".wmv": "video", ".flv": "video",
    ".mp3": "audio", ".wav": "audio", ".flac": "audio",
    ".aac": "audio", ".ogg": "audio", ".wma": "audio",
    ".zip": "file", ".rar": "file", ".7z": "file",
    ".tar": "file", ".gz": "file",
}


def get_file_type(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    return EXTENSION_TO_TYPE.get(ext, "file")


def parse_bitable_url(url: str) -> tuple[str, str, str]:
    """
    解析飞书多维表格 URL，返回 (app_token, table_id, view_id)。
    支持格式：
      https://xxx.feishu.cn/base/APPTOKEN?table=TABLEID&view=VIEWID
      https://xxx.feishu.cn/base/APPTOKEN/TABLEID/VIEWID
    """
    # 先取 app_token（/base/ 后的第一段路径）
    m = re.search(r'/base/([A-Za-z0-9_-]+)', url)
    if not m:
        raise ValueError("无法识别多维表格链接，请检查 URL 格式")
    app_token = m.group(1)

    # table_id: 查询参数或路径
    table_id = ""
    view_id = ""
    t = re.search(r'[?&]table=([A-Za-z0-9_-]+)', url)
    if t:
        table_id = t.group(1)
    v = re.search(r'[?&]view=([A-Za-z0-9_-]+)', url)
    if v:
        view_id = v.group(1)

    # 路径形式 fallback
    if not table_id:
        parts = re.findall(r'/base/[^?]+', url)
        if parts:
            segs = parts[0].strip('/').split('/')
            # segs: ['base', app_token, table_id?, view_id?]
            if len(segs) >= 3:
                table_id = segs[2]
            if len(segs) >= 4:
                view_id = segs[3]

    return app_token, table_id, view_id


# 飞书多维表格字段类型常量
FIELD_TYPE_NAMES = {
    1: "文本", 2: "数字", 3: "单选", 4: "多选", 5: "日期",
    7: "复选框", 11: "人员", 13: "电话", 15: "超链接",
    17: "附件", 18: "单向关联", 19: "查找引用", 20: "公式",
    21: "双向关联", 22: "地理位置", 23: "群组", 1001: "创建时间",
    1002: "更新时间", 1003: "创建人", 1004: "更新人", 1005: "自动编号",
}


class FeishuClient:
    """飞书 API 客户端"""

    def __init__(self, app_id: str, app_secret: str):
        self.app_id = app_id
        self.app_secret = app_secret
        self._tenant_token: str | None = None

    # ---------- 认证 ----------

    def get_tenant_access_token(self) -> str:
        url = f"{BASE_URL}/auth/v3/tenant_access_token/internal"
        resp = requests.post(url, json={
            "app_id": self.app_id,
            "app_secret": self.app_secret,
        }, timeout=10)
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"获取 token 失败: {data.get('msg', '未知错误')}")
        self._tenant_token = data["tenant_access_token"]
        return self._tenant_token

    @property
    def token(self) -> str:
        if not self._tenant_token:
            self.get_tenant_access_token()
        return self._tenant_token

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self.token}"}

    def _request(self, method, url, **kwargs):
        """带自动 token 刷新的请求"""
        kwargs.setdefault("timeout", 30)
        # 合并 headers：调用方的 headers 优先，自动注入 Authorization
        extra_headers = kwargs.pop("headers", {})
        merged = {**self._headers(), **extra_headers}
        resp = requests.request(method, url, headers=merged, **kwargs)
        data = resp.json()
        # token 过期自动刷新重试
        if data.get("code") in (99991663, 99991661):
            self.get_tenant_access_token()
            merged = {**self._headers(), **extra_headers}
            resp = requests.request(method, url, headers=merged, **kwargs)
            data = resp.json()
        return data

    # ---------- 云空间文件上传 ----------

    def upload_file(self, file_path: str, folder_token: str,
                    progress_callback=None) -> dict:
        file_size = os.path.getsize(file_path)
        file_name = os.path.basename(file_path)

        if file_size <= 20 * 1024 * 1024:
            return self._upload_small(file_path, file_name, file_size,
                                      folder_token, progress_callback)
        else:
            return self._upload_large(file_path, file_name, file_size,
                                      folder_token, progress_callback)

    def _upload_small(self, file_path, file_name, file_size,
                      folder_token, progress_callback) -> dict:
        url = f"{BASE_URL}/drive/v1/files/upload_all"
        with open(file_path, "rb") as f:
            file_data = f.read()
        if progress_callback:
            progress_callback(0.3)
        resp = requests.post(
            url, headers=self._headers(),
            data={
                "file_name": file_name, "parent_type": "explorer",
                "parent_node": folder_token, "size": str(file_size),
            },
            files={"file": (file_name, file_data)},
            timeout=300,
        )
        if progress_callback:
            progress_callback(1.0)
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"上传失败: {data.get('msg', '未知错误')}")
        return data.get("data", {})

    def _upload_large(self, file_path, file_name, file_size,
                      folder_token, progress_callback) -> dict:
        block_size = 4 * 1024 * 1024
        block_num = (file_size + block_size - 1) // block_size

        url_prepare = f"{BASE_URL}/drive/v1/files/upload_prepare"
        resp = requests.post(
            url_prepare,
            headers={**self._headers(), "Content-Type": "application/json"},
            json={
                "file_name": file_name, "parent_type": "explorer",
                "parent_node": folder_token, "size": file_size,
            },
            timeout=30,
        )
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"预上传失败: {data.get('msg')}")
        upload_id = data["data"]["upload_id"]

        import time as _time
        url_part = f"{BASE_URL}/drive/v1/files/upload_part"
        max_retries = 3
        with open(file_path, "rb") as f:
            for seq in range(block_num):
                chunk = f.read(block_size)
                last_err = None
                for attempt in range(max_retries):
                    try:
                        resp = requests.post(
                            url_part, headers=self._headers(),
                            data={"upload_id": upload_id, "seq": str(seq),
                                  "size": str(len(chunk))},
                            files={"file": (file_name, chunk)}, timeout=300,
                        )
                        part_data = resp.json()
                        if part_data.get("code") != 0:
                            raise RuntimeError(part_data.get('msg'))
                        last_err = None
                        break
                    except Exception as e:
                        last_err = e
                        if attempt < max_retries - 1:
                            _time.sleep(3 * (attempt + 1))
                if last_err:
                    raise RuntimeError(f"分片{seq}失败(重试{max_retries}次): {last_err}")
                if progress_callback:
                    progress_callback((seq + 1) / (block_num + 1))

        url_finish = f"{BASE_URL}/drive/v1/files/upload_finish"
        resp = requests.post(
            url_finish,
            headers={**self._headers(), "Content-Type": "application/json"},
            json={"upload_id": upload_id, "block_num": block_num},
            timeout=30,
        )
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"完成上传失败: {data.get('msg')}")
        if progress_callback:
            progress_callback(1.0)
        return data.get("data", {})

    # ---------- 多维表格 Bitable API ----------

    def bitable_list_fields(self, app_token: str, table_id: str) -> list[dict]:
        """获取多维表格的字段列表"""
        url = f"{BASE_URL}/bitable/v1/apps/{app_token}/tables/{table_id}/fields"
        all_fields = []
        page_token = None
        while True:
            params = {"page_size": 100}
            if page_token:
                params["page_token"] = page_token
            data = self._request("GET", url, params=params)
            if data.get("code") != 0:
                raise RuntimeError(f"获取字段失败: {data.get('msg')}")
            items = data.get("data", {}).get("items", [])
            all_fields.extend(items)
            if not data.get("data", {}).get("has_more"):
                break
            page_token = data["data"].get("page_token")
        return all_fields

    def bitable_list_records(self, app_token: str, table_id: str,
                             page_size: int = 100) -> list[dict]:
        """获取多维表格中的记录（最多返回 page_size 条）"""
        url = f"{BASE_URL}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
        data = self._request("GET", url, params={"page_size": page_size})
        if data.get("code") != 0:
            raise RuntimeError(f"获取记录失败: {data.get('msg')}")
        return data.get("data", {}).get("items", [])

    def bitable_list_all_records(self, app_token: str, table_id: str) -> list[dict]:
        """获取多维表格中的所有记录（自动翻页）"""
        url = f"{BASE_URL}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
        all_records = []
        page_token = None
        while True:
            params = {"page_size": 500}
            if page_token:
                params["page_token"] = page_token
            data = self._request("GET", url, params=params)
            if data.get("code") != 0:
                raise RuntimeError(f"获取记录失败: {data.get('msg')}")
            items = data.get("data", {}).get("items", [])
            all_records.extend(items)
            if not data.get("data", {}).get("has_more"):
                break
            page_token = data["data"].get("page_token")
        return all_records

    def bitable_update_record(self, app_token: str, table_id: str,
                              record_id: str, fields: dict) -> dict:
        """更新单条记录"""
        url = f"{BASE_URL}/bitable/v1/apps/{app_token}/tables/{table_id}/records/{record_id}"
        data = self._request(
            "PUT", url,
            json={"fields": fields},
            headers={**self._headers(), "Content-Type": "application/json"},
        )
        if data.get("code") != 0:
            raise RuntimeError(f"更新记录失败: {data.get('msg')}")
        return data.get("data", {})

    def bitable_create_records(self, app_token: str, table_id: str,
                               records: list[dict]) -> dict:
        """批量创建记录。records 为 [{"fields": {"列名": "值", ...}}, ...]"""
        url = f"{BASE_URL}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_create"
        data = self._request(
            "POST", url,
            json={"records": records},
            headers={**self._headers(), "Content-Type": "application/json"},
        )
        if data.get("code") != 0:
            raise RuntimeError(f"创建记录失败: {data.get('msg')}")
        return data.get("data", {})

    def bitable_upload_media(self, file_path: str, parent_token: str) -> str:
        """
        上传附件到多维表格，返回 file_token。
        <=20MB 用 upload_all，>20MB 用分片上传。
        """
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)
        ext = Path(file_path).suffix.lower()

        image_exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
        parent_type = "bitable_image" if ext in image_exts else "bitable_file"

        if file_size <= 20 * 1024 * 1024:
            return self._media_upload_small(file_path, file_name, file_size,
                                            parent_type, parent_token)
        else:
            return self._media_upload_large(file_path, file_name, file_size,
                                            parent_type, parent_token)

    def _media_upload_small(self, file_path, file_name, file_size,
                            parent_type, parent_token) -> str:
        """小文件一次性上传 (<=20MB)，带重试"""
        import time as _time
        url = f"{BASE_URL}/drive/v1/medias/upload_all"
        max_retries = 3
        last_err = None
        for attempt in range(max_retries):
            try:
                with open(file_path, "rb") as f:
                    resp = requests.post(
                        url, headers=self._headers(),
                        data={
                            "file_name": file_name,
                            "parent_type": parent_type,
                            "parent_node": parent_token,
                            "size": str(file_size),
                        },
                        files={"file": (file_name, f)},
                        timeout=300,
                    )
                data = resp.json()
                if data.get("code") != 0:
                    raise RuntimeError(
                        f"{data.get('msg')} (code={data.get('code')}, "
                        f"type={parent_type}, size={file_size})")
                return data["data"]["file_token"]
            except Exception as e:
                last_err = e
                if attempt < max_retries - 1:
                    _time.sleep(3 * (attempt + 1))
        raise last_err

    def _media_upload_large(self, file_path, file_name, file_size,
                            parent_type, parent_token) -> str:
        """大文件分片上传 (>20MB)"""
        block_size = 4 * 1024 * 1024  # 4MB
        block_num = (file_size + block_size - 1) // block_size

        # 1. 预上传
        url_prepare = f"{BASE_URL}/drive/v1/medias/upload_prepare"
        data = self._request(
            "POST", url_prepare,
            headers={"Content-Type": "application/json"},
            json={
                "file_name": file_name,
                "parent_type": parent_type,
                "parent_node": parent_token,
                "size": file_size,
            },
        )
        if data.get("code") != 0:
            raise RuntimeError(
                f"预上传失败: {data.get('msg')} (code={data.get('code')})")
        upload_id = data["data"]["upload_id"]

        # 2. 分片上传（带重试）
        import time as _time
        url_part = f"{BASE_URL}/drive/v1/medias/upload_part"
        max_retries = 3
        with open(file_path, "rb") as f:
            for seq in range(block_num):
                chunk = f.read(block_size)
                last_err = None
                for attempt in range(max_retries):
                    try:
                        resp = requests.post(
                            url_part, headers=self._headers(),
                            data={
                                "upload_id": upload_id,
                                "seq": str(seq),
                                "size": str(len(chunk)),
                            },
                            files={"file": (file_name, chunk)},
                            timeout=300,
                        )
                        part_data = resp.json()
                        if part_data.get("code") != 0:
                            raise RuntimeError(part_data.get('msg'))
                        last_err = None
                        break
                    except Exception as e:
                        last_err = e
                        if attempt < max_retries - 1:
                            _time.sleep(3 * (attempt + 1))
                if last_err:
                    raise RuntimeError(
                        f"分片{seq}/{block_num}失败(重试{max_retries}次): {last_err}")

        # 3. 完成上传
        url_finish = f"{BASE_URL}/drive/v1/medias/upload_finish"
        data = self._request(
            "POST", url_finish,
            headers={"Content-Type": "application/json"},
            json={
                "upload_id": upload_id,
                "block_num": block_num,
            },
        )
        if data.get("code") != 0:
            raise RuntimeError(
                f"完成上传失败: {data.get('msg')} (code={data.get('code')})")
        return data["data"]["file_token"]


# ============================================================
# GUI 界面
# ============================================================

class FeishuToolApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("飞书云文档工具")
        self.root.geometry("860x820")
        self.root.resizable(True, True)

        self.client: FeishuClient | None = None
        self.file_list: list[str] = []
        self.uploading = False

        # 多维表格状态
        self.bitable_fields: list[dict] = []
        self.bitable_app_token = ""
        self.bitable_table_id = ""
        self._existing_records: list[dict] = []

        self.config = load_config()
        self._build_ui()

    # ================================================================
    # UI 构建
    # ================================================================

    def _build_ui(self):
        # 顶部凭证区域（两个选项卡共享）
        top = ttk.Frame(self.root, padding=(12, 12, 12, 0))
        top.pack(fill=tk.X)
        self._build_credential_section(top)

        # 选项卡
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=(4, 12))

        tab_upload = ttk.Frame(self.notebook, padding=8)
        tab_bitable = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(tab_upload, text=" 文件上传 ")
        self.notebook.add(tab_bitable, text=" 多维表格 ")

        self._build_upload_tab(tab_upload)
        self._build_bitable_tab(tab_bitable)

        # 底部状态栏
        bottom = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        bottom.pack(fill=tk.X)
        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(bottom, variable=self.progress_var, maximum=100).pack(
            fill=tk.X, pady=(0, 4))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom, textvariable=self.status_var).pack(side=tk.LEFT)

    def _build_credential_section(self, parent):
        cred = ttk.LabelFrame(parent, text="飞书应用凭证", padding=8)
        cred.pack(fill=tk.X)

        ttk.Label(cred, text="App ID:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.app_id_var = tk.StringVar(value=self.config.get("app_id", ""))
        ttk.Entry(cred, textvariable=self.app_id_var, width=50).grid(
            row=0, column=1, sticky=tk.EW, padx=(4, 0), pady=2)

        ttk.Label(cred, text="App Secret:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.app_secret_var = tk.StringVar(value=self.config.get("app_secret", ""))
        ttk.Entry(cred, textvariable=self.app_secret_var, width=50, show="*").grid(
            row=1, column=1, sticky=tk.EW, padx=(4, 0), pady=2)

        btn_f = ttk.Frame(cred)
        btn_f.grid(row=0, column=2, rowspan=2, padx=(8, 0))
        ttk.Button(btn_f, text="保存凭证", command=self._save_credentials).pack(pady=2)
        ttk.Button(btn_f, text="测试连接", command=self._test_connection).pack(pady=2)
        cred.columnconfigure(1, weight=1)

    # ---- 文件上传选项卡 ----

    def _build_upload_tab(self, parent):
        # 目标文件夹
        ff = ttk.LabelFrame(parent, text="目标文件夹", padding=8)
        ff.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(ff, text="文件夹 Token:").grid(row=0, column=0, sticky=tk.W)
        self.folder_token_var = tk.StringVar(value=self.config.get("folder_token", ""))
        ttk.Entry(ff, textvariable=self.folder_token_var, width=50).grid(
            row=0, column=1, sticky=tk.EW, padx=(4, 0))
        ttk.Button(ff, text="获取根目录", command=self._fetch_root_folder).grid(
            row=0, column=2, padx=(8, 0))
        ff.columnconfigure(1, weight=1)
        ttk.Label(ff, text="提示：飞书云文档文件夹 URL 最后一段即为 Token",
                  foreground="gray").grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(4, 0))

        # 文件列表
        fl = ttk.LabelFrame(parent, text="待上传文件", padding=8)
        fl.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        btn_row = ttk.Frame(fl)
        btn_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(btn_row, text="添加文件", command=self._add_files).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text="添加文件夹", command=self._add_folder).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text="清空列表", command=self._clear_files).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text="移除选中", command=self._remove_selected).pack(side=tk.LEFT)

        cols = ("name", "size", "type", "status")
        self.tree = ttk.Treeview(fl, columns=cols, show="headings", height=8)
        self.tree.heading("name", text="文件名")
        self.tree.heading("size", text="大小")
        self.tree.heading("type", text="类型")
        self.tree.heading("status", text="状态")
        self.tree.column("name", width=280)
        self.tree.column("size", width=80, anchor=tk.E)
        self.tree.column("type", width=60, anchor=tk.CENTER)
        self.tree.column("status", width=100, anchor=tk.CENTER)
        sb = ttk.Scrollbar(fl, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.upload_btn = ttk.Button(parent, text="开始上传", command=self._start_upload)
        self.upload_btn.pack(anchor=tk.E)

    # ---- 多维表格选项卡 ----

    def _build_bitable_tab(self, parent):
        # 顶部：链接 + 填充范围（紧凑横排）
        top = ttk.Frame(parent)
        top.pack(fill=tk.X, pady=(0, 4))

        # URL 行
        ttk.Label(top, text="表格链接:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.bitable_url_var = tk.StringVar(value=self.config.get("bitable_url", ""))
        ttk.Entry(top, textvariable=self.bitable_url_var, width=50).grid(
            row=0, column=1, sticky=tk.EW, padx=(4, 4), pady=2)
        ttk.Button(top, text="解析并获取列", command=self._parse_bitable_url).grid(
            row=0, column=2, padx=(0, 4), pady=2)
        self.record_count_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.record_count_var,
                  foreground="blue").grid(row=0, column=3, padx=(4, 0), pady=2)

        # 起始行 + 写入模式 行
        ttk.Label(top, text="起始行:").grid(row=1, column=0, sticky=tk.W, pady=2)
        opt_frame = ttk.Frame(top)
        opt_frame.grid(row=1, column=1, columnspan=3, sticky=tk.W, pady=2)
        self.start_row_var = tk.StringVar(value="1")
        ttk.Spinbox(opt_frame, from_=1, to=99999, width=5,
                     textvariable=self.start_row_var).pack(side=tk.LEFT, padx=(0, 8))
        self.write_mode_var = tk.StringVar(value="append")
        ttk.Radiobutton(opt_frame, text="追加", variable=self.write_mode_var,
                        value="append").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Radiobutton(opt_frame, text="更新已有行", variable=self.write_mode_var,
                        value="update").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(opt_frame, text="查询记录数",
                   command=self._query_record_count).pack(side=tk.LEFT, padx=(8, 0))

        top.columnconfigure(1, weight=1)

        # 底部：提交按钮（先 pack，确保始终可见）
        bottom = ttk.Frame(parent)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))
        self.bitable_submit_btn = ttk.Button(
            bottom, text="  提交到多维表格  ", command=self._submit_bitable)
        self.bitable_submit_btn.pack(side=tk.RIGHT)

        # 中间：三种模式用子选项卡（后 pack，填充剩余空间）
        self.fill_notebook = ttk.Notebook(parent)
        self.fill_notebook.pack(fill=tk.BOTH, expand=True)

        tab_manual = ttk.Frame(self.fill_notebook, padding=4)
        tab_import = ttk.Frame(self.fill_notebook, padding=4)
        tab_file = ttk.Frame(self.fill_notebook, padding=4)
        self.fill_notebook.add(tab_manual, text=" 手动填写 ")
        self.fill_notebook.add(tab_import, text=" Excel/CSV 导入 ")
        self.fill_notebook.add(tab_file, text=" 文件夹上传 ")

        self._build_manual_panel(tab_manual)
        self._build_import_panel(tab_import)
        self._build_file_upload_panel(tab_file)

    def _build_manual_panel(self, parent=None):
        container = parent
        self.manual_frame = ttk.Frame(container)
        self.manual_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(self.manual_frame,
                  text="请先解析链接获取列信息，然后在下方表格中填写数据。",
                  foreground="gray").pack(anchor=tk.W, pady=(0, 4))

        # 手动输入表格（动态生成）
        self.manual_table_frame = ttk.Frame(self.manual_frame)
        self.manual_table_frame.pack(fill=tk.BOTH, expand=True)

        btn_row = ttk.Frame(self.manual_frame)
        btn_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(btn_row, text="添加一行", command=self._manual_add_row).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text="清空所有行", command=self._manual_clear_rows).pack(side=tk.LEFT)

        self.manual_entries: list[list[ttk.Entry]] = []
        self.manual_canvas = None

    def _build_import_panel(self, parent=None):
        container = parent
        self.import_frame = ttk.Frame(container)
        self.import_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择
        file_row = ttk.Frame(self.import_frame)
        file_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(file_row, text="选择文件:").pack(side=tk.LEFT)
        self.import_path_var = tk.StringVar()
        ttk.Entry(file_row, textvariable=self.import_path_var, width=45).pack(
            side=tk.LEFT, padx=(4, 4), fill=tk.X, expand=True)
        ttk.Button(file_row, text="浏览", command=self._browse_import_file).pack(
            side=tk.LEFT, padx=(0, 4))
        ttk.Button(file_row, text="预览并识别映射", command=self._preview_import).pack(side=tk.LEFT)

        ttk.Label(self.import_frame,
                  text="支持 .csv 和 .xlsx 文件。首行为列标题，预览后自动生成映射表。",
                  foreground="gray").pack(anchor=tk.W, pady=(0, 4))

        # 上下拆分：上方预览数据，下方映射表
        paned = ttk.PanedWindow(self.import_frame, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 预览 Treeview
        preview_frame = ttk.LabelFrame(paned, text="数据预览", padding=4)
        self.import_preview_tree = ttk.Treeview(preview_frame, show="headings", height=4)
        sb3 = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL,
                            command=self.import_preview_tree.yview)
        self.import_preview_tree.configure(yscrollcommand=sb3.set)
        self.import_preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb3.pack(side=tk.RIGHT, fill=tk.Y)
        paned.add(preview_frame, weight=1)

        # 列映射表
        map_frame = ttk.LabelFrame(paned, text="列映射（源文件列 → 表格列）", padding=4)
        ttk.Label(map_frame,
                  text="预览数据后自动匹配同名列，你可以通过下拉框调整映射关系。选「不映射」则跳过该列。",
                  foreground="gray").pack(anchor=tk.W, pady=(0, 4))
        self.import_mapping_frame = ttk.Frame(map_frame)
        self.import_mapping_frame.pack(fill=tk.BOTH, expand=True)
        paned.add(map_frame, weight=1)

        self.import_data: list[dict] = []
        # import_mapping_combos: list of (source_col_name, Combobox) — 源列对应目标列的下拉框
        self.import_mapping_combos: list[tuple[str, ttk.Combobox]] = []

    def _build_file_upload_panel(self, parent=None):
        container = parent
        self.file_upload_frame = ttk.Frame(container)
        self.file_upload_frame.pack(fill=tk.BOTH, expand=True)

        # 步骤 1：选择父文件夹 + 排序方式
        step1 = ttk.Frame(self.file_upload_frame)
        step1.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(step1, text="选择文件夹", command=self._bt_select_folder).pack(
            side=tk.LEFT, padx=(0, 4))
        ttk.Button(step1, text="清空", command=self._bt_clear).pack(
            side=tk.LEFT, padx=(0, 8))

        ttk.Label(step1, text="排序:").pack(side=tk.LEFT, padx=(4, 2))
        self.bt_sort_var = tk.StringVar(value="名称_后数字")
        sort_combo = ttk.Combobox(step1, textvariable=self.bt_sort_var, width=14,
                                  state="readonly", values=[
                                      "名称_后数字",
                                      "名称自然排序",
                                      "修改时间↑",
                                      "修改时间↓",
                                      "原始顺序",
                                  ])
        sort_combo.current(0)
        sort_combo.pack(side=tk.LEFT, padx=(0, 4))
        sort_combo.bind("<<ComboboxSelected>>", lambda e: self._bt_resort_folders())

        self.bt_status_var = tk.StringVar(
            value="选择包含多个子文件夹的父目录，每个子文件夹 = 一行数据")
        ttk.Label(step1, textvariable=self.bt_status_var,
                  foreground="gray").pack(side=tk.LEFT)

        # 子文件夹状态
        self.bt_subfolders: list[str] = []  # 子文件夹路径列表（已排序、已过滤）
        self.bt_all_valid_dirs: list[str] = []  # 过滤后的全部有效子文件夹（排序前）
        self.bt_skipped: list[tuple[str, str]] = []  # 被排除的子文件夹
        self.bt_sample_files: list[str] = []  # 第一个子文件夹的文件列表

        # 步骤 2：映射规则（用第一个子文件夹识别）
        map_frame = ttk.LabelFrame(self.file_upload_frame,
                                   text="映射规则（基于第一个子文件夹自动识别）", padding=6)
        map_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        self.bt_mapping_inner = ttk.Frame(map_frame)
        self.bt_mapping_inner.pack(fill=tk.BOTH, expand=True)

        # 步骤 3：子文件夹预览列表
        preview_frame = ttk.LabelFrame(self.file_upload_frame,
                                       text="子文件夹列表（每个 = 一行数据）", padding=4)
        preview_frame.pack(fill=tk.X, pady=(0, 4))

        self.bt_folder_tree = ttk.Treeview(preview_frame,
                                           columns=("name", "match", "status"),
                                           show="headings", height=4)
        self.bt_folder_tree.heading("name", text="子文件夹名")
        self.bt_folder_tree.heading("match", text="匹配")
        self.bt_folder_tree.heading("status", text="状态")
        self.bt_folder_tree.column("name", width=200)
        self.bt_folder_tree.column("match", width=60, anchor=tk.CENTER)
        self.bt_folder_tree.column("status", width=120, anchor=tk.CENTER)
        bt_sb = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL,
                              command=self.bt_folder_tree.yview)
        self.bt_folder_tree.configure(yscrollcommand=bt_sb.set)
        self.bt_folder_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bt_sb.pack(side=tk.RIGHT, fill=tk.Y)

        # 日志文件路径
        self.bt_log_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "upload_log.txt")

        # 映射配置：每个表格列对应什么文件名
        # bt_col_mapping: [{col_name, col_type, widget_type, widget, entry?}]
        self.bt_col_mapping: list[dict] = []

    def _get_active_fill_mode(self) -> str:
        """获取当前选中的填充模式"""
        idx = self.fill_notebook.index(self.fill_notebook.select())
        return ["manual", "import", "file_upload"][idx]

    # ================================================================
    # 凭证操作
    # ================================================================

    def _save_credentials(self):
        config = load_config()
        config["app_id"] = self.app_id_var.get().strip()
        config["app_secret"] = self.app_secret_var.get().strip()
        config["folder_token"] = self.folder_token_var.get().strip()
        config["bitable_url"] = self.bitable_url_var.get().strip()
        save_config(config)
        self.status_var.set("凭证已保存")

    def _test_connection(self):
        app_id = self.app_id_var.get().strip()
        app_secret = self.app_secret_var.get().strip()
        if not app_id or not app_secret:
            messagebox.showwarning("提示", "请先填写 App ID 和 App Secret")
            return
        try:
            client = FeishuClient(app_id, app_secret)
            token = client.get_tenant_access_token()
            self.client = client
            self.status_var.set("连接成功")
            messagebox.showinfo("成功", f"连接成功！\ntoken: {token[:16]}...")
        except Exception as e:
            messagebox.showerror("连接失败", str(e))

    def _ensure_client(self):
        if self.client:
            return True
        app_id = self.app_id_var.get().strip()
        app_secret = self.app_secret_var.get().strip()
        if not app_id or not app_secret:
            messagebox.showwarning("提示", "请先填写 App ID 和 App Secret")
            return False
        self.client = FeishuClient(app_id, app_secret)
        return True

    # ================================================================
    # 文件上传选项卡操作
    # ================================================================

    def _fetch_root_folder(self):
        if not self._ensure_client():
            return
        try:
            url = f"{BASE_URL}/drive/explorer/v2/root_folder/meta"
            resp = requests.get(url, headers=self.client._headers(), timeout=10)
            data = resp.json()
            if data.get("code") != 0:
                raise RuntimeError(data.get("msg"))
            root_token = data["data"]["token"]
            self.folder_token_var.set(root_token)
            self.status_var.set(f"根目录 Token: {root_token}")
        except Exception as e:
            messagebox.showerror("错误", f"获取根目录失败: {e}")

    def _add_files(self):
        paths = filedialog.askopenfilenames(title="选择文件")
        for p in paths:
            if p not in self.file_list:
                self.file_list.append(p)
                self._insert_file_row(p)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if not folder:
            return
        for root_dir, _, files in os.walk(folder):
            for fname in files:
                full = os.path.join(root_dir, fname)
                if full not in self.file_list:
                    self.file_list.append(full)
                    self._insert_file_row(full)

    def _insert_file_row(self, filepath: str):
        name = os.path.basename(filepath)
        size = os.path.getsize(filepath)
        ftype = get_file_type(name)
        self.tree.insert("", tk.END, iid=filepath, values=(
            name, self._format_size(size), ftype, "等待上传"))

    def _clear_files(self):
        self.file_list.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _remove_selected(self):
        for item in self.tree.selection():
            if item in self.file_list:
                self.file_list.remove(item)
            self.tree.delete(item)

    @staticmethod
    def _format_size(size_bytes) -> str:
        size = float(size_bytes)
        for unit in ("B", "KB", "MB", "GB"):
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"

    def _start_upload(self):
        if self.uploading:
            return
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加要上传的文件")
            return
        folder_token = self.folder_token_var.get().strip()
        if not folder_token:
            messagebox.showwarning("提示", "请填写目标文件夹 Token")
            return
        if not self._ensure_client():
            return

        config = load_config()
        config["folder_token"] = folder_token
        save_config(config)

        self.uploading = True
        self.upload_btn.configure(state=tk.DISABLED)
        threading.Thread(target=self._upload_thread, args=(folder_token,), daemon=True).start()

    def _upload_thread(self, folder_token: str):
        total = len(self.file_list)
        success = failed = 0
        try:
            self.client.get_tenant_access_token()
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("认证失败", str(e)))
            self.root.after(0, self._upload_finished)
            return

        for idx, filepath in enumerate(list(self.file_list)):
            self.root.after(0, lambda i=idx, t=total: self.status_var.set(
                f"正在上传 {i + 1}/{t}..."))
            self.root.after(0, lambda fp=filepath: self._update_row(fp, "上传中..."))

            def progress_cb(pct, i=idx):
                self.root.after(0, lambda v=(i + pct) / total * 100: self.progress_var.set(v))

            try:
                self.client.upload_file(filepath, folder_token, progress_cb)
                self.root.after(0, lambda fp=filepath: self._update_row(fp, "成功 ✓"))
                success += 1
            except Exception as e:
                self.root.after(0, lambda fp=filepath, m=str(e)[:40]: self._update_row(
                    fp, f"失败: {m}"))
                failed += 1

        self.root.after(0, lambda: self.progress_var.set(100))
        self.root.after(0, lambda: self.status_var.set(
            f"上传完成 — 成功: {success}, 失败: {failed}"))
        self.root.after(0, lambda s=success, f=failed: messagebox.showinfo(
            "上传完成", f"成功: {s}\n失败: {f}"))
        self.root.after(0, self._upload_finished)

    def _update_row(self, filepath, status):
        try:
            self.tree.set(filepath, "status", status)
        except tk.TclError:
            pass

    def _upload_finished(self):
        self.uploading = False
        self.upload_btn.configure(state=tk.NORMAL)

    # ================================================================
    # 多维表格操作
    # ================================================================

    def _parse_bitable_url(self):
        url = self.bitable_url_var.get().strip()
        if not url:
            messagebox.showwarning("提示", "请先粘贴多维表格链接")
            return
        try:
            app_token, table_id, view_id = parse_bitable_url(url)
        except ValueError as e:
            messagebox.showerror("解析失败", str(e))
            return

        if not table_id:
            messagebox.showwarning("提示", "链接中未找到 table_id，请确保 URL 包含 ?table=xxx 参数")
            return

        self.bitable_app_token = app_token
        self.bitable_table_id = table_id
        self.status_var.set(f"解析成功 — app: {app_token[:10]}... table: {table_id}")

        # 保存 URL
        config = load_config()
        config["bitable_url"] = url
        save_config(config)

        # 获取字段
        if not self._ensure_client():
            return
        try:
            self.client.get_tenant_access_token()
            fields = self.client.bitable_list_fields(app_token, table_id)
            self.bitable_fields = fields
        except Exception as e:
            messagebox.showerror("获取字段失败", str(e))
            return

        # 显示解析结果
        writable_types = {19, 20, 1001, 1002, 1003, 1004, 1005}
        lines = []
        for f in fields:
            col_name = f.get("field_name", "")
            col_type = f.get("type", 0)
            type_name = FIELD_TYPE_NAMES.get(col_type, f"未知({col_type})")
            readonly = "（只读）" if col_type in writable_types else ""
            lines.append(f"  {col_name}  [{type_name}] {readonly}")

        msg = (f"解析成功！识别到 {len(fields)} 个列：\n\n"
               + "\n".join(lines)
               + f"\n\n表格 ID: {table_id}")
        messagebox.showinfo("解析结果", msg)
        self.status_var.set(f"已识别 {len(fields)} 个字段")

        # 查询现有记录数
        self._query_record_count()

        # 刷新手动填写表格和映射表
        self._rebuild_manual_table()
        self._rebuild_mapping_table()

    def _query_record_count(self):
        if not self.bitable_app_token or not self.bitable_table_id:
            messagebox.showwarning("提示", "请先解析多维表格链接")
            return
        if not self._ensure_client():
            return
        try:
            records = self.client.bitable_list_all_records(
                self.bitable_app_token, self.bitable_table_id)
            self._existing_records = records
            count = len(records)
            self.record_count_var.set(f"当前共 {count} 行记录")
            self.status_var.set(f"查询完成：表中现有 {count} 行记录")
        except Exception as e:
            self.record_count_var.set("查询失败")
            messagebox.showerror("查询失败", str(e))

    # ---- 手动填写 ----

    def _get_writable_fields(self) -> list[dict]:
        """过滤出可写入的字段（排除公式、自动编号等只读类型）"""
        readonly_types = {19, 20, 1001, 1002, 1003, 1004, 1005}
        return [f for f in self.bitable_fields if f.get("type") not in readonly_types]

    def _rebuild_manual_table(self):
        for w in self.manual_table_frame.winfo_children():
            w.destroy()
        self.manual_entries.clear()

        writable = self._get_writable_fields()
        if not writable:
            ttk.Label(self.manual_table_frame, text="没有可写入的字段",
                      foreground="gray").pack()
            return

        # 使用 Canvas + scrollbar 支持横向滚动
        canvas = tk.Canvas(self.manual_table_frame, height=150)
        h_scroll = ttk.Scrollbar(self.manual_table_frame, orient=tk.HORIZONTAL,
                                 command=canvas.xview)
        v_scroll = ttk.Scrollbar(self.manual_table_frame, orient=tk.VERTICAL,
                                 command=canvas.yview)
        canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        inner = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner, anchor=tk.NW)

        # 表头
        for col_idx, f in enumerate(writable):
            type_name = FIELD_TYPE_NAMES.get(f.get("type"), "")
            lbl = ttk.Label(inner, text=f"{f['field_name']}\n({type_name})",
                            anchor=tk.CENTER, width=16, relief=tk.GROOVE)
            lbl.grid(row=0, column=col_idx, sticky=tk.EW, padx=1, pady=1)

        # 默认添加一行
        self._manual_add_row_inner(inner, writable, 1)

        inner.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(fill=tk.BOTH, expand=True)
        self.manual_canvas = canvas
        self.manual_inner = inner
        self.manual_writable = writable

    def _manual_add_row_inner(self, parent, writable, row_num):
        row_entries = []
        for col_idx, f in enumerate(writable):
            e = ttk.Entry(parent, width=18)
            e.grid(row=row_num, column=col_idx, sticky=tk.EW, padx=1, pady=1)
            row_entries.append(e)
        self.manual_entries.append(row_entries)

    def _manual_add_row(self):
        if not hasattr(self, 'manual_inner') or not self.bitable_fields:
            messagebox.showwarning("提示", "请先解析多维表格链接")
            return
        writable = self._get_writable_fields()
        row_num = len(self.manual_entries) + 1
        self._manual_add_row_inner(self.manual_inner, writable, row_num)
        self.manual_inner.update_idletasks()
        if self.manual_canvas:
            self.manual_canvas.configure(scrollregion=self.manual_canvas.bbox("all"))

    def _manual_clear_rows(self):
        if hasattr(self, 'manual_inner'):
            self._rebuild_manual_table()

    # ---- Excel/CSV 导入 ----

    def _browse_import_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 或 CSV 文件",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")])
        if path:
            self.import_path_var.set(path)

    def _preview_import(self):
        path = self.import_path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择文件")
            return
        try:
            rows = self._read_tabular_file(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return

        if not rows:
            messagebox.showinfo("提示", "文件为空")
            return

        self.import_data = rows
        headers = list(rows[0].keys())

        # 刷新预览 Treeview
        self.import_preview_tree.delete(*self.import_preview_tree.get_children())
        self.import_preview_tree["columns"] = headers
        for h in headers:
            self.import_preview_tree.heading(h, text=h)
            self.import_preview_tree.column(h, width=100)

        for row in rows[:50]:  # 最多预览 50 行
            vals = [str(row.get(h, "")) for h in headers]
            self.import_preview_tree.insert("", tk.END, values=vals)

        # 构建映射表
        self._rebuild_import_mapping(headers)

        matched = sum(1 for _, cb in self.import_mapping_combos if cb.get() != "（不映射）")
        self.status_var.set(f"预览: {len(rows)} 行数据, {matched}/{len(headers)} 列已映射")

    def _rebuild_import_mapping(self, source_headers: list[str]):
        """根据源文件列头和多维表格字段，构建导入映射表"""
        for w in self.import_mapping_frame.winfo_children():
            w.destroy()
        self.import_mapping_combos.clear()

        # 下拉选项：「不映射」+ 所有可写入的表格列名
        writable = self._get_writable_fields()
        target_options = ["（不映射）"] + [f["field_name"] for f in writable]
        target_names_lower = {f["field_name"].lower(): f["field_name"] for f in writable}

        # 表头
        header = ttk.Frame(self.import_mapping_frame)
        header.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(header, text="源文件列", width=20, anchor=tk.W,
                  font=("", 9, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(header, text="示例值", width=18, anchor=tk.W,
                  font=("", 9, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(header, text="→ 映射到表格列", width=22, anchor=tk.W,
                  font=("", 9, "bold")).pack(side=tk.LEFT)

        # 取第一行数据做示例
        sample = self.import_data[0] if self.import_data else {}

        for src_col in source_headers:
            row = ttk.Frame(self.import_mapping_frame)
            row.pack(fill=tk.X, pady=1)

            ttk.Label(row, text=src_col, width=20, anchor=tk.W).pack(
                side=tk.LEFT, padx=(0, 8))

            # 示例值（截取前 20 字符）
            sample_val = str(sample.get(src_col, ""))[:20]
            ttk.Label(row, text=sample_val, width=18, anchor=tk.W,
                      foreground="gray").pack(side=tk.LEFT, padx=(0, 8))

            combo = ttk.Combobox(row, values=target_options, width=22, state="readonly")
            combo.pack(side=tk.LEFT)

            # 自动匹配：精确匹配 > 小写匹配 > 包含匹配
            matched_target = self._auto_match_import_col(src_col, writable, target_names_lower)
            if matched_target:
                idx = target_options.index(matched_target) if matched_target in target_options else 0
                combo.current(idx)
            else:
                combo.current(0)  # 不映射

            self.import_mapping_combos.append((src_col, combo))

    @staticmethod
    def _auto_match_import_col(src_col: str, writable: list[dict],
                               target_lower: dict) -> str | None:
        """自动匹配源列到目标列：精确 > 小写 > 包含"""
        # 精确匹配
        for f in writable:
            if f["field_name"] == src_col:
                return f["field_name"]
        # 小写匹配
        src_lower = src_col.lower().strip()
        if src_lower in target_lower:
            return target_lower[src_lower]
        # 包含匹配（源列名包含目标列名，或反过来）
        for f in writable:
            tgt = f["field_name"]
            if tgt in src_col or src_col in tgt:
                return tgt
        return None

    def _get_import_mapping(self) -> dict[str, str]:
        """获取用户配置的导入映射，返回 {源列名: 目标列名}"""
        mapping = {}
        for src_col, combo in self.import_mapping_combos:
            target = combo.get()
            if target and target != "（不映射）":
                mapping[src_col] = target
        return mapping

    @staticmethod
    def _read_tabular_file(path: str) -> list[dict]:
        ext = Path(path).suffix.lower()
        if ext == ".csv":
            with open(path, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                return list(reader)
        elif ext in (".xlsx", ".xls"):
            try:
                import openpyxl
            except ImportError:
                raise RuntimeError(
                    "读取 Excel 需要 openpyxl 库，请运行: pip install openpyxl")
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            rows_iter = ws.iter_rows(values_only=True)
            headers = [str(c) if c else f"列{i}" for i, c in enumerate(next(rows_iter))]
            result = []
            for row in rows_iter:
                d = {}
                for h, v in zip(headers, row):
                    d[h] = v if v is not None else ""
                result.append(d)
            wb.close()
            return result
        else:
            raise ValueError(f"不支持的文件格式: {ext}")

    # ---- 文件上传回填 ----

    # ---- 文件夹上传：核心逻辑 ----

    @staticmethod
    def _bt_read_text_file(filepath: str) -> str | None:
        """读取文本文件内容，非文本返回 None"""
        text_exts = {".txt", ".md", ".csv", ".json", ".xml", ".html", ".log", ".srt"}
        if Path(filepath).suffix.lower() not in text_exts:
            return None
        try:
            with open(filepath, "r", encoding="utf-8-sig", errors="ignore") as f:
                return f.read(4000).strip()
        except Exception:
            return None

    @staticmethod
    def _bt_scan_files(folder: str) -> dict[str, str]:
        """
        递归扫描文件夹，返回 {文件名: 完整路径} 字典。
        如果有同名文件，优先保留层级更浅的。
        """
        result = {}
        for root_dir, dirs, files in os.walk(folder):
            # 跳过隐藏/系统子目录
            dirs[:] = [d for d in dirs if not d.startswith(('.', '__', '$'))
                       and d != 'node_modules']
            depth = root_dir.replace(folder, '').count(os.sep)
            for fname in files:
                if fname.startswith('.'):
                    continue
                if fname not in result:
                    result[fname] = os.path.join(root_dir, fname)
        return result

    # 需要排除的文件夹名前缀
    _SKIP_FOLDER_PREFIXES = (".", "__", "node_modules", "$")

    def _bt_is_valid_subfolder(self, folder_path: str) -> tuple[bool, str]:
        """
        判断子文件夹是否有效，返回 (是否有效, 原因)。
        递归检查是否包含任何文件。
        """
        name = os.path.basename(folder_path)
        for prefix in self._SKIP_FOLDER_PREFIXES:
            if name.startswith(prefix):
                return False, f"系统/隐藏文件夹 ({prefix}...)"

        # 递归检查是否有任何文件
        file_map = self._bt_scan_files(folder_path)
        if not file_map:
            return False, "空文件夹（递归无文件）"

        return True, ""

    def _bt_calc_match_score(self, subfolder: str, mapped_files: list[str]) -> tuple[int, int]:
        """
        计算子文件夹对映射规则的匹配度（递归扫描，支持扩展名匹配）。
        """
        file_map = self._bt_scan_files(subfolder)
        matched = sum(1 for f in mapped_files if self._bt_find_file(file_map, f) is not None)
        return matched, len(mapped_files)

    def _bt_get_mapped_filenames(self) -> list[str]:
        """从当前映射配置中提取需要的文件名列表"""
        names = []
        for m in self.bt_col_mapping:
            selected = m["widget"].get()
            if selected and selected not in ("（跳过）", "（手动输入）"):
                names.append(selected)
        return names

    @staticmethod
    def _natural_sort_key(path: str):
        """自然排序 key：数字部分按数值排序，如 item2 < item10"""
        name = os.path.basename(path).lower()
        return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', name)]

    @staticmethod
    def _underscore_num_key(path: str):
        """
        按子文件夹内压缩包文件名的 _ 后数字排序。
        如子文件夹内有 yanyu_520.zip → 排序值为 520。
        找不到压缩包则按文件夹名中的数字排序。
        """
        # 先在子文件夹内找压缩包
        if os.path.isdir(path):
            for root_dir, _, files in os.walk(path):
                for f in files:
                    if f.lower().endswith(('.zip', '.rar', '.7z')):
                        m = re.search(r'_(\d+)', f)
                        if m:
                            return int(m.group(1))
                        nums = re.findall(r'\d+', f)
                        if nums:
                            return int(nums[-1])
                break  # 只看第一层，不递归太深

        # fallback: 按文件夹名中的数字
        name = os.path.basename(path)
        m = re.search(r'_(\d+)', name)
        if m:
            return int(m.group(1))
        nums = re.findall(r'\d+', name)
        if nums:
            return int(nums[-1])
        return 0

    def _bt_sort_dirs(self, dirs: list[str]) -> list[str]:
        """根据用户选择的排序方式排序文件夹列表"""
        mode = self.bt_sort_var.get()
        if mode == "原始顺序":
            return list(dirs)
        elif mode == "修改时间↑":
            return sorted(dirs, key=lambda d: os.path.getmtime(d))
        elif mode == "修改时间↓":
            return sorted(dirs, key=lambda d: os.path.getmtime(d), reverse=True)
        elif mode == "名称自然排序":
            return sorted(dirs, key=self._natural_sort_key)
        else:  # 名称_后数字（默认）
            return sorted(dirs, key=self._underscore_num_key)

    def _bt_resort_folders(self):
        """切换排序方式后重新排列列表"""
        if not self.bt_all_valid_dirs:
            return
        mapped_files = self._bt_get_mapped_filenames()
        sorted_dirs = self._bt_sort_dirs(self.bt_all_valid_dirs)
        self._bt_refresh_folder_list(sorted_dirs, mapped_files, self.bt_skipped)

    def _bt_select_folder(self):
        """选择父文件夹，扫描子文件夹并过滤"""
        folder = filedialog.askdirectory(title="选择包含子文件夹的父目录")
        if not folder:
            return

        if not self.bitable_fields:
            messagebox.showwarning("提示", "请先在上方粘贴表格链接并点击「解析并获取列」")
            return

        # 扫描直接子文件夹（保持文件系统原始顺序）
        all_dirs = []
        for name in os.listdir(folder):
            sub_path = os.path.join(folder, name)
            if os.path.isdir(sub_path):
                all_dirs.append(sub_path)

        if not all_dirs:
            messagebox.showwarning("提示",
                                   "该文件夹下没有子文件夹。\n\n"
                                   "正确结构应为：\n"
                                   "  父文件夹/\n"
                                   "    ├── 子文件夹A/  → 第1行\n"
                                   "    ├── 子文件夹B/  → 第2行\n"
                                   "    └── ...")
            return

        # 第一轮过滤：排除隐藏/空文件夹
        valid_dirs = []
        skipped = []
        for d in all_dirs:
            ok, reason = self._bt_is_valid_subfolder(d)
            if ok:
                valid_dirs.append(d)
            else:
                skipped.append((os.path.basename(d), reason))

        if not valid_dirs:
            skip_detail = "\n".join(f"  {n}: {r}" for n, r in skipped[:10])
            messagebox.showwarning("提示",
                                   f"所有 {len(all_dirs)} 个子文件夹均被排除：\n\n{skip_detail}")
            return

        # 按用户选择排序
        valid_dirs = self._bt_sort_dirs(valid_dirs)

        # 保存过滤结果（供切换排序时使用）
        self.bt_all_valid_dirs = valid_dirs
        self.bt_skipped = skipped

        # 用排序后第一个子文件夹作为模板（递归扫描）
        first = valid_dirs[0]
        sample_map = self._bt_scan_files(first)
        self.bt_sample_files = list(sample_map.values())
        self.bt_sample_map = sample_map

        # 构建映射 UI
        self.bt_subfolders = valid_dirs
        self._bt_build_mapping_ui()

        # 用映射规则做第二轮过滤
        mapped_files = self._bt_get_mapped_filenames()
        self._bt_refresh_folder_list(valid_dirs, mapped_files, skipped)

    def _bt_refresh_folder_list(self, valid_dirs: list[str],
                                mapped_files: list[str],
                                pre_skipped: list[tuple[str, str]]):
        """刷新子文件夹列表，计算匹配度并排除无匹配项"""
        self.bt_folder_tree.delete(*self.bt_folder_tree.get_children())
        self.bt_subfolders = []

        included = 0
        excluded_no_match = 0

        for sf in valid_dirs:
            name = os.path.basename(sf)

            if mapped_files:
                hit, total = self._bt_calc_match_score(sf, mapped_files)
                match_text = f"{hit}/{total}"
            else:
                hit, total = 0, 0
                match_text = "-"

            if mapped_files and hit == 0:
                self.bt_folder_tree.insert("", tk.END, values=(
                    name, match_text, "排除(无匹配)"))
                excluded_no_match += 1
            else:
                self.bt_subfolders.append(sf)
                if total > 0 and hit == total:
                    status = "待上传"
                elif total > 0:
                    status = f"缺{total - hit}个文件"
                else:
                    status = "待上传"
                self.bt_folder_tree.insert("", tk.END, values=(
                    name, match_text, status))
                included += 1

        for skip_name, reason in pre_skipped:
            self.bt_folder_tree.insert("", tk.END, values=(
                skip_name, "-", reason))

        total_scanned = len(valid_dirs) + len(pre_skipped)
        self.bt_status_var.set(
            f"扫描 {total_scanned} 个文件夹 → "
            f"有效 {included} 个 | "
            f"排除 {excluded_no_match + len(pre_skipped)} 个"
            f"（无匹配 {excluded_no_match}, 隐藏/空 {len(pre_skipped)}）")

        if included == 0:
            messagebox.showwarning("提示",
                                   f"没有有效的子文件夹。\n"
                                   f"请检查映射规则和文件夹内容是否一致。")

    def _bt_clear(self):
        """清空所有状态"""
        self.bt_subfolders.clear()
        self.bt_all_valid_dirs.clear()
        self.bt_skipped.clear()
        self.bt_sample_files.clear()
        self.bt_sample_map = {}
        self.bt_col_mapping.clear()
        for w in self.bt_mapping_inner.winfo_children():
            w.destroy()
        self.bt_folder_tree.delete(*self.bt_folder_tree.get_children())
        self.bt_status_var.set("选择包含多个子文件夹的父目录，每个子文件夹 = 一行数据")

    def _bt_build_mapping_ui(self):
        """用第一个子文件夹的文件（递归扫描），构建 列 ↔ 文件 映射 UI"""
        for w in self.bt_mapping_inner.winfo_children():
            w.destroy()
        self.bt_col_mapping.clear()

        writable = self._get_writable_fields()
        if not writable:
            ttk.Label(self.bt_mapping_inner, text="没有可写入的列",
                      foreground="gray").pack(anchor=tk.W)
            return

        # 模板文件夹的文件（递归扫描结果）
        sample_map = getattr(self, 'bt_sample_map', {})
        sample_names = sorted(sample_map.keys())
        sample_stems = {}  # stem_lower → basename
        for name in sample_names:
            stem = Path(name).stem.lower().strip()
            if stem not in sample_stems:
                sample_stems[stem] = name

        # 下拉选项
        file_options_attach = ["（跳过）"] + sample_names
        file_options_text = ["（手动输入）"] + sample_names

        # 表头
        first_name = os.path.basename(self.bt_subfolders[0])
        ttk.Label(self.bt_mapping_inner,
                  text=f"模板：{first_name}/  "
                       f"（递归扫描到 {len(sample_names)} 个文件，所有子文件夹按相同规则处理）",
                  foreground="blue").pack(anchor=tk.W, pady=(0, 6))

        matched_count = 0

        for f in writable:
            col_name = f["field_name"]
            col_type = f.get("type", 0)
            type_name = FIELD_TYPE_NAMES.get(col_type, f"未知")
            is_attachment = (col_type == 17)

            row = ttk.Frame(self.bt_mapping_inner)
            row.pack(fill=tk.X, pady=2, padx=4)

            ttk.Label(row, text=col_name, width=14, anchor=tk.W,
                      font=("", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
            ttk.Label(row, text=f"[{type_name}]", width=8, anchor=tk.W,
                      foreground="gray").pack(side=tk.LEFT, padx=(0, 6))

            # 自动匹配：文件名（去扩展名）== 列名
            col_lower = col_name.lower().strip()
            matched_name = sample_stems.get(col_lower)

            if is_attachment:
                combo = ttk.Combobox(row, values=file_options_attach, width=28, state="readonly")
                if matched_name:
                    idx = file_options_attach.index(matched_name)
                    combo.current(idx)
                    status = "✓ 自动匹配"
                    matched_count += 1
                else:
                    combo.current(0)
                    status = "— 可选择文件"
                combo.pack(side=tk.LEFT, padx=(0, 4))
                ttk.Label(row, text=status,
                          foreground="green" if matched_name else "orange").pack(side=tk.LEFT)

                self.bt_col_mapping.append({
                    "col_name": col_name, "col_type": col_type,
                    "mode": "attach", "widget": combo,
                })
            else:
                combo = ttk.Combobox(row, values=file_options_text, width=16, state="readonly")
                entry = ttk.Entry(row, width=22)

                if matched_name:
                    matched_fp = sample_map.get(matched_name)
                    content = self._bt_read_text_file(matched_fp) if matched_fp else None
                    if content:
                        idx = file_options_text.index(matched_name)
                        combo.current(idx)
                        preview = content.split('\n')[0][:60]
                        entry.insert(0, preview)
                        entry.configure(state="readonly")
                        status = "✓ 读取文件内容"
                        matched_count += 1
                    else:
                        combo.current(0)
                        status = "可输入或选文件"
                else:
                    combo.current(0)
                    status = "可输入或选文件"

                combo.pack(side=tk.LEFT, padx=(0, 4))

                # 下拉变化 → 更新 entry 预览（用 sample_map 查找路径）
                def _on_change(event, cb=combo, ent=entry, smap=sample_map):
                    selected = cb.get()
                    ent.configure(state="normal")
                    ent.delete(0, tk.END)
                    if selected != "（手动输入）":
                        fp = smap.get(selected)
                        if fp:
                            content = self._bt_read_text_file(fp)
                            if content:
                                ent.insert(0, content.split('\n')[0][:200])
                            else:
                                ent.insert(0, "[非文本文件]")
                                ent.configure(state="readonly")
                combo.bind("<<ComboboxSelected>>", _on_change)

                entry.pack(side=tk.LEFT, padx=(0, 4))
                check_fp = sample_map.get(matched_name) if matched_name else None
                fg = "green" if check_fp and self._bt_read_text_file(check_fp) else "gray"
                ttk.Label(row, text=status, foreground=fg).pack(side=tk.LEFT)

                self.bt_col_mapping.append({
                    "col_name": col_name, "col_type": col_type,
                    "mode": "text", "widget": combo, "entry": entry,
                })

        unmatched = len(writable) - matched_count
        self.status_var.set(
            f"识别完成：{matched_count} 列已自动匹配，{unmatched} 列待确认，"
            f"共 {len(self.bt_subfolders)} 个子文件夹将生成 {len(self.bt_subfolders)} 行数据")

    def _rebuild_mapping_table(self):
        """解析链接后，如果已有子文件夹则重新识别"""
        if self.bt_subfolders and self.bitable_fields:
            self._bt_build_mapping_ui()

    # ---- 统一提交 ----

    def _submit_bitable(self):
        if not self.bitable_app_token or not self.bitable_table_id:
            messagebox.showwarning("提示", "请先解析多维表格链接")
            return
        if not self._ensure_client():
            return

        mode = self._get_active_fill_mode()
        self.bitable_submit_btn.configure(state=tk.DISABLED)

        if mode == "manual":
            threading.Thread(target=self._submit_manual, daemon=True).start()
        elif mode == "import":
            threading.Thread(target=self._submit_import, daemon=True).start()
        elif mode == "file_upload":
            threading.Thread(target=self._submit_file_upload, daemon=True).start()

    def _submit_done(self, success, failed, total):
        self.root.after(0, lambda: self.progress_var.set(100))
        self.root.after(0, lambda: self.status_var.set(
            f"提交完成 — 成功: {success}, 失败: {failed}"))
        self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
        self.root.after(0, lambda s=success, f=failed: messagebox.showinfo(
            "提交完成", f"成功写入: {s} 条\n失败: {f} 条"))

    def _get_start_row(self) -> int:
        """获取用户设置的起始行（从 1 开始）"""
        try:
            return max(1, int(self.start_row_var.get()))
        except ValueError:
            return 1

    def _is_update_mode(self) -> bool:
        return self.write_mode_var.get() == "update"

    def _ensure_existing_records(self):
        """确保已有记录列表是最新的（更新模式下需要）"""
        if not self._existing_records:
            records = self.client.bitable_list_all_records(
                self.bitable_app_token, self.bitable_table_id)
            self._existing_records = records

    def _ask_user_sync(self, title: str, message: str) -> bool:
        """在工作线程中同步弹出确认对话框，返回用户选择"""
        result = threading.Event()
        user_choice = [False]

        def _ask():
            user_choice[0] = messagebox.askyesno(title, message)
            result.set()

        self.root.after(0, _ask)
        result.wait()
        return user_choice[0]

    def _preflight_check(self, records: list[dict]) -> bool:
        """
        提交前预检查：
        - 更新模式：检查表格是否有足够的行，不够则提示用户自动补齐
        - 追加模式：检查起始行是否超出源数据范围
        返回 True 表示检查通过可继续，False 表示用户取消
        """
        start_row = self._get_start_row()

        if not self._is_update_mode():
            # 追加模式：检查起始行不超出数据范围
            actual_count = len(records) - (start_row - 1)
            if actual_count <= 0:
                self.root.after(0, lambda: messagebox.showwarning(
                    "数据不足",
                    f"源数据共 {len(records)} 行，起始行设为第 {start_row} 行，\n"
                    f"没有可提交的数据。请调整起始行。"))
                return False
            return True

        # 更新模式：检查表格行数
        self.root.after(0, lambda: self.status_var.set("正在检查表格现有记录数..."))
        self._ensure_existing_records()
        existing_count = len(self._existing_records)
        update_start = start_row - 1
        needed_end = update_start + len(records)  # 需要覆盖到第几行
        gap = needed_end - existing_count  # 缺少的行数

        if gap <= 0:
            # 行数足够
            return True

        # 起始行就超出已有行数
        if update_start >= existing_count:
            msg = (
                f"表格当前共 {existing_count} 行，\n"
                f"你设置从第 {start_row} 行开始更新，但表格只有 {existing_count} 行。\n"
                f"需要写入 {len(records)} 条数据，共需要新增 {gap} 行。\n\n"
                f"是否自动新增 {gap} 行空记录后继续？"
            )
        else:
            msg = (
                f"表格当前共 {existing_count} 行，\n"
                f"从第 {start_row} 行开始更新 {len(records)} 条数据，\n"
                f"需要写到第 {needed_end} 行，超出现有行数 {gap} 行。\n\n"
                f"是否自动新增 {gap} 行空记录后继续？"
            )

        if not self._ask_user_sync("行数不足", msg):
            return False

        # 自动创建空行
        self.root.after(0, lambda g=gap: self.status_var.set(f"正在自动新增 {g} 行..."))
        writable = self._get_writable_fields()
        # 用第一个可写字段创建空记录（飞书不允许完全空的 fields）
        placeholder_field = writable[0]["field_name"] if writable else None
        empty_records = []
        for _ in range(gap):
            if placeholder_field:
                empty_records.append({"fields": {placeholder_field: ""}})
            else:
                empty_records.append({"fields": {}})

        # 分批创建空行
        batch_size = 500
        for i in range(0, len(empty_records), batch_size):
            batch = empty_records[i:i + batch_size]
            self.client.bitable_create_records(
                self.bitable_app_token, self.bitable_table_id, batch)

        # 刷新记录缓存
        self._existing_records = self.client.bitable_list_all_records(
            self.bitable_app_token, self.bitable_table_id)
        new_count = len(self._existing_records)
        self.root.after(0, lambda c=new_count: self.record_count_var.set(f"当前共 {c} 行记录"))
        self.root.after(0, lambda g=gap: self.status_var.set(f"已自动新增 {g} 行，继续写入..."))
        return True

    def _submit_records(self, records: list[dict]):
        """
        根据 write_mode 和 start_row 提交记录。
        - 追加模式：跳过源数据前 start_row-1 行，剩余追加为新记录
        - 更新模式：从表格第 start_row 行开始逐行覆盖
        """
        # 前置检查
        if not self._preflight_check(records):
            return 0, 0

        start_row = self._get_start_row()

        if not self._is_update_mode():
            # 追加模式：跳过源数据前 N-1 行
            records = records[start_row - 1:]
            self._batch_create_records(records)
            return len(records), 0

        # 更新模式
        existing = self._existing_records
        update_start = start_row - 1
        success = failed = 0
        total = len(records)

        for i, rec in enumerate(records):
            target_idx = update_start + i
            self.root.after(0, lambda v=(i + 1) / total * 100: self.progress_var.set(v))
            self.root.after(0, lambda ii=i, t=total, s=update_start: self.status_var.set(
                f"正在写入第 {s + ii + 1} 行 ({ii + 1}/{t})..."))

            try:
                if target_idx < len(existing):
                    record_id = existing[target_idx]["record_id"]
                    self.client.bitable_update_record(
                        self.bitable_app_token, self.bitable_table_id,
                        record_id, rec["fields"])
                else:
                    self.client.bitable_create_records(
                        self.bitable_app_token, self.bitable_table_id, [rec])
                success += 1
            except Exception as e:
                failed += 1
                self.root.after(0, lambda m=str(e)[:50], r=target_idx: self.status_var.set(
                    f"第 {r + 1} 行写入失败: {m}"))

        return success, failed

    def _submit_manual(self):
        writable = self._get_writable_fields()
        records = []
        for row_entries in self.manual_entries:
            fields = {}
            for entry, fdef in zip(row_entries, writable):
                val = entry.get().strip()
                if val:
                    fields[fdef["field_name"]] = self._cast_value(val, fdef)
            if fields:
                records.append({"fields": fields})

        if not records:
            self.root.after(0, lambda: messagebox.showwarning("提示", "请至少填写一行数据"))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
            return

        self.root.after(0, lambda: self.status_var.set(f"正在提交 {len(records)} 条记录..."))
        try:
            self.client.get_tenant_access_token()
            success, failed = self._submit_records(records)
            if success == 0 and failed == 0:
                # 用户取消了预检查
                self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
                return
            self._submit_done(success, failed, len(records))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("提交失败", str(e)))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))

    def _submit_import(self):
        if not self.import_data:
            self.root.after(0, lambda: messagebox.showwarning("提示", "请先预览导入数据"))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
            return

        # 使用用户配置的映射
        col_mapping = self._get_import_mapping()
        if not col_mapping:
            self.root.after(0, lambda: messagebox.showwarning(
                "提示", "没有映射任何列，请在映射表中至少将一列映射到表格列"))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
            return

        field_defs = {f["field_name"]: f for f in self.bitable_fields}
        records = []
        for row in self.import_data:
            fields = {}
            for src_col, tgt_col in col_mapping.items():
                val = row.get(src_col)
                if val not in (None, ""):
                    if tgt_col in field_defs:
                        fields[tgt_col] = self._cast_value(str(val), field_defs[tgt_col])
                    else:
                        fields[tgt_col] = str(val)
            if fields:
                records.append({"fields": fields})

        if not records:
            self.root.after(0, lambda: messagebox.showwarning("提示", "没有有效的数据行"))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
            return

        self.root.after(0, lambda: self.status_var.set(f"正在提交 {len(records)} 条记录..."))
        try:
            self.client.get_tenant_access_token()
            success, failed = self._submit_records(records)
            if success == 0 and failed == 0:
                self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
                return
            self._submit_done(success, failed, len(records))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("提交失败", str(e)))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))

    @staticmethod
    def _bt_find_file(file_map: dict, template_name: str) -> str | None:
        """
        在 file_map 中查找文件。
        优先精确匹配文件名，找不到则按扩展名匹配（同扩展名只有一个时自动匹配）。
        """
        # 1. 精确匹配
        if template_name in file_map:
            return file_map[template_name]

        # 2. 按扩展名匹配
        ext = Path(template_name).suffix.lower()
        if ext:
            candidates = [fp for name, fp in file_map.items()
                          if Path(name).suffix.lower() == ext]
            if len(candidates) == 1:
                return candidates[0]
            # 多个同扩展名文件：按 stem 相似度选最接近的
            if candidates:
                template_stem = Path(template_name).stem.lower()
                # 优先选 stem 前缀相同的
                for fp in candidates:
                    if Path(fp).stem.lower().startswith(template_stem[:3]):
                        return fp
                return candidates[0]  # 兜底返回第一个

        return None

    def _bt_build_record_for_subfolder(self, subfolder: str) -> tuple[dict, list]:
        """
        根据映射规则，为一个子文件夹构建 fields 和附件列表。
        递归扫描子文件夹，按扩展名模糊匹配文件。
        返回 (fields_dict, [(col_name, filepath), ...])
        """
        file_map = self._bt_scan_files(subfolder)

        fields = {}
        attachments = []

        for m in self.bt_col_mapping:
            col_name = m["col_name"]
            selected = m["widget"].get()

            if m["mode"] == "attach":
                if selected and selected != "（跳过）":
                    fp = self._bt_find_file(file_map, selected)
                    if fp and os.path.isfile(fp):
                        attachments.append((col_name, fp))
            elif m["mode"] == "text":
                if selected == "（手动输入）":
                    val = m["entry"].get().strip()
                    if val:
                        fields[col_name] = self._cast_value(val, {"type": m["col_type"]})
                else:
                    fp = self._bt_find_file(file_map, selected)
                    if fp and os.path.isfile(fp):
                        content = self._bt_read_text_file(fp)
                        if content:
                            fields[col_name] = content
        return fields, attachments

    def _bt_update_tree_status(self, idx: int, status: str):
        """更新子文件夹列表中第 idx 行的状态列"""
        try:
            children = self.bt_folder_tree.get_children()
            if idx < len(children):
                self.bt_folder_tree.set(children[idx], "status", status)
                self.bt_folder_tree.see(children[idx])
        except Exception:
            pass

    def _bt_log_init(self):
        """初始化日志文件"""
        import datetime
        with open(self.bt_log_path, "w", encoding="utf-8") as f:
            f.write(f"飞书上传日志 — {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")

    def _bt_log_msg(self, msg: str):
        """追加一行到日志文件"""
        with open(self.bt_log_path, "a", encoding="utf-8") as f:
            f.write(msg + "\n")

    def _bt_log_open(self):
        """用系统默认程序打开日志文件"""
        import subprocess
        try:
            os.startfile(self.bt_log_path)
        except AttributeError:
            subprocess.Popen(["open", self.bt_log_path])

    def _submit_file_upload(self):
        """遍历所有子文件夹，逐个上传并实时更新状态，日志写入文件"""
        if not self.bt_subfolders:
            self.root.after(0, lambda: messagebox.showwarning("提示", "请先选择文件夹"))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
            return
        if not self.bt_col_mapping:
            self.root.after(0, lambda: messagebox.showwarning("提示", "没有映射规则，请重新选择文件夹"))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
            return

        # 初始化日志文件
        self._bt_log_init()

        try:
            self.client.get_tenant_access_token()
            self._bt_log_msg("认证成功\n")
        except Exception as e:
            err = str(e)
            self._bt_log_msg(f"认证失败: {err}")
            self.root.after(0, lambda m=err: messagebox.showerror("认证失败", m))
            self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))
            self._bt_log_open()
            return

        total = len(self.bt_subfolders)
        success = failed = skipped = 0
        fail_details = []

        start_row = self._get_start_row()
        is_update = self._is_update_mode()
        mode_text = "更新" if is_update else "追加"

        self._bt_log_msg(f"共 {total} 个子文件夹待处理")
        self._bt_log_msg(f"模式: {mode_text}  起始行: {start_row}\n")

        # 更新模式下预加载已有记录
        existing_records = []
        if is_update:
            try:
                self._bt_log_msg("正在获取现有记录...")
                existing_records = self.client.bitable_list_all_records(
                    self.bitable_app_token, self.bitable_table_id)
                self._bt_log_msg(f"现有 {len(existing_records)} 条记录\n")
            except Exception as e:
                self._bt_log_msg(f"获取记录失败: {e}\n")

        for idx, sf in enumerate(self.bt_subfolders):
            sf_name = os.path.basename(sf)
            row_num = start_row + idx  # 当前写入第几行
            num = f"[{idx+1}/{total}] 第{row_num}行"

            self.root.after(0, lambda n=sf_name, s=num: self.status_var.set(f"{s} {n}"))
            self.root.after(0, lambda i=idx, t=total: self.progress_var.set((i + 1) / t * 100))
            self.root.after(0, lambda i=idx: self._bt_update_tree_status(i, "处理中"))

            self._bt_log_msg(f"{num} {sf_name}")

            # 构建记录
            fields, attachments = self._bt_build_record_for_subfolder(sf)
            self._bt_log_msg(f"    文本字段: {list(fields.keys()) if fields else '无'}")
            self._bt_log_msg(f"    待上传附件: {[os.path.basename(fp) for _, fp in attachments] if attachments else '无'}")

            # 上传附件
            for col_name, fp in attachments:
                fname = os.path.basename(fp)
                fsize = os.path.getsize(fp)
                self.root.after(0, lambda i=idx, n=fname: self._bt_update_tree_status(
                    i, f"上传 {n}"))
                try:
                    file_token = self.client.bitable_upload_media(
                        fp, self.bitable_app_token)
                    fields[col_name] = [{"file_token": file_token}]
                    self._bt_log_msg(f"    附件OK: {fname} ({self._format_size(fsize)})")
                except Exception as e:
                    err = str(e)
                    self._bt_log_msg(f"    附件失败: {fname} → {err}")
                    fail_details.append(f"{sf_name}/{fname}: {err}")

            if not fields:
                self._bt_log_msg(f"    → 跳过（无有效数据）\n")
                self.root.after(0, lambda i=idx: self._bt_update_tree_status(i, "跳过"))
                skipped += 1
                continue

            # 写入记录（根据模式选择追加或更新）
            self.root.after(0, lambda i=idx: self._bt_update_tree_status(i, "写入中"))
            try:
                target_idx = row_num - 1  # 0-based
                if is_update and target_idx < len(existing_records):
                    record_id = existing_records[target_idx]["record_id"]
                    self.client.bitable_update_record(
                        self.bitable_app_token, self.bitable_table_id,
                        record_id, fields)
                    self._bt_log_msg(f"    → 更新第{row_num}行成功\n")
                else:
                    self.client.bitable_create_records(
                        self.bitable_app_token, self.bitable_table_id,
                        [{"fields": fields}])
                    self._bt_log_msg(f"    → 追加成功（第{row_num}行）\n")
                success += 1
                self.root.after(0, lambda i=idx: self._bt_update_tree_status(i, "成功"))
            except Exception as e:
                failed += 1
                err = str(e)
                self._bt_log_msg(f"    → 写入失败: {err}\n")
                self.root.after(0, lambda i=idx: self._bt_update_tree_status(i, "失败"))
                fail_details.append(f"{sf_name}: {err}")

        # 写入汇总
        self._bt_log_msg("=" * 60)
        self._bt_log_msg(f"成功: {success}  失败: {failed}  跳过: {skipped}  共: {total}")
        if fail_details:
            self._bt_log_msg(f"\n失败详情:")
            for d in fail_details:
                self._bt_log_msg(f"  - {d}")
        self._bt_log_msg(f"\n日志文件: {self.bt_log_path}")

        # 界面更新
        self.root.after(0, lambda: self.progress_var.set(100))
        self.root.after(0, lambda s=success, f=failed, sk=skipped, t=total:
                        self.status_var.set(f"完成 — 成功{s} 失败{f} 跳过{sk} 共{t}"))
        self.root.after(0, lambda: self.bitable_submit_btn.configure(state=tk.NORMAL))

        # 直接打开日志文件，用户可自由查看和复制
        self.root.after(0, self._bt_log_open)

    def _batch_create_records(self, records: list[dict]):
        """分批提交（每批最多 500 条）"""
        batch_size = 500
        for i in range(0, len(records), batch_size):
            batch = records[i:i + batch_size]
            self.client.bitable_create_records(
                self.bitable_app_token, self.bitable_table_id, batch)
            pct = min((i + batch_size) / len(records), 1.0) * 100
            self.root.after(0, lambda v=pct: self.progress_var.set(v))

    @staticmethod
    def _cast_value(val: str, field_def: dict):
        """根据字段类型转换值"""
        ftype = field_def.get("type")
        if ftype == 2:  # 数字
            try:
                return float(val) if '.' in val else int(val)
            except ValueError:
                return val
        if ftype == 7:  # 复选框
            return val.lower() in ("true", "1", "yes", "是")
        if ftype == 4:  # 多选
            return [v.strip() for v in val.split(",") if v.strip()]
        return val


# ============================================================
# 入口
# ============================================================

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    FeishuToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
