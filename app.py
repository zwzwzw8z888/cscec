#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""中建四局公文格式化 Web 服务"""

import os
import time
import tempfile
import uuid
from pathlib import Path

from flask import Flask, request, render_template, send_file, redirect, url_for, flash

from formatter_core import format_document

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-" + str(uuid.uuid4()))
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB max

EXPIRY_SECONDS = 10 * 60  # 文件10分钟后过期

UPLOAD_DIR = Path(tempfile.gettempdir()) / "cscec-formatter"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# 内存中记录文件创建时间（服务重启后清空，与 /tmp 保持同步）
_file_registry: dict[str, float] = {}


def _cleanup_expired():
    """清理过期文件"""
    now = time.time()
    expired = [
        name for name, created in _file_registry.items()
        if now - created > EXPIRY_SECONDS
    ]
    for name in expired:
        (UPLOAD_DIR / name).unlink(missing_ok=True)
        _file_registry.pop(name, None)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/format", methods=["POST"])
def do_format():
    file = request.files.get("file")
    if not file or file.filename == "":
        flash("请先选择文件再上传", "error")
        return redirect(url_for("index"))

    if not file.filename.lower().endswith(".docx"):
        flash("仅支持 .docx 格式的 Word 文档", "error")
        return redirect(url_for("index"))

    _cleanup_expired()

    uid = uuid.uuid4().hex[:8]
    src_path = UPLOAD_DIR / f"src_{uid}.docx"
    dst_temp = UPLOAD_DIR / f"out_{uid}.docx"

    # 下载时的友好文件名：原文件名-格式化版.docx
    base = file.filename.rsplit(".", 1)[0] if "." in file.filename else file.filename
    download_name = f"{base}-格式化版.docx"
    dst_final = UPLOAD_DIR / download_name

    try:
        file.save(str(src_path))
        dst, warnings = format_document(str(src_path), str(dst_temp))
        dst_temp.rename(dst_final)

        # 清理源文件
        src_path.unlink(missing_ok=True)

        # 注册文件，记录创建时间
        _file_registry[download_name] = time.time()

        return render_template(
            "result.html",
            filename=file.filename,
            download_name=download_name,
            warnings=warnings,
            expiry_minutes=EXPIRY_SECONDS // 60,
        )

    except Exception as e:
        src_path.unlink(missing_ok=True)
        dst_temp.unlink(missing_ok=True)
        dst_final.unlink(missing_ok=True)
        flash(f"处理失败：{e}", "error")
        return redirect(url_for("index"))


@app.route("/download/<name>")
def download(name):
    path = UPLOAD_DIR / name

    # 检查是否过期
    created = _file_registry.get(name)
    if created and time.time() - created > EXPIRY_SECONDS:
        path.unlink(missing_ok=True)
        _file_registry.pop(name, None)
        flash("文件已过期（超过10分钟），请重新上传格式化", "error")
        return redirect(url_for("index"))

    if not path.exists():
        flash("文件不存在，请重新上传", "error")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True, download_name=name, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/cleanup/<name>")
def cleanup(name):
    """下载完成后清理输出文件"""
    (UPLOAD_DIR / name).unlink(missing_ok=True)
    _file_registry.pop(name, None)
    return "", 204


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
