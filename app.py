# -*- coding: utf-8 -*-
import os
import uuid
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify, url_for

from processor import process_excel

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 内存日志：够比赛演示用
TASK_LOGS = {}  # task_id -> list[str]


def add_log(task_id: str, msg: str):
    t = datetime.now().strftime("%H:%M:%S")
    TASK_LOGS.setdefault(task_id, []).append(f"[{t}] {msg}")


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/api/start", methods=["POST"])
def start():
    """
    接收两个 Excel：
      - param_file: 参数表
      - report_file: 报表
    调用 processor.process_excel 生成 outputs/{task_id}_最终结果.xlsx
    """
    if "param_file" not in request.files or "report_file" not in request.files:
        return jsonify({"error": "请上传 参数表 和 报表 两个文件"}), 400

    param_file = request.files["param_file"]
    report_file = request.files["report_file"]

    if param_file.filename == "" or report_file.filename == "":
        return jsonify({"error": "文件名为空，请重新选择文件"}), 400

    task_id = str(uuid.uuid4())
    TASK_LOGS[task_id] = []

    add_log(task_id, f"收到参数表：{param_file.filename}")
    add_log(task_id, f"收到报表：{report_file.filename}")
    add_log(task_id, "开始处理…")

    # 保存上传文件
    param_path = os.path.join(UPLOAD_DIR, f"{task_id}_param.xlsx")
    report_path = os.path.join(UPLOAD_DIR, f"{task_id}_report.xlsx")
    param_file.save(param_path)
    report_file.save(report_path)
    add_log(task_id, "文件已保存到服务器临时目录")

    # 输出文件固定命名（下载时按磁盘文件找，不依赖内存）
    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")

    try:
        add_log(task_id, "开始执行真实处理逻辑…")
        process_excel(param_path, report_path, out_path)
        add_log(task_id, "真实处理完成，已生成最终结果.xlsx")
    except Exception as e:
        add_log(task_id, f"处理失败：{type(e).__name__}: {e}")
        return jsonify({"error": f"处理失败：{type(e).__name__}: {e}", "task_id": task_id}), 500

    add_log(task_id, "处理结束，可以下载结果")

    return jsonify({
        "task_id": task_id,
        "download_url": url_for("download", task_id=task_id)
    })


@app.route("/api/log/<task_id>", methods=["GET"])
def get_log(task_id):
    return jsonify({"logs": TASK_LOGS.get(task_id, [])})


@app.route("/api/download/<task_id>", methods=["GET"])
def download(task_id):
    # 关键修复：不再依赖 TASK_OUTPUT（WSGI 多进程/重载会丢内存）
    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")
    if not os.path.exists(out_path):
        return "任务不存在或尚未生成结果", 404

    return send_file(out_path, as_attachment=True, download_name="最终结果.xlsx")


if __name__ == "__main__":
    # 本地运行：访问 http://127.0.0.1:5000
    app.run(host="0.0.0.0", port=5000, debug=True)
