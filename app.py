# -*- coding: utf-8 -*-
import os
import uuid
import threading
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify, url_for

from processor import process_excel

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 日志与状态（演示够用）
TASK_LOGS = {}      # task_id -> list[str]
TASK_STATUS = {}    # task_id -> {"state": "running|done|error", "error": str|None}


def add_log(task_id: str, msg: str):
    t = datetime.now().strftime("%H:%M:%S")
    TASK_LOGS.setdefault(task_id, []).append(f"[{t}] {msg}")


def run_task(task_id: str, param_path: str, report_path: str, out_path: str):
    try:
        add_log(task_id, "后台任务开始执行…")
        process_excel(param_path, report_path, out_path, log_fn=lambda m: add_log(task_id, m))
        add_log(task_id, "后台任务完成 ✅")
        TASK_STATUS[task_id] = {"state": "done", "error": None}
    except Exception as e:
        add_log(task_id, f"后台任务失败 ❌：{type(e).__name__}: {e}")
        TASK_STATUS[task_id] = {"state": "error", "error": f"{type(e).__name__}: {e}"}


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/api/start", methods=["POST"])
def start():
    # 只负责：收文件 + 保存 + 启动后台线程 + 立刻返回 task_id
    if "param_file" not in request.files or "report_file" not in request.files:
        return jsonify({"error": "请上传 参数表 和 报表 两个文件"}), 400

    param_file = request.files["param_file"]
    report_file = request.files["report_file"]

    if param_file.filename == "" or report_file.filename == "":
        return jsonify({"error": "文件名为空，请重新选择文件"}), 400

    task_id = str(uuid.uuid4())
    TASK_LOGS[task_id] = []
    TASK_STATUS[task_id] = {"state": "running", "error": None}

    add_log(task_id, f"收到参数表：{param_file.filename}")
    add_log(task_id, f"收到报表：{report_file.filename}")
    add_log(task_id, "保存上传文件…")

    param_path = os.path.join(UPLOAD_DIR, f"{task_id}_param.xlsx")
    report_path = os.path.join(UPLOAD_DIR, f"{task_id}_report.xlsx")
    param_file.save(param_path)
    report_file.save(report_path)

    add_log(task_id, "文件保存完成，启动后台处理…")

    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")

    t = threading.Thread(target=run_task, args=(task_id, param_path, report_path, out_path), daemon=True)
    t.start()

    return jsonify({
        "task_id": task_id,
        "download_url": url_for("download", task_id=task_id),
        "status_url": url_for("status", task_id=task_id),
        "log_url": url_for("get_log", task_id=task_id),
    })


@app.route("/api/log/<task_id>", methods=["GET"])
def get_log(task_id):
    return jsonify({"logs": TASK_LOGS.get(task_id, [])})


@app.route("/api/status/<task_id>", methods=["GET"])
def status(task_id):
    s = TASK_STATUS.get(task_id)
    if not s:
        return jsonify({"state": "unknown"}), 404

    # 双保险：文件存在也视为 done
    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")
    if os.path.exists(out_path):
        TASK_STATUS[task_id] = {"state": "done", "error": None}
        s = TASK_STATUS[task_id]

    return jsonify(s)


@app.route("/api/download/<task_id>", methods=["GET"])
def download(task_id):
    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")
    if not os.path.exists(out_path):
        return "任务不存在或尚未生成结果", 404
    return send_file(out_path, as_attachment=True, download_name="最终结果.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)