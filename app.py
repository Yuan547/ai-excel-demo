# -*- coding: utf-8 -*-
import os
import uuid
import json
import threading
from datetime import datetime

from flask import Flask, render_template, request, send_file, jsonify, url_for

from processor import process_excel

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

# ✅ 新增：任务日志与状态落盘（解决 PythonAnywhere 上“假卡住/丢日志”）
LOG_DIR = os.path.join(BASE_DIR, "task_logs")
STATUS_DIR = os.path.join(BASE_DIR, "task_status")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(STATUS_DIR, exist_ok=True)

# 可选：限制上传大小（默认 50MB，可按需调大/调小）
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "50")) * 1024 * 1024


def _log_path(task_id: str) -> str:
    return os.path.join(LOG_DIR, f"{task_id}.log")


def _status_path(task_id: str) -> str:
    return os.path.join(STATUS_DIR, f"{task_id}.json")


def add_log(task_id: str, msg: str):
    """把日志追加写入 task_logs/<task_id>.log"""
    t = datetime.now().strftime("%H:%M:%S")
    line = f"[{t}] {msg}\n"
    p = _log_path(task_id)
    with open(p, "a", encoding="utf-8") as f:
        f.write(line)
        f.flush()


def write_status(task_id: str, state: str, error: str | None = None):
    """把状态写入 task_status/<task_id>.json"""
    p = _status_path(task_id)
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"state": state, "error": error}, f, ensure_ascii=False)


def read_status(task_id: str):
    p = _status_path(task_id)
    if not os.path.exists(p):
        return None
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def run_task(task_id: str, param_path: str, report_path: str, out_path: str):
    try:
        add_log(task_id, "后台任务开始执行…")
        write_status(task_id, "running", None)

        t0 = datetime.now()
        add_log(task_id, "开始处理 Excel…")
        process_excel(
            param_path,
            report_path,
            out_path,
            log_fn=lambda m: add_log(task_id, m),
        )
        dt = (datetime.now() - t0).total_seconds()
        add_log(task_id, f"处理完成，用时 {dt:.1f}s ✅")

        add_log(task_id, "后台任务完成 ✅")
        write_status(task_id, "done", None)
    except Exception as e:
        add_log(task_id, f"后台任务失败 ❌：{type(e).__name__}: {e}")
        write_status(task_id, "error", f"{type(e).__name__}: {e}")


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

    # 初始化状态/日志（落盘）
    write_status(task_id, "running", None)
    add_log(task_id, f"收到参数表：{param_file.filename}")
    add_log(task_id, f"收到报表：{report_file.filename}")
    add_log(task_id, "保存上传文件…")

    param_path = os.path.join(UPLOAD_DIR, f"{task_id}_param.xlsx")
    report_path = os.path.join(UPLOAD_DIR, f"{task_id}_report.xlsx")
    param_file.save(param_path)
    report_file.save(report_path)

    add_log(task_id, "文件保存完成，启动后台处理…")

    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")

    # ⚠️ 仍然使用后台线程（演示用最省事）
    # 但 PythonAnywhere 可能会回收 worker；日志落盘后至少不会“假卡住”
    t = threading.Thread(target=run_task, args=(task_id, param_path, report_path, out_path), daemon=True)
    t.start()

    return jsonify(
        {
            "task_id": task_id,
            "download_url": url_for("download", task_id=task_id),
            "status_url": url_for("status", task_id=task_id),
            "log_url": url_for("get_log", task_id=task_id),
        }
    )


@app.route("/api/log/<task_id>", methods=["GET"])
def get_log(task_id):
    p = _log_path(task_id)
    if not os.path.exists(p):
        return jsonify({"logs": []})
    with open(p, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()
    # 只返回最后 400 行，防止日志太大拖慢页面
    return jsonify({"logs": lines[-400:]})


@app.route("/api/status/<task_id>", methods=["GET"])
def status(task_id):
    s = read_status(task_id)
    if not s:
        return jsonify({"state": "unknown"}), 404

    # 双保险：文件存在就视为 done
    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")
    if os.path.exists(out_path) and s.get("state") != "done":
        write_status(task_id, "done", None)
        s = read_status(task_id) or {"state": "done", "error": None}

    return jsonify(s)


@app.route("/api/download/<task_id>", methods=["GET"])
def download(task_id):
    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")
    if not os.path.exists(out_path):
        return "任务不存在或尚未生成结果", 404
    return send_file(out_path, as_attachment=True, download_name="最终结果.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)