# -*- coding: utf-8 -*-

import os
import uuid
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 用一个很简单的“内存日志”，演示够用（比赛现场足够）
TASK_LOGS = {}  # task_id -> list[str]
TASK_OUTPUT = {}  # task_id -> output_filepath


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
    先做假处理：生成一个空的“最终结果.xlsx”，并返回 task_id
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
    add_log(task_id, "开始处理（当前为假流程：仅演示网页跑通）…")

    # 保存上传文件
    param_path = os.path.join(UPLOAD_DIR, f"{task_id}_param.xlsx")
    report_path = os.path.join(UPLOAD_DIR, f"{task_id}_report.xlsx")
    param_file.save(param_path)
    report_file.save(report_path)
    add_log(task_id, "文件已保存到服务器临时目录")

    # ====== 假处理：生成一个简单的结果文件（后续替换为你们真实逻辑）======
    import pandas as pd
    out_path = os.path.join(OUTPUT_DIR, f"{task_id}_最终结果.xlsx")

    demo = pd.DataFrame([
        ["A网格", 10, 3, "30%", 300, 120, "40%", 85, "某旗县", "产品1"],
        ["B网格", 8, 5, "62.5%", 200, 160, "80%", 92, "某旗县", "产品2"],
    ], columns=["单位名称","日目标","日发展","日发展完成率","月目标","月累计发展","月完成率","得分","旗县","产品"])

    demo.to_excel(out_path, index=False)
    add_log(task_id, "生成演示用 最终结果.xlsx 完成")

    TASK_OUTPUT[task_id] = out_path
    add_log(task_id, "处理结束，可以下载结果")

    return jsonify({"task_id": task_id})


@app.route("/api/log/<task_id>", methods=["GET"])
def get_log(task_id):
    return jsonify({"logs": TASK_LOGS.get(task_id, [])})


@app.route("/api/download/<task_id>", methods=["GET"])
def download(task_id):
    if task_id not in TASK_OUTPUT:
        return "任务不存在或尚未生成结果", 404
    path = TASK_OUTPUT[task_id]
    return send_file(path, as_attachment=True, download_name="最终结果.xlsx")


if __name__ == "__main__":
    # 本地运行：访问 http://127.0.0.1:5000
    app.run(host="0.0.0.0", port=5000, debug=True)
