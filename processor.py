# -*- coding: utf-8 -*-
import os
import re
import json
import ast
from typing import Any, Optional, List

import pandas as pd
from openai import OpenAI
from openpyxl import Workbook


def read_excel_sheet(file_path: str, sheet_name: str, skiprows: int = 1, nrows: int = None, usecols: str = None):
    """读取Excel指定sheet的函数"""
    return pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        skiprows=skiprows,
        nrows=nrows,
        usecols=usecols
    )


def _parse_llm_list_of_lists(raw_text: str) -> List[List[Any]]:
    """
    尽量把模型返回解析为 Python 的 list[list]
    兼容：
    - JSON 格式
    - Python list 字符串（用 literal_eval）
    """
    if raw_text is None:
        raise ValueError("模型返回为空")

    text = raw_text.strip()

    # 常见：模型输出被包进 ``` ```，先去掉
    if text.startswith("```"):
        text = text.strip("`").strip()

    # 先尝试 JSON
    try:
        return json.loads(text)
    except Exception:
        pass

    # 再尝试 Python list
    try:
        return ast.literal_eval(text)
    except Exception:
        pass

    # 最后尝试你队友那种“去转义再 loads”
    try:
        clean_text = text.replace('\\"', '"').replace('"\"', '"')
        return json.loads(clean_text)
    except Exception as e:
        raise ValueError(f"无法解析模型输出为二维列表：{e}\n原始输出开头：{text[:200]}")


def analyze_data_with_llm(
    df: Any,
    api_key: Optional[str] = None,
    base_url: str = "https://dashscope.aliyuncs.com/compatible-mode/v1",
    model: str = "qwen3-max"
) -> List[List[Any]]:
    """
    调用通义千问（OpenAI 兼容接口）做映射清洗，返回 list[list]
    """
    if api_key is None:
        api_key = os.getenv("DASHSCOPE_API_KEY")
    if not api_key:
        raise RuntimeError("未检测到环境变量 DASHSCOPE_API_KEY（请在 PythonAnywhere Web->Environment Variables 配置）")

    client = OpenAI(api_key=api_key, base_url=base_url)

    # 重要：不要把 df 直接 f-string 成超长文本（会爆 token）
    # 这里保留必要信息：列名 + 前几行（如果你确定数据量不大，也可以改成完整 df.to_csv）
    preview = df.head(30).to_csv(index=False)

    completion = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"""
新建一个对话，你是一个专业的数据分析助手，请严格按照以下要求处理数据：
数据来源（CSV预览，含表头+部分数据）：\n{preview}

这个表前面行是表头，后面是数据，要从前面映射各产品，获取后面对应的数据。
一、数据清洗与预处理：
1. 数据筛选规则：
   - 不处理单位名称包含"未划分"、"合计"、"战客"的数据
   - 没有网格的：以"旗县"字段作为单位名称
   - 有网格的：以"网格"字段作为单位名称，但"旗县"仍需要输出为单独列
2. 指标映射规则（关键部分 - 严格匹配顺序）：
   需要从数据表中提取以下指标，按优先级顺序映射：
   | 目标字段 | 数据中的可能字段名（按匹配优先级降序） |
   |----------|-------------------------------------|
   | 日目标 | 日目标、日发展目标、当日目标、本日目标、日发展 |
   | 日发展 | 日发展、当日发展、本日发展、日新增 |
   | 日发展完成率 | 日完成率、日发展完成率、当日完成率；必须包含"日"且不包含"月" |
   | 月目标 | 月目标、月度目标、本月目标 |
   | 月累计发展 | 月累计发展、月度累计、累计发展、月发展 |
   | 月完成率 | 月完成率、月度完成率、累计完成率；必须包含"月"或"累计"且不包含"日"

重要映射规则：
- 先按优先级顺序匹配字段名
- 字段名必须包含指定的关键词（如"日"或"月"），避免交叉混淆
- 如果没有找到对应字段，该指标输出为"-1.5"

二、输出格式要求：
- 只返回一个合法的 JSON-like 二维列表（Python list of lists）
- 以 "[[" 开始，以 "]]" 结束，中间无额外内容
- 不要表头（只输出数据行）
- 每行顺序严格为：["单位名称","日目标","日发展","日发展完成率","月目标","月累计发展","月完成率","得分","旗县"]
- 数值保留原始精度，不要四舍五入
"""}]
    )

    content = completion.choices[0].message.content
    return _parse_llm_list_of_lists(content)


def process_excel(param_path: str, report_path: str, out_path: str) -> None:
    """
    网站入口：读取参数表 + 报表，调用 AI 做映射，写出 out_path
    """
    # 1) 读取参数表，判断是否“简版”
    df_param_all = pd.read_excel(param_path)
    shifou = df_param_all.iloc[0, 1] if df_param_all.shape[1] > 1 else 0

    numbers_list = []
    letters_list = []
    canshu_list = []

    if shifou == 0:
        # 按你队友逻辑：A:C + nrows=4 + skiprows=1
        canshu = pd.read_excel(param_path, usecols="A:C", nrows=4, skiprows=1)
        canshu_list = canshu.values.tolist()

        # 解析第2、3列的行列范围
        for _, row in canshu.iterrows():
            row_numbers = []
            row_letters = []
            for col_index in [1, 2]:
                item = row.iloc[col_index]
                if pd.isna(item):
                    numbers = ''
                    letters = ''
                else:
                    item_str = str(item)
                    numbers = ''.join(re.findall(r'\d+', item_str))
                    letters = ''.join(re.findall(r'[a-zA-Z]+', item_str))
                row_numbers.append(numbers)
                row_letters.append(letters)
            numbers_list.append(row_numbers)
            letters_list.append(row_letters)
    else:
        # 如果不是简版：默认处理报表所有 sheet
        xls_report = pd.ExcelFile(report_path)
        canshu_list = [[s] for s in xls_report.sheet_names]

    # 2) 获取报表 sheet
    xls = pd.ExcelFile(report_path)
    sheet_names = xls.sheet_names

    # 3) 写出 Excel（openpyxl）
    wb = Workbook()
    ws1 = wb.active

    all_headers = ["单位名称", "日目标", "日发展", "日发展完成率",
                   "月目标", "月累计发展", "月完成率", "得分", "旗县", "产品", "排名", "备注"]
    for col, header in enumerate(all_headers, start=1):
        ws1.cell(row=1, column=col, value=header)

    product_col_index = all_headers.index("产品") + 1
    hangshu = 2

    # 4) 遍历参数表指定的 sheet
    for i in range(len(canshu_list)):
        target_sheet = canshu_list[i][0]
        if target_sheet not in sheet_names:
            continue

        # 读取 df
        if shifou == 0:
            # 简版：按范围读
            number = int(numbers_list[i][1]) - int(numbers_list[i][0]) - 4
            letter1 = letters_list[i][0]
            letter2 = letters_list[i][1]
            range_str = f"{letter1}:{letter2}"
            df = read_excel_sheet(
                file_path=report_path,
                sheet_name=target_sheet,
                skiprows=4,
                nrows=number,
                usecols=range_str,
            )
        else:
            df = read_excel_sheet(
                file_path=report_path,
                sheet_name=target_sheet,
                skiprows=1
            )

        # 5) 调 AI（返回二维列表，不含表头）
        all_data = analyze_data_with_llm(df)

        # 6) 写入结果
        for row_data in all_data:
            for col_idx, value in enumerate(row_data, 1):
                ws1.cell(row=hangshu, column=col_idx, value=value)
            ws1.cell(row=hangshu, column=product_col_index, value=target_sheet)
            hangshu += 1

    # 7) 保存到 out_path（网站要求）
    wb.save(out_path)