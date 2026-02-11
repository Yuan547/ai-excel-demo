# -*- coding: utf-8 -*-

# processor.py
import os
import pandas as pd

def process_excel(param_path: str, report_path: str, out_path: str) -> None:
    """
    读取参数表 + 报表，生成 out_path
    先用你现有逻辑填充这里
    """
    # TODO: 把你现有代码中“读参数表、遍历sheet、写最终结果.xlsx”的核心逻辑搬进来
    # 临时：先写一个demo，确认流程通
    demo = pd.DataFrame([["测试网格", 1, 1, "100%", 10, 2, "20%", 99, "某旗县", "某产品"]],
                        columns=["单位名称","日目标","日发展","日发展完成率","月目标","月累计发展","月完成率","得分","旗县","产品"])
    demo.to_excel(out_path, index=False)
