import pandas as pd
import os

def process_excel(input_path, table_name, output_path=None, sheet_name=0):
    """
    自动化：读取Excel → 生成处理后的结果 → 可选导出
    
    参数:
    --------
    input_path : str
        输入Excel文件路径（支持相对/绝对路径）
    table_name : str
        需要筛选的表名
    output_path : str, optional
        导出Excel的路径，例如 'result.xlsx'；如果为None则不导出
    sheet_name : str/int, optional
        Excel工作表名或序号，默认0（第一个sheet）
    
    返回:
    --------
    doc_final : pd.DataFrame
        处理后的结果
    """
    
    # Step 1: 读取Excel
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"输入文件不存在: {input_path}")
    doc = pd.read_excel(input_path, sheet_name=sheet_name)

    # Step 2: 检查必要列
    required_cols = ['表名', '序号', '问题描述（必填）', '问题案例', '答复（必填）']
    missing_cols = [col for col in required_cols if col not in doc.columns]
    if missing_cols:
        raise ValueError(f"缺少必要列: {missing_cols}")

    # Step 3: 过滤表名
    doc_new = doc[doc['表名'] == table_name]
    if doc_new.empty:
        print(f"⚠️ 未找到表名 '{table_name}' 的数据")
        return pd.DataFrame()

    # Step 4: 聚合
    doc_new = doc_new.groupby(['表名', '序号'])[required_cols].first().reset_index(drop=True)
    doc_new = doc_new.fillna('')

    # Step 5: 组合列
    doc_new['答疑筛查[监督]'] = (
        '[问题序号]: ' + doc_new['序号'].astype(str) + '\n' +
        '[问题描述]: ' + doc_new['问题描述（必填）'] + '\n' +
        '[问题案例]: ' + doc_new['问题案例'] + '\n' +
        '[同业建议方案]: \n' +
        '[监管答复]: ' + doc_new['答复（必填）']
    )

    doc_final = doc_new[['表名', '序号', '答疑筛查[监督]']]

    # Step 6: 导出
    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc_final.to_excel(output_path, index=False, sheet_name=table_name)
        print(f"✅ 已导出到: {output_path}")

    return doc_final


# 用法示例
doc_final = process_excel(
    input_path="监管口径答疑文档_v1.0.xlsx",
    table_name="对公信贷业务借据表",
    output_path="结果/输出结果.xlsx"
)
