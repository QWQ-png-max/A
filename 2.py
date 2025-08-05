# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import logging
import sys
import os
import webbrowser

# 配置日志
log_path = os.path.join(os.path.expanduser("~"), "material_processor.log")
logging.basicConfig(
    filename=log_path,
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8"
)

# 检查 Streamlit 版本
if "streamlit" not in sys.modules:
    logging.error("Streamlit 未安装")
    st.error("Streamlit 未安装")
    sys.exit(1)

# 防止重复运行
if "app_initialized" not in st.session_state:
    st.session_state.app_initialized = False

if st.session_state.app_initialized:
    logging.debug("应用已初始化，跳过重复运行")
    st.stop()
else:
    st.session_state.app_initialized = True
    try:
        st.set_page_config(page_title="Material Processor")
        logging.debug("页面配置成功")
    except Exception as e:
        logging.error(f"页面配置失败: {e}", exc_info=True)
        st.error(f"页面配置失败: {e}")
        sys.exit(1)

# 初始化 session_state
if "conditions_path" not in st.session_state:
    st.session_state.conditions_path = None
if "database_path" not in st.session_state:
    st.session_state.database_path = None
if "output_path" not in st.session_state:
    st.session_state.output_path = ""
if "production_qty" not in st.session_state:
    st.session_state.production_qty = 1

# 添加 CSS 样式
st.markdown(
    """
    <style>
    div.stButton > button {
        width: 120px;
        height: 50px;
        font-size: 16px;
        font-weight: bold;
        color: black;
        background-color: #FFD700;
        border: none;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
        transition: all 0.3s ease;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 10px auto;
    }
    div.stButton > button:hover {
        background-color: #4682B4;
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.3);
        transform: translateY(-2px);
    }
    div.stButton > button:active {
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
        transform: translateY(0);
    }
    </style>
    """,
    unsafe_allow_html=True
)

def process_new_material_codes():
    if st.session_state.conditions_path is None or st.session_state.database_path is None:
        st.error("请上传所有所需文件！")
        logging.error("缺少上传文件")
        return
    if not st.session_state.output_path:
        st.error("请输入有效的保存路径！")
        logging.error("缺少保存路径")
        return
    try:
        conditions = pd.read_excel(st.session_state.conditions_path)
        database = pd.read_excel(st.session_state.database_path)
        output_path = Path(st.session_state.output_path)
        required_conditions_cols = ["原物料代码", "新编码"]
        required_database_cols = ["编码", "新系统编码"]
        if not all(col in conditions.columns for col in required_conditions_cols):
            st.error(f"设备物料清单缺少以下必备列：{', '.join(required_conditions_cols)}")
            logging.error(f"设备物料清单缺少列：{', '.join(required_conditions_cols)}")
            return
        if not all(col in database.columns for col in required_database_cols):
            st.error(f"旧物料表格缺少以下必备列：{', '.join(required_database_cols)}")
            logging.error(f"旧物料表格缺少列：{', '.join(required_database_cols)}")
            return
        conditions["新编码"] = conditions["新编码"].astype("string")
        for index, row in conditions.iterrows():
            condition_code = row["原物料代码"]
            match = database[database["编码"] == condition_code]
            if not match.empty:
                conditions.at[index, "新编码"] = match["新系统编码"].iloc[0]
            else:
                conditions.at[index, "新编码"] = ""
        conditions.to_excel(output_path, index=False)
        st.success(f"完成！已成功将新物料代码同步到 {output_path}")
        logging.debug(f"成功同步新物料代码到 {output_path}")
    except Exception as e:
        logging.error(f"处理新物料编码失败: {e}", exc_info=True)
        st.error(f"处理失败：{e}")

def process_inventory():
    if st.session_state.conditions_path is None or st.session_state.database_path is None:
        st.error("请上传所有所需文件！")
        logging.error("缺少上传文件")
        return
    if not st.session_state.output_path:
        st.error("请输入有效的保存路径！")
        logging.error("缺少保存路径")
        return
    try:
        conditions = pd.read_excel(st.session_state.conditions_path)
        database = pd.read_excel(st.session_state.database_path)
        output_path = Path(st.session_state.output_path)
        required_conditions_cols = ["新编码", "库存"]
        required_database_cols = ["物料代码", "基本计量单位数量"]
        if not all(col in conditions.columns for col in required_conditions_cols):
            st.error(f"设备物料清单缺少以下必备列：{', '.join(required_conditions_cols)}")
            logging.error(f"设备物料清单缺少列：{', '.join(required_conditions_cols)}")
            return
        if not all(col in database.columns for col in required_database_cols):
            st.error(f"库存表格缺少以下必备列：{', '.join(required_database_cols)}")
            logging.error(f"库存表格缺少列：{', '.join(required_database_cols)}")
            return
        conditions_columns = "\n".join(f"第{i+1}列：{col}" for i, col in enumerate(conditions.columns))
        database_columns = "\n".join(f"第{i+1}列：{col}" for i, col in enumerate(database.columns))
        st.info(f"设备物料清单列名：\n{conditions_columns}\n\n库存表格列名：\n{database_columns}")
        logging.debug(f"设备物料清单列名：\n{conditions_columns}\n库存表格列名：\n{database_columns}")
        conditions["库存"] = conditions["库存"].astype("string")
        for index, row in conditions.iterrows():
            condition_code = row["新编码"]
            match = database[database["物料代码"] == condition_code]
            if not match.empty:
                conditions.at[index, "库存"] = match["基本计量单位数量"].iloc[0]
            else:
                conditions.at[index, "库存"] = "0"
        conditions.to_excel(output_path, index=False)
        st.success(f"完成！已成功将库存同步到 {output_path}")
        logging.debug(f"成功同步库存到 {output_path}")
    except Exception as e:
        logging.error(f"处理库存失败: {e}", exc_info=True)
        st.error(f"处理失败：{e}")

def generate_purchase_list():
    if st.session_state.conditions_path is None:
        st.error("请上传设备物料清单！")
        logging.error("缺少设备物料清单")
        return
    if not st.session_state.output_path:
        st.error("请输入有效的保存路径！")
        logging.error("缺少保存路径")
        return
    try:
        conditions = pd.read_excel(st.session_state.conditions_path)
        output_path = Path(st.session_state.output_path)
        production_qty = st.session_state.production_qty
        required_cols = ["数量", "库存", "参考材料单价"]
        if not all(col in conditions.columns for col in required_cols):
            st.error(f"设备物料清单缺少以下必备列：{', '.join(required_cols)}")
            logging.error(f"设备物料清单缺少列：{', '.join(required_cols)}")
            return
        for col in ["数量", "库存", "参考材料单价"]:
            conditions[col] = conditions[col].replace({",": ""}, regex=True)
        conditions["数量"] = pd.to_numeric(conditions["数量"], errors="coerce")
        conditions["库存"] = pd.to_numeric(conditions["库存"], errors="coerce")
        conditions["参考材料单价"] = pd.to_numeric(conditions["参考材料单价"], errors="coerce")
        conditions["缺口量"] = conditions["数量"] * production_qty - conditions["库存"]
        conditions["需求量"] = conditions["缺口量"].clip(lower=0)
        conditions["成本"] = conditions["需求量"] * conditions["参考材料单价"]
        total_cost = conditions["成本"].sum()
        purchase_list = conditions[conditions["需求量"] > 0][["物料名称", "需求量", "新编码", "参考材料单价", "成本"]]
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            conditions.to_excel(writer, sheet_name="物料报表", index=False)
            purchase_list.to_excel(writer, sheet_name="购买清单", index=False)
            summary = pd.DataFrame([["成本", total_cost]], columns=["项目", "金额"])
            summary.to_excel(writer, sheet_name="成本汇总", index=False)
        st.success(f"处理完成！报表已保存到 {output_path}\n总成本：{total_cost:.2f}")
        logging.debug(f"成功生成采购清单到 {output_path}，总成本：{total_cost:.2f}")
    except Exception as e:
        logging.error(f"生成采购清单失败: {e}", exc_info=True)
        st.error(f"处理失败：{e}")

def main():
    st.info("运行程序后，浏览器将自动打开")
    logging.debug("尝试打开浏览器")
    webbrowser.open("https://material-processor.streamlit.app")
    st.title("注意！在开始运行程序前，请确保表格中有以下列名：")
    st.write("采购清单中：新编码，原物料代码，数量，参考材料单价，库存")
    st.write("整理后新旧物料编码对照表：编码，新系统编码")
    st.write("库存：物料代码，基本计量单位数量")
    option = st.selectbox(
        "你想执行哪项任务？",
        ("同步新物料编码", "同步库存数量", "生成采购清单"),
        index=0,
        placeholder="任务"
    )
    st.subheader("文件上传")
    if option == "同步新物料编码":
        st.session_state.conditions_path = st.file_uploader("请选择设备物料清单：")
        st.session_state.database_path = st.file_uploader("请选择旧物料表格：")
        st.subheader("请选择保存路径")
        st.session_state.output_path = st.text_input("请输入文件保存路径：", placeholder="例如：/tmp/output.xlsx")
    elif option == "同步库存数量":
        st.session_state.conditions_path = st.file_uploader("请选择已同步新物料代码的设备物料清单：")
        st.session_state.database_path = st.file_uploader("请选择库存表格：")
        st.subheader("请选择保存路径")
        st.session_state.output_path = st.text_input("请输入文件保存路径：", placeholder="例如：/tmp/output.xlsx")
    elif option == "生成采购清单":
        st.session_state.conditions_path = st.file_uploader("请选择已同步库存的设备物料清单：")
        st.subheader("请输入生产设备数量")
        st.session_state.production_qty = st.number_input("生产设备数量：", min_value=1, value=1)
        st.subheader("请选择保存路径")
        st.session_state.output_path = st.text_input("请输入文件保存路径：", placeholder="例如：/tmp/output.xlsx")
    if st.button("开始运行"):
        logging.debug(f"选择任务：{option}")
        if option == "同步新物料编码":
            process_new_material_codes()
        elif option == "同步库存数量":
            process_inventory()
        elif option == "生成采购清单":
            generate_purchase_list()
        else:
            st.error("请正确选择任务后再执行应用")
            logging.error("未选择有效任务")
    if st.button("清除缓存"):
        st.session_state.clear()
        st.session_state.app_initialized = False
        st.success("状态已重置")
        logging.debug("状态已重置")

if __name__ == "__main__":
    try:
        logging.debug("开始运行 main()")
        main()
    except Exception as e:
        logging.error(f"主程序运行失败: {e}", exc_info=True)
        st.error(f"主程序运行失败: {e}")
