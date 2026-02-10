import streamlit as st
import os
import sys
import time
import ssl

# ==========================================
# 🔍 终极调试探针 (新增)
# ==========================================
st.write("### 🔍 系统路径自检")
current_dir = os.getcwd()
st.write(f"📂 当前工作目录: `{current_dir}`")

st.write("📂 目录下的文件清单:")
files = os.listdir(current_dir)
st.code(str(files))

services_path = os.path.join(current_dir, 'services')
if os.path.exists(services_path):
    st.success("✅ 找到了 services 文件夹！")
    st.write("📂 services 文件夹里的内容:")
    st.code(str(os.listdir(services_path)))
else:
    st.error(f"❌ 找不到 services 文件夹！它应该在: {services_path}")
    # 尝试递归查找
    st.write("🕵️‍♀️ 正在尝试全盘搜索 services...")
    found = False
    for root, dirs, files in os.walk(current_dir):
        if 'services' in dirs:
            found_path = os.path.join(root, 'services')
            st.warning(f"⚠️ 找到了！但是它藏在这里: `{found_path}`")
            sys.path.append(root) # 自动修复路径
            st.info(f"🔧 已自动修正路径，尝试重新加载...")
            found = True
            break
    if not found:
        st.error("☠️ 全盘搜索失败，services 文件夹真的不在代码库里。")
        st.stop()

# ==========================================
# 0. 前端环境 SSL 修复
# ==========================================
# ... (后面接原来的代码)
try:
    from services.model_engine import create_model, fetch_data
except ImportError:
    st.error("❌ 无法导入后端引擎，请确保 'services/model_engine.py' 文件存在且路径正确。")
    st.stop()

# --- 页面配置 ---
st.set_page_config(
    page_title="DeepInsight | 智能投研平台",
    page_icon="📊",
    layout="centered"
)

# --- 侧边栏 ---
with st.sidebar:
    st.title("DeepInsight V15")
    st.caption("全量标准版")
    st.markdown("---")
    st.markdown("### 🛠️ 模型能力")
    st.info("✅ 历史财报极致还原")
    st.info("✅ 业务/成本多维拆分")
    st.info("✅ 资产负债表自动配平")
    st.info("✅ 现金流量表间接法")
    st.markdown("---")
    st.markdown("Created by AI Industry Agent")

# --- 主界面 ---
st.title("🚀 A股上市公司估值建模系统")
st.markdown("输入股票代码，一键生成 **华尔街标准 (Standardized)** 的 Excel 财务模型。")

# 输入区域
with st.container():
    col1, col2 = st.columns([3, 1])
    with col1:
        symbol = st.text_input("股票代码", value="000895", placeholder="例如: 000895, 600519")
    with col2:
        st.write("") 
        st.write("") 
        run_btn = st.button("🚀 开始建模", type="primary", use_container_width=True)

# --- 逻辑处理 ---
if run_btn:
    if not symbol:
        st.warning("请输入有效的股票代码")
    else:
        # 初始化状态
        status_box = st.status("正在连接交易所数据中心...", expanded=True)
        
        try:
            # 1. 获取预览数据 (用于前端展示)
            status_box.write(f"🔍 正在抓取 {symbol} 的核心财务数据...")
            data_pool, years = fetch_data(symbol)
            
            if not data_pool:
                status_box.update(label="❌ 数据获取失败", state="error")
                st.error(f"无法获取代码 {symbol} 的数据，请检查代码是否正确（如：000895）。")
            else:
                # 2. 调用引擎生成 Excel
                status_box.write("⚙️ 正在构建三张报表勾稽关系...")
                create_model(symbol) # 核心生成步骤
                
                # 3. 检查文件是否生成
                file_prefix = "SZ" if symbol.startswith("0") or symbol.startswith("3") else "SH"
                if symbol.lower().startswith("sz") or symbol.lower().startswith("sh"):
                    file_prefix = "" # 如果用户自己输了前缀
                    
                filename = f"generated_models/{file_prefix}{symbol}_DeepInsight_V15_Standard.xlsx"
                # 简单的模糊查找，防止前缀大小写问题
                if not os.path.exists(filename):
                    # 尝试找一下目录下包含该代码的文件
                    import glob
                    files = glob.glob(f"generated_models/*{symbol}*V15*.xlsx")
                    if files:
                        filename = files[0]

                if os.path.exists(filename):
                    status_box.update(label="✅ 建模完成！", state="complete", expanded=False)
                    
                    # --- 结果展示区 ---
                    st.divider()
                    st.success(f"🎉 **{symbol} 估值模型已生成**")
                    
                    # 核心指标卡片
                    latest_year = years[-1]
                    latest_data = data_pool[latest_year]
                    
                    st.subheader(f"📊 核心指标预览 ({latest_year})")
                    k1, k2, k3 = st.columns(3)
                    
                    rev = latest_data.get('TOTAL_OPERATE_INCOME', 0)
                    profit = latest_data.get('PARENT_NETPROFIT', 0)
                    cash = latest_data.get('NETCASH_OPERATE', 0)
                    
                    k1.metric("营业总收入", f"{rev/1e8:,.2f} 亿")
                    k2.metric("归母净利润", f"{profit/1e8:,.2f} 亿", delta_color="normal")
                    k3.metric("经营性现金流", f"{cash/1e8:,.2f} 亿")

                    # 下载按钮
                    with open(filename, "rb") as file:
                        st.download_button(
                            label="📥 点击下载 Excel 估值模型 (.xlsx)",
                            data=file,
                            file_name=os.path.basename(filename),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                else:
                    status_box.update(label="❌ 生成失败", state="error")
                    st.error("模型文件未生成，请检查后端日志。")

        except Exception as e:
            status_box.update(label="❌ 发生系统错误", state="error")
            st.error(f"Error: {e}")
            st.code(traceback.format_exc())
