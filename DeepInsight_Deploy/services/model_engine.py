import akshare as ak
import pandas as pd
import xlsxwriter
from io import BytesIO
import datetime
import sys
import traceback
import ssl

# ==========================================
# 0. SSL 证书验证绕过 (环境兼容补丁)
# ==========================================
try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    pass
else:
    ssl._create_default_https_context = _create_unverified_https_context

# ==========================================
# 1. 审计级全量科目库 (不再进行删减)
# ==========================================
FULL_SCHEMA = {
    "IS": [ # 利润表
        ("一、营业总收入", "TOTAL_OPERATE_INCOME", 0, True),
        ("    其中：营业收入", "OPERATE_INCOME", 1, False),
        ("二、营业总成本", "TOTAL_OPERATE_COST", 0, True),
        ("    其中：营业成本", "OPERATE_COST", 1, False),
        ("        税金及附加", "TAX_BUSINESSSURCHARGE", 1, False),
        ("        销售费用", "SALE_EXPENSE", 1, False),
        ("        管理费用", "MANAGE_EXPENSE", 1, False),
        ("        研发费用", "RESEARCH_EXPENSE", 1, False),
        ("        财务费用", "FINANCE_EXPENSE", 1, False),
        ("            其中：利息费用", "INTEREST_EXPENSE", 2, False),
        ("            利息收入", "INTEREST_INCOME", 2, False),
        ("    加：其他收益", "OTHER_INCOME", 1, False),
        ("        投资收益", "INVEST_INCOME", 1, False),
        ("        公允价值变动收益", "FAIRVALUE_CHANGE_INCOME", 1, False),
        ("        信用减值损失", "CREDIT_IMPAIRMENT_LOSS", 1, False),
        ("        资产减值损失", "ASSET_IMPAIRMENT_LOSS", 1, False),
        ("        资产处置收益", "ASSET_DISPOSAL_INCOME", 1, False),
        ("三、营业利润", "OPERATE_PROFIT", 0, True),
        ("    加：营业外收入", "NONBUSINESS_INCOME", 1, False),
        ("    减：营业外支出", "NONBUSINESS_EXPENSE", 1, False),
        ("四、利润总额", "TOTAL_PROFIT", 0, True),
        ("    减：所得税费用", "INCOME_TAX", 1, False),
        ("五、净利润", "NETPROFIT", 0, True),
        ("    归母净利润", "PARENT_NETPROFIT", 1, True),
        ("    少数股东损益", "MINORITY_INTEREST", 1, False),
        ("六、EBITDA (参考)", "EBITDA_CALC", 0, True)
    ],
    "BS": [ # 资产负债表 (强制全量显示)
        ("流动资产：", "", 0, True),
        ("    货币资金", "MONETARYFUNDS", 1, False),
        ("    交易性金融资产", "TRADE_FINASSET_NOTFVTPL", 1, False),
        ("    应收票据", "NOTES_RECE", 1, False),
        ("    应收账款", "ACCOUNTS_RECE", 1, False),
        ("    应收款项融资", "RECEIVABLE_FINANCING", 1, False),
        ("    预付款项", "PREPAYMENT", 1, False),
        ("    其他应收款", "OTHER_RECE", 1, False),
        ("    存货", "INVENTORY", 1, False),
        ("    合同资产", "CONTRACT_ASSET", 1, False),
        ("    一年内到期的非流动资产", "NONCURRENT_ASSET_ONE_YEAR", 1, False),
        ("    其他流动资产", "OTHER_CURRENT_ASSET", 1, False),
        ("  流动资产合计", "TOTAL_CURRENT_ASSETS", 0, True),
        ("非流动资产：", "", 0, True),
        ("    长期股权投资", "LONG_EQUITY_INVEST", 1, False),
        ("    其他权益工具投资", "OTHER_EQUITY_INVEST", 1, False),
        ("    投资性房地产", "INVEST_REALESTATE", 1, False),
        ("    固定资产", "FIXED_ASSET", 1, False),
        ("    在建工程", "CONSTRUCTION_IN_PROCESS", 1, False),
        ("    使用权资产", "RIGHT_USE_ASSETS", 1, False),
        ("    无形资产", "INTANGIBLE_ASSET", 1, False),
        ("    商誉", "GOODWILL", 1, False),
        ("    长期待摊费用", "LONG_PREPAID_EXPENSE", 1, False),
        ("    递延所得税资产", "DEFERRED_TAX_ASSET", 1, False),
        ("    其他非流动资产", "OTHER_NONCURRENT_ASSET", 1, False),
        ("  非流动资产合计", "TOTAL_NONCURRENT_ASSETS", 0, True),
        ("资产总计", "TOTAL_ASSETS", 0, True),
        ("流动负债：", "", 0, True),
        ("    短期借款", "SHORT_LOAN", 1, False),
        ("    应付票据", "NOTES_PAYABLE", 1, False),
        ("    应付账款", "ACCOUNTS_PAYABLE", 1, False),
        ("    预收款项", "PRECEIVE", 1, False),
        ("    合同负债", "CONTRACT_LIABILITIES", 1, False),
        ("    应付职工薪酬", "PAYROLL_PAYABLE", 1, False),
        ("    应交税费", "TAX_PAYABLE", 1, False),
        ("    其他应付款", "OTHER_PAYABLE", 1, False),
        ("    一年内到期的非流动负债", "NONCURRENT_LIAB_ONE_YEAR", 1, False),
        ("    其他流动负债", "OTHER_CURRENT_LIAB", 1, False),
        ("  流动负债合计", "TOTAL_CURRENT_LIAB", 0, True),
        ("非流动负债：", "", 0, True),
        ("    长期借款", "LONG_LOAN", 1, False),
        ("    应付债券", "BOND_PAYABLE", 1, False),
        ("    租赁负债", "LEASE_LIAB", 1, False),
        ("    长期应付款", "LONG_PAYABLE", 1, False),
        ("    递延收益", "DEFERRED_REVENUE", 1, False),
        ("    递延所得税负债", "DEFERRED_TAX_LIAB", 1, False),
        ("    预计负债", "ANTICIPATE_LIAB", 1, False),
        ("    其他非流动负债", "OTHER_NONCURRENT_LIAB", 1, False),
        ("  非流动负债合计", "TOTAL_NONCURRENT_LIAB", 0, True),
        ("负债合计", "TOTAL_LIABILITIES", 0, True),
        ("股东权益：", "", 0, True),
        ("    实收资本(或股本)", "SHARE_CAPITAL", 1, False),
        ("    资本公积", "CAPITAL_RESERVE", 1, False),
        ("    盈余公积", "SURPLUS_RESERVE", 1, False),
        ("    未分配利润", "UNDISTRIBUTED_PROFIT", 1, False),
        ("  归属于母公司股东权益合计", "TOTAL_EQUITY", 0, True),
        ("    少数股东权益", "MINORITY_EQUITY", 1, False),
        ("  股东权益合计", "TOTAL_LIAB_EQUITY", 0, True),
        ("报表配平项 (Plug)", "BS_PLUG", 0, True),
        ("负债和股东权益总计", "TOTAL_LIABILITIES_AND_EQUITY_CALC", 0, True),
        ("CHECK (配平检查)", "BALANCE_CHECK", 0, True)
    ],
    "CF": [ # 现金流量表
        ("一、经营活动产生的现金流量：", "", 0, True),
        ("    销售商品、提供劳务收到的现金", "SALES_SERVICES", 1, False),
        ("    收到的税费返还", "RECEIVE_TAX_REFUND", 1, False),
        ("    收到其他与经营活动有关的现金", "RECEIVE_OTHER_OPERATE", 1, False),
        ("  经营活动现金流入小计", "TOTAL_OPERATE_INFLOW", 0, True),
        ("    购买商品、接受劳务支付的现金", "BUY_GOODS_SERVICES", 1, False),
        ("    支付给职工以及为职工支付的现金", "PAY_STAFF_CASH", 1, False),
        ("    支付的各项税费", "PAY_ALL_TAX", 1, False),
        ("    支付其他与经营活动有关的现金", "PAY_OTHER_OPERATE", 1, False),
        ("  经营活动现金流出小计", "TOTAL_OPERATE_OUTFLOW", 0, True),
        ("  经营活动产生的现金流量净额", "NETCASH_OPERATE", 0, True),
        ("二、投资活动产生的现金流量：", "", 0, True),
        ("    收回投资收到的现金", "WITHDRAW_INVEST", 1, False),
        ("    取得投资收益收到的现金", "INVEST_INCOME_CASH", 1, False),
        ("    处置固定资产、无形资产和其他长期资产收回的现金净额", "DISPOSAL_LONG_ASSET", 1, False),
        ("  投资活动现金流入小计", "TOTAL_INVEST_INFLOW", 0, True),
        ("    购建固定资产、无形资产和其他长期资产支付的现金", "CONSTRUCT_LONG_ASSET", 1, False),
        ("    投资支付的现金", "INVEST_PAY_CASH", 1, False),
        ("  投资活动现金流出小计", "TOTAL_INVEST_OUTFLOW", 0, True),
        ("  投资活动产生的现金流量净额", "NETCASH_INVEST", 0, True),
        ("三、筹资活动产生的现金流量：", "", 0, True),
        ("    吸收投资收到的现金", "ABSORB_INVEST_RECEIVED", 1, False),
        ("    取得借款收到的现金", "BORROW_CASH", 1, False),
        ("  筹资活动现金流入小计", "TOTAL_FINANCE_INFLOW", 0, True),
        ("    偿还债务支付的现金", "PAY_DEBT_CASH", 1, False),
        ("    分配股利、利润或偿付利息支付的现金", "ASSIGN_DIVIDEND_PORFIT", 1, False),
        ("  筹资活动现金流出小计", "TOTAL_FINANCE_OUTFLOW", 0, True),
        ("  筹资活动产生的现金流量净额", "NETCASH_FINANCE", 0, True),
        ("四、现金及现金等价物净增加额", "CASH_NETINCREASE", 0, True),
        ("五、期末现金及现金等价物余额", "YEAR_END_CASH", 0, True)
    ]
}

def normalize_key(key): return str(key).upper().strip()

# ==========================================
# 2. 数据获取
# ==========================================
def fetch_data(symbol):
    code = f"SZ{symbol}" if symbol.startswith("3") or symbol.startswith("0") else f"SH{symbol}"
    print(f"🚀 [DeepInsight V15.0] 启动全量标准版: {code}...")
    try:
        df_is = ak.stock_profit_sheet_by_report_em(symbol=code)
        df_bs = ak.stock_balance_sheet_by_report_em(symbol=code)
        df_cf = ak.stock_cash_flow_sheet_by_report_em(symbol=code)
        current_year = datetime.datetime.now().year
        target_years = [str(y) for y in range(current_year - 7, current_year)]
        data_pool = {}
        def process(df):
            if df is None or df.empty: return
            col_map = {col: normalize_key(col) for col in df.columns}
            for _, row in df.iterrows():
                r_date = str(row.get('REPORT_DATE') or row.get('report_date', ''))
                if "12-31" in r_date:
                    year = r_date[:4]
                    if year in target_years:
                        if year not in data_pool: data_pool[year] = {}
                        for col, val in row.items():
                            std_key = col_map.get(col, col)
                            try:
                                if val and str(val).replace('.', '', 1).replace('-', '', 1).isdigit():
                                    data_pool[year][std_key] = float(val)
                                else: data_pool[year][std_key] = 0.0
                            except: data_pool[year][std_key] = 0.0
        process(df_is); process(df_bs); process(df_cf)
        
        for y in data_pool:
            d = data_pool[y]
            if d.get("FE_INTEREST_EXPENSE", 0) == 0: d["FE_INTEREST_EXPENSE"] = d.get("FINANCE_EXPENSE", 0)
            if d.get("INTEREST_EXPENSE", 0) == 0: d["INTEREST_EXPENSE"] = d.get("FE_INTEREST_EXPENSE", 0)
            d["EBIT_CALC"] = d.get("TOTAL_PROFIT", 0) + d.get("FINANCE_EXPENSE", 0)
            d["EBITDA_CALC"] = d["EBIT_CALC"]
            total_liab = d.get("TOTAL_LIABILITIES", 0)
            total_eq = d.get("TOTAL_LIAB_EQUITY", 0) - total_liab
            if total_eq <= 0: total_eq = d.get("TOTAL_EQUITY", 0) + d.get("MINORITY_EQUITY", 0)
            d["TOTAL_LIABILITIES_AND_EQUITY_CALC"] = total_liab + total_eq
            d["BALANCE_CHECK"] = d.get("TOTAL_ASSETS", 0) - d["TOTAL_LIABILITIES_AND_EQUITY_CALC"]

        return data_pool, sorted(list(data_pool.keys()))
    except Exception as e: print(f"❌ 数据获取失败: {e}"); return None, None

# ==========================================
# 3. 模版构建
# ==========================================
def create_model(symbol):
    data_pool, years = fetch_data(symbol)
    if not data_pool: return None
    
    # ❌ 移除 get_active_schema，强制使用 FULL_SCHEMA
    schema = FULL_SCHEMA
    
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    
    font = 'Arial'
    st_title = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#003366', 'font_name': font})
    st_th_hist = wb.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#D9E1F2', 'font_name': font})
    st_th_proj = wb.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#FFF2CC', 'font_name': font})
    st_item0 = wb.add_format({'bold': True, 'indent': 0, 'font_name': font})
    st_item1 = wb.add_format({'indent': 2, 'font_name': font})
    st_item2 = wb.add_format({'indent': 4, 'font_color': '#666666', 'font_name': font})
    st_num_h = wb.add_format({'num_format': '#,##0', 'font_color': '#003366', 'font_name': font})
    st_num_f = wb.add_format({'num_format': '#,##0', 'font_color': 'black', 'font_name': font})
    st_pct_h = wb.add_format({'num_format': '0.0%', 'font_color': '#003366', 'font_name': font})
    st_inp = wb.add_format({'bg_color': '#FFFFCC', 'border': 1, 'font_color': 'blue', 'num_format': '#,##0'})
    st_inp_pct = wb.add_format({'bg_color': '#FFFFCC', 'border': 1, 'font_color': 'blue', 'num_format': '0.00%'})
    st_plug = wb.add_format({'bg_color': '#E6E6E6', 'font_color': 'red', 'bold': True, 'num_format': '#,##0'})
    
    last_y = int(years[-1])
    proj_years = [str(last_y + i) for i in range(1, 6)]
    all_years = years + proj_years
    
    ref_map = {s: {} for s in ['HIST', 'REV', 'ASSUMP', 'INV', 'FIN', 'WC', 'IS', 'BS']}

    # Sheet 1: History
    s1 = wb.add_worksheet("1.历史财务报表")
    s1.hide_gridlines(2); s1.set_column(0,0,50); s1.set_column(1, len(years)+1, 14); s1.set_tab_color('#336699')
    s1.write(0, 0, f"{symbol} 历史数据底稿 (Standardized)", st_title)
    s1.write(2, 0, "会计科目", st_th_hist)
    for i, y in enumerate(years): s1.write(2, i+1, y, st_th_hist)
    curr = 3
    for sht in ["IS", "BS", "CF"]:
        for cn, key, indent, bold in schema[sht]:
            fmt = st_item0 if bold else (st_item1 if indent==1 else st_item2)
            s1.write(curr, 0, cn, fmt)
            if key:
                for i, y in enumerate(years):
                    s1.write(curr, i+1, data_pool[y].get(key, 0)/1e6, st_num_h)
                ref_map['HIST'][key] = curr
            curr += 1
        curr += 1

    # Sheet 2: Assumptions
    s_assump = wb.add_worksheet("2.基本假设")
    s_assump.hide_gridlines(2); s_assump.set_column(0,0,45); s_assump.set_column(1, len(all_years)+1, 12); s_assump.set_tab_color('#FF0000')
    s_assump.write(0, 0, "核心驱动假设", st_title)
    for i, y in enumerate(all_years): s_assump.write(2, i+1, y, st_th_hist if i<len(years) else st_th_proj)
    curr = 3
    drivers = [
        ("REV_GROWTH", "营业收入增长率 (YoY)", "TOTAL_OPERATE_INCOME", "TOTAL_OPERATE_INCOME", 0.10, True),
        ("TAX_RATE_REV", "税金及附加率 (%收入)", "TAX_BUSINESSSURCHARGE", "TOTAL_OPERATE_INCOME", 0.005, False),
        ("SELL_RATE", "销售费用率 (%收入)", "SALE_EXPENSE", "TOTAL_OPERATE_INCOME", 0.04, False),
        ("MANAGE_RATE", "管理费用率 (%收入)", "MANAGE_EXPENSE", "TOTAL_OPERATE_INCOME", 0.03, False),
        ("RD_RATE", "研发费用率 (%收入)", "RESEARCH_EXPENSE", "TOTAL_OPERATE_INCOME", 0.02, False),
        ("INCOME_TAX_RATE", "有效所得税率 (%EBT)", "INCOME_TAX", "TOTAL_PROFIT", 0.15, False),
        ("DSO", "应收账款周转天数 (DSO)", "ACCOUNTS_RECE", "TOTAL_OPERATE_INCOME", 30, False),
        ("DIO", "存货周转天数 (DIO)", "INVENTORY", "OPERATE_COST", 60, False),
        ("DPO", "应付账款周转天数 (DPO)", "ACCOUNTS_PAYABLE", "OPERATE_COST", 60, False),
        ("CAPEX_RATE", "CAPEX占收入比", "CONSTRUCT_LONG_ASSET", "TOTAL_OPERATE_INCOME", 0.05, False),
        ("DIV_PAYOUT", "股利支付率 (%净利润)", "ASSIGN_DIVIDEND_PORFIT", "NETPROFIT", 0.30, False)
    ]
    for code, name, num_key, den_key, default, is_growth in drivers:
        s_assump.write(curr, 0, name, st_item1)
        ref_map['ASSUMP'][code] = curr
        for i, y in enumerate(years):
            col = xlsxwriter.utility.xl_col_to_name(i+1)
            if is_growth:
                if i == 0: s_assump.write(curr, i+1, 0, st_pct_h)
                else:
                    prev_col = xlsxwriter.utility.xl_col_to_name(i)
                    if num_key in ref_map['HIST']:
                        row_idx = ref_map['HIST'][num_key] + 1
                        s_assump.write_formula(curr, i+1, f"=('1.历史财务报表'!{col}{row_idx}/'1.历史财务报表'!{prev_col}{row_idx})-1", st_pct_h)
                    else: s_assump.write(curr, i+1, 0, st_pct_h)
            else:
                if num_key in ref_map['HIST'] and den_key in ref_map['HIST']:
                    num_row = ref_map['HIST'][num_key] + 1
                    den_row = ref_map['HIST'][den_key] + 1
                    formula = f"=IFERROR('1.历史财务报表'!{col}{num_row}/'1.历史财务报表'!{col}{den_row}, 0)"
                    if "周转天数" in name: formula += "*360"
                    fmt = st_num_h if "周转天数" in name else st_pct_h
                    s_assump.write_formula(curr, i+1, formula, fmt)
                else: s_assump.write(curr, i+1, 0, st_num_h)
        start_avg_col = xlsxwriter.utility.xl_col_to_name(max(1, len(years)-2))
        end_avg_col = xlsxwriter.utility.xl_col_to_name(len(years))
        for i, y in enumerate(proj_years):
            col_idx = len(years) + i + 1
            avg_formula = f"=AVERAGE({start_avg_col}{curr+1}:{end_avg_col}{curr+1})"
            fmt = st_inp if "周转天数" in name else st_inp_pct
            s_assump.write_formula(curr, col_idx, avg_formula, fmt)
        curr += 1
    s_assump.write(curr, 0, "综合折旧率 (%期初固定资产)", st_item1)
    for i in range(len(all_years)): s_assump.write(curr, i+1, 0.10, st_inp_pct if i>=len(years) else st_pct_h)
    ref_map['ASSUMP']['DEPR_RATE'] = curr; curr += 1
    s_assump.write(curr, 0, "平均债务利率", st_item1)
    for i in range(len(all_years)): s_assump.write(curr, i+1, 0.04, st_inp_pct if i>=len(years) else st_pct_h)
    ref_map['ASSUMP']['DEBT_RATE'] = curr; curr += 1

    # Sheet 3 (Revenue)
    s2 = wb.add_worksheet("3.业务拆分预测"); s2.hide_gridlines(2); s2.set_column(0,0,35); s2.set_column(1, len(all_years)+1, 13); s2.set_tab_color('#FF9900')
    s2.write(0,0,"业务量价与成本模型", st_title)
    for i, y in enumerate(all_years): s2.write(2, i+1, y, st_th_hist)
    curr=3; segments = ["核心业务A", "核心业务B", "其他业务"]
    total_rev_rows = []; total_cost_rows = []
    for seg in segments:
        s2.write(curr, 0, seg, st_item0); curr += 1
        s2.write(curr, 0, "  销量 (Vol)", st_item1); curr += 1
        s2.write(curr, 0, "  单价 (ASP)", st_item1); curr += 1
        s2.write(curr, 0, "  单位成本 (Unit Cost)", st_item1); curr += 1
        s2.write(curr, 0, f"  {seg}收入", st_item1)
        ratio = 0.6 if "A" in seg else 0.2
        for i, y in enumerate(all_years):
            col = xlsxwriter.utility.xl_col_to_name(i+1)
            if i < len(years):
                hist_ref = f"'1.历史财务报表'!{col}{ref_map['HIST']['TOTAL_OPERATE_INCOME']+1}"
                s2.write_formula(curr, i+1, f"={hist_ref}*{ratio}", st_num_h)
            else:
                prev=xlsxwriter.utility.xl_col_to_name(i); growth=f"'2.基本假设'!{col}{ref_map['ASSUMP']['REV_GROWTH']+1}"
                s2.write_formula(curr, i+1, f"={prev}{curr+1}*(1+{growth})", st_num_f)
        total_rev_rows.append(curr); curr += 1
        s2.write(curr, 0, f"  {seg}成本", st_item1)
        for i, y in enumerate(all_years):
            col = xlsxwriter.utility.xl_col_to_name(i+1)
            if i < len(years):
                hist_ref = f"'1.历史财务报表'!{col}{ref_map['HIST']['OPERATE_COST']+1}"
                s2.write_formula(curr, i+1, f"={hist_ref}*{ratio}", st_num_h)
            else: s2.write_formula(curr, i+1, f"={col}{curr}*0.75", st_num_f)
        total_cost_rows.append(curr); curr += 2
    s2.write(curr, 0, "营业总收入合计", st_item0)
    for i in range(len(all_years)):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        f = "=" + "+".join([f"{col}{r+1}" for r in total_rev_rows])
        s2.write_formula(curr, i+1, f, st_num_f)
    ref_map['REV']['TOTAL'] = curr; curr += 1
    s2.write(curr, 0, "营业总成本合计", st_item0)
    for i in range(len(all_years)):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        f = "=" + "+".join([f"{col}{r+1}" for r in total_cost_rows])
        s2.write_formula(curr, i+1, f, st_num_f)
    ref_map['REV']['COST'] = curr

    # Sheet 4-6 (Schedules)
    s4 = wb.add_worksheet("4.投资预测"); s4.hide_gridlines(2); s4.set_column(0,0,35); s4.set_column(1, len(all_years)+1, 13)
    s4.write(0,0,"CAPEX", st_title); curr=3
    s4.write(curr,0,"期初PPE", st_item1); beg=curr; curr+=1
    s4.write(curr,0,"CAPEX", st_item1); capex=curr; curr+=1
    s4.write(curr,0,"Depr", st_item1); da=curr; curr+=1
    s4.write(curr,0,"期末PPE", st_item0); end=curr
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1); prev = xlsxwriter.utility.xl_col_to_name(i)
        if i==0: s4.write(beg, i+1, 0, st_num_f)
        else: s4.write_formula(beg, i+1, f"={prev}{end+1}", st_num_f)
        if i<len(years): s4.write_formula(capex, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST'].get('CONSTRUCT_LONG_ASSET', 0)+1}", st_num_h)
        else: s4.write_formula(capex, i+1, f"='3.业务拆分预测'!{col}{ref_map['REV']['TOTAL']+1}*'2.基本假设'!{col}{ref_map['ASSUMP']['CAPEX_RATE']+1}", st_num_f)
        s4.write_formula(da, i+1, f"={col}{beg+1}*0.1", st_num_f)
        s4.write_formula(end, i+1, f"={col}{beg+1}+{col}{capex+1}-{col}{da+1}", st_num_f)
    ref_map['INV']['DA']=da; ref_map['INV']['PPE']=end; ref_map['INV']['CAPEX']=capex

    s5 = wb.add_worksheet("5.筹资预测"); s5.hide_gridlines(2); s5.set_column(0,0,35); s5.set_column(1, len(all_years)+1, 13)
    s5.write(0,0,"Debt", st_title); curr=3
    s5.write(curr,0,"Debt Bal", st_item0); debt=curr; curr+=1
    s5.write(curr,0,"Interest", st_item1); inte=curr
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1); prev=xlsxwriter.utility.xl_col_to_name(i)
        if i<len(years): s5.write_formula(debt, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST'].get('SHORT_LOAN',0)+1}", st_num_h)
        else: s5.write_formula(debt, i+1, f"={prev}{debt+1}", st_num_f)
        s5.write_formula(inte, i+1, f"={col}{debt+1}*0.04", st_num_f)
    ref_map['FIN']['DEBT']=debt; ref_map['FIN']['INT']=inte

    s6 = wb.add_worksheet("6.营运资金"); s6.hide_gridlines(2); s6.set_column(0,0,35); s6.set_column(1, len(all_years)+1, 13)
    s6.write(0,0,"WC", st_title); curr=3
    s6.write(curr,0,"AR", st_item0); ar=curr; curr+=1
    s6.write(curr,0,"Inv", st_item0); inv=curr; curr+=1
    s6.write(curr,0,"Change", st_item0); chg=curr
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1); prev=xlsxwriter.utility.xl_col_to_name(i)
        if i<len(years):
            s6.write_formula(ar, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST']['ACCOUNTS_RECE']+1}", st_num_h)
            s6.write_formula(inv, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST']['INVENTORY']+1}", st_num_h)
        else:
            s6.write_formula(ar, i+1, f"='3.业务拆分预测'!{col}{ref_map['REV']['TOTAL']+1}/360*30", st_num_f)
            s6.write_formula(inv, i+1, f"='3.业务拆分预测'!{col}{ref_map['REV']['COST']+1}/360*60", st_num_f)
        if i==0: s6.write(chg, i+1, 0, st_num_f)
        else: s6.write_formula(chg, i+1, f"=-({col}{ar+1}-{prev}{ar+1} + {col}{inv+1}-{prev}{inv+1})", st_num_f)
    ref_map['WC']['AR']=ar; ref_map['WC']['INV']=inv; ref_map['WC']['CHG']=chg

    # Sheet 7: IS
    s7 = wb.add_worksheet("7.利润表预测"); s7.hide_gridlines(2); s7.set_column(0,0,45)
    s7.write(0,0,"IS", st_title); curr=3
    s7.write(curr, 0, "营业总收入", st_item0)
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        s7.write_formula(curr, i+1, f"='3.业务拆分预测'!{col}{ref_map['REV']['TOTAL']+1}", st_num_f)
    curr += 1
    s7.write(curr, 0, "营业成本", st_item0)
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        s7.write_formula(curr, i+1, f"='3.业务拆分预测'!{col}{ref_map['REV']['COST']+1}", st_num_f)
    curr += 1
    for cn, key, indent, bold in schema['IS']:
        if key in ["TOTAL_OPERATE_INCOME", "OPERATE_COST", "OPERATE_INCOME", "TOTAL_OPERATE_COST"]: continue
        if key == "": continue
        fmt = st_item0 if bold else (st_item1 if indent==1 else st_item2)
        s7.write(curr, 0, cn, fmt)
        for i, y in enumerate(all_years):
            col = xlsxwriter.utility.xl_col_to_name(i+1)
            if "INTEREST" in key and i >= len(years):
                 s7.write_formula(curr, i+1, f"='5.筹资预测'!{col}{ref_map['FIN']['INT']+1}", st_num_f)
            elif i < len(years):
                if key in ref_map['HIST']: s7.write_formula(curr, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST'][key]+1}", st_num_h)
            else:
                rev = f"='3.业务拆分预测'!{col}{ref_map['REV']['TOTAL']+1}"
                if "TAX" in key: s7.write_formula(curr, i+1, f"={rev}*'2.基本假设'!{col}{ref_map['ASSUMP']['TAX_RATE_REV']+1}", st_num_f)
                elif "SALE" in key: s7.write_formula(curr, i+1, f"={rev}*'2.基本假设'!{col}{ref_map['ASSUMP']['SELL_RATE']+1}", st_num_f)
                elif "MANAGE" in key: s7.write_formula(curr, i+1, f"={rev}*'2.基本假设'!{col}{ref_map['ASSUMP']['MANAGE_RATE']+1}", st_num_f)
                else: s7.write_formula(curr, i+1, f"={rev}*0.01", st_num_f)
        ref_map['IS'][key] = curr; curr += 1

    # Sheet 8: BS
    s8 = wb.add_worksheet("8.资产负债表预测"); s8.hide_gridlines(2); s8.set_column(0,0,45)
    s8.write(0,0,"BS", st_title); curr=3
    asset_rows=[]; liab_rows=[]; equity_rows=[]; plug_row=-1
    s8.write(curr, 0, "货币资金", st_item1)
    bs_row_idx={'MONETARYFUNDS': curr}; asset_rows.append(curr); cash_row=curr
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        if i < len(years): s8.write_formula(curr, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST']['MONETARYFUNDS']+1}", st_num_h)
        else: s8.write_formula(curr, i+1, f"='9.现金流量表预测'!{col}25", st_num_f)
    curr+=1
    for cn, key, indent, bold in schema['BS']:
        if key in ["MONETARYFUNDS", "TOTAL_ASSETS", "BALANCE_CHECK"] or key.startswith("TOTAL_") or key == "": continue
        if key == "BS_PLUG": plug_row = curr; s8.write(curr, 0, "报表配平项 (Plug)", st_plug); curr += 1; continue
        fmt = st_item0 if bold else (st_item1 if indent==1 else st_item2)
        s8.write(curr, 0, cn, fmt)
        if "ASSET" in key or "RECE" in key or "INVENTORY" in key or "INVEST" in key: asset_rows.append(curr)
        elif "CAPITAL" in key or "PROFIT" in key or "EQUITY" in key or "RESERVE" in key: equity_rows.append(curr)
        else: liab_rows.append(curr)
        for i, y in enumerate(all_years):
            col = xlsxwriter.utility.xl_col_to_name(i+1)
            if key == "ACCOUNTS_RECE": s8.write_formula(curr, i+1, f"='6.营运资金'!{col}{ref_map['WC']['AR']+1}", st_num_f)
            elif key == "INVENTORY": s8.write_formula(curr, i+1, f"='6.营运资金'!{col}{ref_map['WC']['INV']+1}", st_num_f)
            elif key == "FIXED_ASSET": s8.write_formula(curr, i+1, f"='4.投资预测'!{col}{ref_map['INV']['PPE']+1}", st_num_f)
            elif key == "SHORT_LOAN": s8.write_formula(curr, i+1, f"='5.筹资预测'!{col}{ref_map['FIN']['DEBT']+1}", st_num_f)
            elif key == "UNDISTRIBUTED_PROFIT" and i >= len(years):
                prev = xlsxwriter.utility.xl_col_to_name(i)
                ni = f"='7.利润表预测'!{col}{ref_map['IS']['NETPROFIT']+1}"
                div = f"'2.基本假设'!{col}{ref_map['ASSUMP']['DIV_PAYOUT']+1}"
                s8.write_formula(curr, i+1, f"={prev}{curr+1} + {ni}*(1-{div})", st_num_f)
            elif i < len(years) and key in ref_map['HIST']:
                s8.write_formula(curr, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST'][key]+1}", st_num_h)
            else:
                prev = xlsxwriter.utility.xl_col_to_name(i)
                s8.write_formula(curr, i+1, f"={prev}{curr+1}", st_num_f)
        curr += 1
    s8.write(curr, 0, "资产总计", st_item0); asset_total_row = curr
    for i in range(len(all_years)):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        f_sum = "+".join([f"{col}{r+1}" for r in asset_rows])
        s8.write_formula(curr, i+1, f"={f_sum}", st_num_f)
    curr += 2
    s8.write(curr, 0, "负债权益合计", st_item0)
    for i in range(len(all_years)):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        f_sum_l = "+".join([f"{col}{r+1}" for r in liab_rows])
        f_sum_e = "+".join([f"{col}{r+1}" for r in equity_rows])
        s8.write_formula(curr, i+1, f"={f_sum_l}+{f_sum_e}+{col}{plug_row+1}", st_num_f)
    for i in range(len(all_years)):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        f_sum_l = "+".join([f"{col}{r+1}" for r in liab_rows])
        f_sum_e = "+".join([f"{col}{r+1}" for r in equity_rows])
        s8.write_formula(plug_row, i+1, f"={col}{asset_total_row+1} - ({f_sum_l}+{f_sum_e})", st_plug)

    # Sheet 9: CF
    s9 = wb.add_worksheet("9.现金流量表预测"); s9.hide_gridlines(2); s9.set_column(0,0,45)
    s9.write(0,0,"CF (Indirect)", st_title); curr=3
    s9.write(curr, 0, "一、经营活动", st_item0); curr += 1
    s9.write(curr, 0, "净利润", st_item1)
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        if i < len(years): s9.write_formula(curr, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST']['NETPROFIT']+1}", st_num_h)
        else: s9.write_formula(curr, i+1, f"='7.利润表预测'!{col}{ref_map['IS']['NETPROFIT']+1}", st_num_f)
    curr += 1
    s9.write(curr, 0, "加: 折旧摊销", st_item1)
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        if i < len(years): 
            cfo = f"='1.历史财务报表'!{col}{ref_map['HIST']['NETCASH_OPERATE']+1}"
            ni = f"='1.历史财务报表'!{col}{ref_map['HIST']['NETPROFIT']+1}"
            s9.write_formula(curr, i+1, f"={cfo}-{ni}", st_num_h) 
        else: s9.write_formula(curr, i+1, f"='4.投资预测'!{col}{ref_map['INV']['DA']+1}", st_num_f)
    curr += 1
    s9.write(curr, 0, "加: 营运资金变动", st_item1)
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        if i < len(years): s9.write(curr, i+1, 0, st_num_h)
        else: s9.write_formula(curr, i+1, f"='6.营运资金'!{col}{ref_map['WC']['CHG']+1}", st_num_f)
    curr += 1
    s9.write(curr, 0, "经营活动现金流净额", st_item0); cfo=curr
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        if i < len(years): s9.write_formula(curr, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST']['NETCASH_OPERATE']+1}", st_num_h)
        else: s9.write_formula(curr, i+1, f"=SUM({col}{curr-3}:{col}{curr-1})", st_num_f)
    curr += 2
    s9.write(curr, 0, "二、投资活动", st_item0); curr += 1
    s9.write(curr, 0, "CAPEX", st_item1)
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        if i < len(years): s9.write_formula(curr, i+1, f"=-'1.历史财务报表'!{col}{ref_map['HIST'].get('CONSTRUCT_LONG_ASSET', 0)+1}", st_num_h)
        else: s9.write_formula(curr, i+1, f"=-'4.投资预测'!{col}{ref_map['INV']['CAPEX']+1}", st_num_f)
    cfi=curr; curr += 2
    s9.write(curr, 0, "三、筹资活动", st_item0); curr += 1
    s9.write(curr, 0, "债务变动", st_item1); curr += 1
    s9.write(curr, 0, "股利", st_item1)
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        if i < len(years): s9.write_formula(curr, i+1, f"=-'1.历史财务报表'!{col}{ref_map['HIST'].get('ASSIGN_DIVIDEND_PORFIT',0)+1}", st_num_h)
        else:
            ni = f"='7.利润表预测'!{col}{ref_map['IS']['NETPROFIT']+1}"
            rate = f"'2.基本假设'!{col}{ref_map['ASSUMP']['DIV_PAYOUT']+1}"
            s9.write_formula(curr, i+1, f"=-{ni}*{rate}", st_num_f)
    cff=curr; curr += 2
    s9.write(curr, 0, "现金净增加额", st_item0); net_chg=curr
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1)
        if i < len(years): s9.write_formula(curr, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST']['CASH_NETINCREASE']+1}", st_num_h)
        else: s9.write_formula(curr, i+1, f"={col}{cfo+1}+{col}{cfi+1}+{col}{cff+1}", st_num_f)
    curr += 1
    s9.write(curr, 0, "期初现金", st_item1); beg_c=curr; curr+=1
    s9.write(curr, 0, "期末现金", st_item0); end_c=curr
    for i, y in enumerate(all_years):
        col = xlsxwriter.utility.xl_col_to_name(i+1); prev = xlsxwriter.utility.xl_col_to_name(i)
        if i==0: s9.write_formula(beg_c, i+1, f"='1.历史财务报表'!{col}{ref_map['HIST']['MONETARYFUNDS']+1}", st_num_h)
        else: s9.write_formula(beg_c, i+1, f"={prev}{end_c+1}", st_num_f)
        s9.write_formula(end_c, i+1, f"={col}{beg_c+1}+{col}{net_chg+1}", st_num_f)
    s8.write_formula(cash_row, len(years)+1, f"='9.现金流量表预测'!{xlsxwriter.utility.xl_col_to_name(len(years))}{end_c+1}", st_num_f)

    wb.close()
    output.seek(0)
    
    import os
    if not os.path.exists("generated_models"): os.makedirs("generated_models")
    filename = f"generated_models/{symbol}_DeepInsight_V15_Standard.xlsx"
    with open(filename, "wb") as f: f.write(output.read())
    print(f"✅ [V15.0] 标准化全量模型已生成: {filename}")

if __name__ == "__main__":
    symbol = sys.argv[1] if len(sys.argv) > 1 else "000895"
    create_model(symbol)