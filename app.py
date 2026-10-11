import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# ===== 页面配置 =====
st.set_page_config(page_title="淘宝利润计算器", page_icon="💰", layout="wide")
st.title("💰 淘宝利润计算器（网页版·多店铺 + 图表）")
st.markdown("上传订单、售后数据和商品成本配置表，点击“开始计算”即可得出利润结果。支持一个文件多个 sheet，每个 sheet 名 = 店铺名。")

# ===== 初始化 Session State =====
if 'config_df' not in st.session_state:
    st.session_state.config_df = pd.DataFrame(columns=['商品标题', '商品成本', '履约费', '税率'])
if 'result' not in st.session_state:
    st.session_state.result = None
if 'shop_summary' not in st.session_state:
    st.session_state.shop_summary = None
if 'orders' not in st.session_state:
    st.session_state.orders = None
if 'total_profit' not in st.session_state:
    st.session_state.total_profit = 0
if 'total_margin' not in st.session_state:
    st.session_state.total_margin = 0
if 'params_df' not in st.session_state:
    st.session_state.params_df = pd.DataFrame({
        '参数名称': ['软件费', '默认履约费', '默认税率(%)'],
        '参数值': [15.0, 1.5, 1.0]
    })

# ===== 智能读取（单个 sheet / 单文件） =====
def smart_read(file):
    if file.name.lower().endswith('.csv'):
        for enc in ['utf-8-sig', 'gbk', 'gb18030', 'utf-8']:
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError(f"无法识别文件编码：{file.name}")
    else:
        try:
            file.seek(0)
            return pd.read_excel(file)
        except Exception:
            file.seek(0)
            return pd.read_csv(file, encoding='gbk')

# ===== 智能读取所有 sheet，返回 {店铺名: DataFrame} =====
def smart_read_all_sheets(file):
    fname = file.name
    if fname.lower().endswith('.csv'):
        df = smart_read(file)
        shop_name = os.path.splitext(fname)[0]
        return {shop_name: df}
    else:
        try:
            file.seek(0)
            sheets = pd.read_excel(file, sheet_name=None)
        except Exception:
            file.seek(0)
            df = pd.read_csv(file, encoding='gbk')
            shop_name = os.path.splitext(fname)[0]
            return {shop_name: df}

        result = {}
        for name, df in sheets.items():
            if df is None or df.dropna(how='all').empty:
                continue
            df.columns = df.columns.astype(str).str.strip()
            result[str(name).strip()] = df
        if not result:
            raise ValueError(f"文件 {fname} 中没有找到有效数据 sheet。")
        return result

# ===== 侧边栏：全局参数（可增删改的表） =====
with st.sidebar:
    st.header("⚙️ 全局参数设置")
    st.caption("双击单元格可编辑；点行尾可删除；点表尾可添加新行。")
    edited_params = st.data_editor(
        st.session_state.params_df,
        num_rows="dynamic",
        width='stretch',
        hide_index=True,
        column_config={
            "参数名称": st.column_config.TextColumn("参数名称", width="medium"),
            "参数值": st.column_config.NumberColumn("参数值", width="small", format="%.2f"),
        },
        key="params_editor"
    )
    st.caption("💡 除「软件费」「默认履约费」「默认税率(%)」外，新增的参数将作为 **每单固定费用** 扣除（例如：包装费 0.5）。")
    st.caption("💡 多店铺时，软件费只从总利润里扣 **一次**。")

# 解析侧边栏参数
def parse_params(edited_params):
    soft_fee = 0.0
    default_fee = 0.0
    default_tax = 0.0
    extra_fees = {}
    for _, row in edited_params.iterrows():
        name = str(row['参数名称']).strip() if pd.notna(row['参数名称']) else ''
        if not name:
            continue
        try:
            value = float(row['参数值'])
        except (ValueError, TypeError):
            value = 0.0
        if name == '软件费':
            soft_fee = value
        elif name == '默认履约费':
            default_fee = value
        elif name in ('默认税率(%)', '默认税率'):
            default_tax = value / 100.0
        else:
            extra_fees[name] = value
    return soft_fee, default_fee, default_tax, extra_fees

soft_fee, default_fee, default_tax, extra_fees = parse_params(edited_params)

# ===== 文件上传 =====
st.markdown("---")
st.subheader("📁 数据上传")
st.markdown("订单明细表支持一个文件多个 sheet，**每个 sheet 名就是店铺名**。csv 文件会用文件名作为店铺名。")

col1, col2 = st.columns(2)
with col1:
    orders_file = st.file_uploader("📄 上传订单明细表（可多 sheet）", type=['xls', 'xlsx', 'csv'], key="orders_up")
with col2:
    afters_file = st.file_uploader("📄 上传售后数据（可多 sheet）", type=['xls', 'xlsx', 'csv'], key="afters_up")

st.markdown("---")
st.subheader("🛠️ 商品成本配置表")
st.markdown("上传商品成本配置表，需包含列：**商品标题、商品成本**（履约费、税率可留空，留空则使用左侧默认值）。配置表对所有店铺通用。")
config_file = st.file_uploader("📄 上传商品成本配置表", type=['xls', 'xlsx', 'csv'], key="config_up")

if config_file is not None:
    try:
        config_df = smart_read(config_file)
        config_df.columns = config_df.columns.astype(str).str.strip()
        if '商品标题' not in config_df.columns or '商品成本' not in config_df.columns:
            st.error("❌ 配置表必须包含「商品标题」和「商品成本」两列！")
        else:
            st.success(f"✅ 配置表已加载，共 {len(config_df)} 条商品")
            st.dataframe(config_df, width='stretch')
            st.session_state.config_df = config_df
    except Exception as e:
        st.error(f"❌ 配置表读取失败：{e}")

# 模板下载
st.markdown("---")
st.markdown("**第一次用？先下载配置表模板，按模板填写商品信息后再上传：**")
template_df = pd.DataFrame({
    '商品标题': ['示例商品A', '示例商品B'],
    '商品成本': [10.0, 20.0],
    '履约费': [1.5, 1.5],
    '税率': [1.0, 1.0]
})
template_buf = io.BytesIO()
template_df.to_excel(template_buf, index=False, engine='openpyxl')
template_buf.seek(0)
st.download_button(
    label="📥 下载配置表模板",
    data=template_buf,
    file_name="商品成本配置表模板.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ===== 读取多 sheet 并合并 =====
def load_multi_sheet(file, label):
    sheets = smart_read_all_sheets(file)
    frames = []
    shop_names = []
    for shop, df in sheets.items():
        if df is None or df.empty:
            continue
        df = df.copy()
        df.columns = df.columns.astype(str).str.strip()
        df['店铺'] = shop
        frames.append(df)
        shop_names.append(shop)
    if not frames:
        raise ValueError(f"{label} 中没有读取到有效数据。")
    merged = pd.concat(frames, ignore_index=True, sort=False)
    return merged, shop_names

# ===== 核心计算逻辑 =====
def calculate(orders, afters, config_df, soft_fee, default_fee, default_tax, extra_fees):
    config = config_df.copy()
    if '商品成本' not in config.columns:
        config['商品成本'] = 0
    if '履约费' not in config.columns:
        config['履约费'] = default_fee
    if '税率' not in config.columns:
        config['税率'] = default_tax

    config['商品成本'] = pd.to_numeric(config['商品成本'], errors='coerce').fillna(0)
    config['履约费'] = pd.to_numeric(config['履约费'], errors='coerce').fillna(default_fee)
    config['税率'] = pd.to_numeric(config['税率'], errors='coerce').fillna(default_tax)
    config['税率'] = config['税率'].apply(lambda x: x / 100 if x > 1 else x)

    cost_map = dict(zip(config['商品标题'], config['商品成本']))
    fee_map = dict(zip(config['商品标题'], config['履约费']))
    tax_map = dict(zip(config['商品标题'], config['税率']))

    orders = orders.copy()
    afters = afters.copy()

    if '店铺' not in orders.columns:
        orders['店铺'] = '默认店铺'
    if '店铺' not in afters.columns:
        afters['店铺'] = None

    orders['订单编号'] = orders['订单编号'].astype(str).str.strip()
    afters['订单编号'] = afters['订单编号'].astype(str).str.strip()

    # 兼容各种金额列名
    possible_amount_cols = ['买家应付货款', '总金额', '订单金额', '实付金额', '买家实付金额', '应收金额', '买家实际支付金额']
    found_col = None
    for col in possible_amount_cols:
        if col in orders.columns:
            found_col = col
            break

    if found_col:
        orders['计算基数'] = pd.to_numeric(orders[found_col], errors='coerce').fillna(0)
    else:
        st.error(f"❌ 订单明细表中找不到金额列！请确保表头包含以下任一列名：{', '.join(possible_amount_cols)}")
        st.stop()

    orders['退款金额'] = pd.to_numeric(orders['退款金额'], errors='coerce').fillna(0)
    orders['宝贝总数量'] = pd.to_numeric(orders['宝贝总数量'], errors='coerce').fillna(0)

    # ===== 兼容日期列名，用于每日利润曲线 =====
    possible_date_cols = ['下单时间', '成交时间', '订单创建时间', '付款时间', '订单付款时间', '创建时间', '订单日期', '日期']
    found_date = None
    for col in possible_date_cols:
        if col in orders.columns:
            found_date = col
            break
    if found_date:
        orders['订单日期'] = pd.to_datetime(orders[found_date], errors='coerce')
    else:
        orders['订单日期'] = pd.NaT

    valid_products = set(config['商品标题'].dropna())
    missing_products = set(orders['商品标题'].dropna()) - valid_products
    if missing_products:
        st.warning(f"⚠️ 以下商品未在配置表中，已按成本0、履约费{default_fee}、税率{default_tax*100}% 暂估：")
        st.write(list(missing_products))
        for p in missing_products:
            cost_map[p] = 0
            fee_map[p] = default_fee
            tax_map[p] = default_tax

    orders = orders[orders['计算基数'] > 0].copy()

    if '发货时间' in orders.columns:
        orders['是否已发货'] = orders['发货时间'].notna() & (orders['发货时间'].astype(str).str.strip() != '')
    else:
        orders['是否已发货'] = orders['订单状态'].isin(['卖家已发货，等待买家确认', '交易成功'])

    orders['商品成本'] = orders['商品标题'].map(cost_map).fillna(0) * orders['宝贝总数量']
    orders['退款后实收'] = orders['计算基数'] - orders['退款金额']
    orders['履约费'] = orders['商品标题'].map(fee_map).fillna(default_fee)
    orders['履约成本'] = orders['是否已发货'].astype(int) * orders['履约费']
    orders['税率'] = orders['商品标题'].map(tax_map).fillna(default_tax)
    
    # 【修改点】税费按照退款后实收计算，且最低为0（避免退款比原价还高导致负数税费）
    orders['税费'] = orders['退款后实收'].clip(lower=0) * orders['税率']

    # 额外费用（每单固定扣，A 方案）
    extra_fee_per_order = sum(extra_fees.values()) if extra_fees else 0.0
    orders['额外费用'] = extra_fee_per_order
    for fee_name, fee_val in extra_fees.items():
        orders[fee_name] = fee_val

    # 售后数据：过滤「出账」
    if '操作类型' in afters.columns:
        afters = afters[afters['操作类型'] == '出账'].copy()
    afters['收支金额'] = pd.to_numeric(afters['收支金额(元)'], errors='coerce').fillna(0)
    afters['交易赔付金额'] = afters.apply(lambda row: abs(row['收支金额']) if '交易赔付' in str(row['原因']) else 0, axis=1)
    afters['平台垫资运费'] = afters.apply(lambda row: abs(row['收支金额']) if any(k in str(row['原因']) for k in ['欠费划扣', '平台垫资欠费', '运费垫资欠费']) else 0, axis=1)

    afters_has_shop = afters['店铺'].notna().any() and (afters['店铺'].astype(str).str.strip() != '').any()
    if afters_has_shop:
        after_summary = afters.groupby(['店铺', '订单编号']).agg({'交易赔付金额': 'sum', '平台垫资运费': 'sum'}).reset_index()
        orders = orders.merge(after_summary, on=['店铺', '订单编号'], how='left')
    else:
        after_summary = afters.groupby('订单编号').agg({'交易赔付金额': 'sum', '平台垫资运费': 'sum'}).reset_index()
        orders = orders.merge(after_summary, on='订单编号', how='left')

    orders['交易赔付金额'] = orders['交易赔付金额'].fillna(0)
    orders['平台垫资运费'] = orders['平台垫资运费'].fillna(0)
    orders['退货运费'] = orders['平台垫资运费']

    orders['净利润'] = (orders['退款后实收'] - orders['商品成本'] - orders['履约成本']
                        - orders['税费'] - orders['交易赔付金额'] - orders['退货运费']
                        - orders['额外费用'])

    # ===== 商品维度汇总 =====
    result = orders.groupby(['店铺', '商品标题', '商品属性SKU']).agg({
        '订单编号': 'count', '宝贝总数量': 'sum', '计算基数': 'sum',
        '退款金额': 'sum', '退款后实收': 'sum', '商品成本': 'sum',
        '履约成本': 'sum', '税费': 'sum', '交易赔付金额': 'sum',
        '退货运费': 'sum', '额外费用': 'sum', '净利润': 'sum'
    }).reset_index()

    result = result.rename(columns={
        '计算基数': '买家应付货款', '订单编号': '订单数', '宝贝总数量': '总销量',
        '退款金额': '退款总额', '商品成本': '商品总成本', '履约成本': '履约总成本',
        '税费': '综合扣费', '交易赔付金额': '交易赔付', '退货运费': '平台垫资运费'
    })

    result['净利率'] = result['净利润'] / result['买家应付货款'].replace(0, np.nan)
    result['净利润'] = result['净利润'].round(2)
    result['退款率'] = (result['退款总额'] / result['买家应付货款'].replace(0, np.nan) * 100).round(2).fillna(0).astype(str) + '%'
    result['净利率'] = (result['净利率'] * 100).round(2).fillna(0).astype(str) + '%'
    result = result.sort_values('净利润', ascending=False)

    # ===== 店铺维度汇总 =====
    shop_summary = orders.groupby('店铺').agg({
        '订单编号': 'count', '宝贝总数量': 'sum', '计算基数': 'sum',
        '退款金额': 'sum', '退款后实收': 'sum', '商品成本': 'sum',
        '履约成本': 'sum', '税费': 'sum', '交易赔付金额': 'sum',
        '退货运费': 'sum', '额外费用': 'sum', '净利润': 'sum'
    }).reset_index()

    shop_summary = shop_summary.rename(columns={
        '计算基数': '买家应付货款', '订单编号': '订单数', '宝贝总数量': '总销量',
        '退款金额': '退款总额', '商品成本': '商品总成本', '履约成本': '履约总成本',
        '税费': '综合扣费', '交易赔付金额': '交易赔付', '退货运费': '平台垫资运费'
    })

    shop_summary['净利率'] = shop_summary['净利润'] / shop_summary['买家应付货款'].replace(0, np.nan)
    shop_summary['净利润'] = shop_summary['净利润'].round(2)
    shop_summary['退款率'] = (shop_summary['退款总额'] / shop_summary['买家应付货款'].replace(0, np.nan) * 100).round(2).fillna(0).astype(str) + '%'
    shop_summary['净利率'] = (shop_summary['净利率'] * 100).round(2).fillna(0).astype(str) + '%'
    shop_summary = shop_summary.sort_values('净利润', ascending=False)

    total_profit = orders['净利润'].sum() - soft_fee
    total_revenue = orders['退款后实收'].sum()
    total_margin = (total_profit / total_revenue * 100) if total_revenue != 0 else 0

    return result, shop_summary, orders, total_profit, total_margin

# ===== 触发计算 =====
if st.button("🚀 开始计算", type="primary"):
    if not orders_file or not afters_file:
        st.error("❌ 请先上传订单明细表和售后数据。")
    else:
        try:
            with st.spinner("正在读取多 sheet 并计算中..."):
                orders, order_shops = load_multi_sheet(orders_file, "订单明细表")
                afters, after_shops = load_multi_sheet(afters_file, "售后数据")

                st.info(f"📦 识别到订单表店铺：{', '.join(order_shops)}")
                if after_shops:
                    st.info(f"🧾 识别到售后表店铺：{', '.join(after_shops)}")

                result, shop_summary, calculated_orders, total_profit, total_margin = calculate(
                    orders, afters, st.session_state.config_df,
                    soft_fee, default_fee, default_tax, extra_fees
                )
                st.session_state.result = result
                st.session_state.shop_summary = shop_summary
                st.session_state.orders = calculated_orders
                st.session_state.total_profit = total_profit
                st.session_state.total_margin = total_margin
                st.rerun()
        except Exception as e:
            st.error(f"❌ 计算出错：{e}")
            st.exception(e)

# ===== 显示结果 =====
if st.session_state.result is not None:
    st.markdown("---")
    st.subheader("📊 计算结果")
    col_a, col_b = st.columns(2)
    col_a.metric("总净利润（已扣软件费，全部店铺合计）", f"￥{st.session_state.total_profit:.2f}")
    col_b.metric("总净利率", f"{st.session_state.total_margin:.2f}%")

    st.markdown("### 🏬 按店铺汇总")
    st.dataframe(st.session_state.shop_summary, width='stretch')

    st.markdown("### 🛍️ 按店铺 + 商品汇总")
    st.dataframe(st.session_state.result, width='stretch')

    # ===== 图表分析区 =====
    st.markdown("---")
    st.subheader("📈 图表分析")

    orders_df = st.session_state.orders
    result_df = st.session_state.result
    shop_df = st.session_state.shop_summary

    tab1, tab2, tab3 = st.tabs(["📅 每日利润曲线", "🏆 商品利润排行", "💸 店铺退款率"])

    # ----- Tab 1：每日利润曲线 -----
    with tab1:
        if '订单日期' in orders_df.columns and orders_df['订单日期'].notna().any():
            daily = orders_df.dropna(subset=['订单日期']).copy()
            daily['日期'] = daily['订单日期'].dt.date
            daily_profit = daily.groupby(['日期', '店铺'])['净利润'].sum().reset_index()
            pivot = daily_profit.pivot(index='日期', columns='店铺', values='净利润').fillna(0)
            pivot = pivot.sort_index()
            st.markdown("**每天各店铺的净利润走势**（多条线代表多个店铺）")
            st.line_chart(pivot)
            st.caption("提示：鼠标放到线上会显示具体金额。")
        else:
            st.info("ℹ️ 未识别到日期列，暂时无法画每日利润曲线。支持的日期列名："
                    + "、".join(['下单时间', '成交时间', '订单创建时间', '付款时间', '订单付款时间', '创建时间', '订单日期', '日期'])
                    + "。你可以把订单表里的日期列改名成上面任意一个再重新上传。")

    # ----- Tab 2：商品利润排行 -----
    with tab2:
        st.markdown("**Top N 商品净利润排行**（按店铺 + 商品）")
        top_n = st.slider("显示前 N 个商品", min_value=5, max_value=50, value=15, step=5, key="top_n")
        top = result_df.head(top_n).copy()
        top['净利润_num'] = pd.to_numeric(top['净利润'], errors='coerce').fillna(0)
        top['标签'] = top['店铺'].astype(str) + " | " + top['商品标题'].astype(str)
        chart_df = top.set_index('标签')[['净利润_num']].rename(columns={'净利润_num': '净利润'})
        # 降序，使利润最高的显示在最左边
        chart_df = chart_df.sort_values('净利润', ascending=False)
        st.bar_chart(chart_df)
        st.caption("横条越长代表这个商品净利润越高；如果是负数，说明这个商品在亏钱。从左到右利润由高到低排列。")

    # ----- Tab 3：店铺退款率 -----
    with tab3:
        st.markdown("**各店铺退款率**（退款总额 ÷ 买家应付货款）")
        refund = orders_df.groupby('店铺').agg({
            '计算基数': 'sum', '退款金额': 'sum'
        }).reset_index()
        refund['退款率(%)'] = (refund['退款金额'] / refund['计算基数'].replace(0, np.nan) * 100).round(2).fillna(0)
        refund = refund.sort_values('退款率(%)', ascending=False)
        st.bar_chart(refund.set_index('店铺')[['退款率(%)']])
        st.dataframe(refund[['店铺', '计算基数', '退款金额', '退款率(%)']].rename(
            columns={'计算基数': '买家应付货款', '退款金额': '退款总额'}
        ), width='stretch')

    # ===== 导出 Excel =====
    def to_excel_with_style(df_result, df_shop, df_orders):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_shop.to_excel(writer, index=False, sheet_name='店铺汇总')
            df_result.to_excel(writer, index=False, sheet_name='商品汇总')
            df_orders.to_excel(writer, index=False, sheet_name='每单明细')
        wb = load_workbook(output)

        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

        for sheet_name in ['店铺汇总', '商品汇总']:
            ws = wb[sheet_name]
            header = [cell.value for cell in ws[1]]
            if '净利润' not in header or '净利率' not in header:
                continue
            profit_col = header.index('净利润') + 1
            margin_col = header.index('净利率') + 1
            for row in range(2, ws.max_row + 1):
                profit = ws.cell(row=row, column=profit_col).value
                margin_str = ws.cell(row=row, column=margin_col).value
                margin = 0
                if isinstance(margin_str, str) and margin_str.endswith('%'):
                    try:
                        margin = float(margin_str.strip('%')) / 100
                    except:
                        margin = 0
                if profit is not None:
                    if profit < 0:
                        for col in range(1, ws.max_column + 1):
                            ws.cell(row=row, column=col).fill = red_fill
                    elif margin < 0.05:
                        for col in range(1, ws.max_column + 1):
                            ws.cell(row=row, column=col).fill = yellow_fill

        ws = wb['商品汇总']
        for idx, col in enumerate(df_result.columns):
            max_len = max(df_result[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[get_column_letter(idx + 1)].width = max_len

        wb.save(output)
        output.seek(0)
        return output

    excel_data = to_excel_with_style(
        st.session_state.result,
        st.session_state.shop_summary,
        st.session_state.orders
    )
    st.download_button(
        label="📥 下载带预警的 Excel 报表",
        data=excel_data,
        file_name="多店铺利润结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
