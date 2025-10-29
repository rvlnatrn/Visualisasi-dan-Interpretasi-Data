import pandas as pd
import plotly.express as px
import plotly.io as pio
from pathlib import Path
from datetime import datetime

DATA_PATH = r"D:\SEM 5\Visualisasi data\UTS\tokopedia.xlsx"
OUT_HTML = "index.html"

def load_excel(path: str, sheet_hint: str = "df") -> pd.DataFrame:
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheet = sheet_hint if sheet_hint in xls.sheet_names else xls.sheet_names[0]
    df = xls.parse(sheet)

    if "order_date" in df.columns:
        df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")
        df["year"]  = df["order_date"].dt.year
        df["month"] = df["order_date"].dt.month
        df["ym"]    = df["order_date"].dt.to_period("M").astype(str)

    for col in ["before_discount", "discount_amount", "after_discount", "price", "qty_ordered", "cogs"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in ["payment_method", "category", "sku_name", "customer_id", "region", "is_valid"]:
        if col in df.columns:
            df[col] = df[col].astype(str)
    return df

p = Path(DATA_PATH)
if not p.exists():
    raise FileNotFoundError(f"File tidak ditemukan: {p}\n→ Pastikan path benar atau letakkan file di folder proyek ini.")

df = load_excel(str(p))

print(f"Memuat: {p.resolve()}")
print("Kolom tersedia:", list(df.columns))

total_sales = float(df["after_discount"].sum()) if "after_discount" in df.columns else 0.0
total_orders = int(df.shape[0])
unique_customers = int(df["customer_id"].nunique()) if "customer_id" in df.columns else 0
aov = total_sales / max(total_orders, 1)

print("\n===== KPI =====")
print(f"Total Sales (After Discount): ${total_sales:,.0f}")
print(f"Total Orders               : {total_orders:,}")
print(f"Unique Customers           : {unique_customers:,}")
print(f"Average Order Value (AOV)  : ${aov:,.0f}")

px.defaults.template = "plotly_dark"
figs = {}

if {"ym", "after_discount"}.issubset(df.columns):
    trend = df.groupby("ym", as_index=False)["after_discount"].sum().sort_values("ym")
    fig_trend = px.line(trend, x="ym", y="after_discount", markers=True, title="Sales Trend by Month (after_discount)")
    fig_trend.update_layout(xaxis_title="", yaxis_title="Sales (after discount)")
    figs["trend"] = fig_trend
else:
    print("Lewati grafik tren: 'ym' atau 'after_discount' tidak ada.")

if {"category", "after_discount"}.issubset(df.columns):
    cat_sales = (
        df.groupby("category", as_index=False)["after_discount"]
          .sum().sort_values("after_discount", ascending=False).head(12)
    )
    fig_cat = px.bar(cat_sales, x="after_discount", y="category", orientation="h", title="Top Categories by Sales")
    fig_cat.update_layout(xaxis_title="Sales", yaxis_title="Category")
    figs["category"] = fig_cat
else:
    print("Lewati kategori: 'category' / 'after_discount' tidak ada.")

if {"payment_method", "after_discount"}.issubset(df.columns):
    pay = (df.groupby("payment_method", as_index=False)["after_discount"]
             .sum().sort_values("after_discount", ascending=False))
    fig_pay = px.pie(pay, names="payment_method", values="after_discount", title="Sales by Payment Method", hole=0.35)
    figs["payment"] = fig_pay
else:
    print("Lewati payment: 'payment_method' / 'after_discount' tidak ada.")

if {"sku_name", "after_discount"}.issubset(df.columns):
    top_prod = (df.groupby("sku_name", as_index=False)["after_discount"]
                  .sum().sort_values("after_discount", ascending=False).head(10))
    fig_tp = px.bar(top_prod, x="after_discount", y="sku_name", orientation="h", title="Top 10 Products by Sales")
    fig_tp.update_layout(xaxis_title="Sales", yaxis_title="Product")
    figs["top_products"] = fig_tp
else:
    print("Lewati Top Products: 'sku_name' / 'after_discount' tidak ada.")

if {"customer_id", "after_discount"}.issubset(df.columns):
    top_cust = (df.groupby("customer_id", as_index=False)["after_discount"]
                  .sum().sort_values("after_discount", ascending=False).head(10))
    fig_tc = px.bar(top_cust, x="after_discount", y="customer_id", orientation="h", title="Top 10 Customers by Sales")
    fig_tc.update_layout(xaxis_title="Sales", yaxis_title="Customer ID")
    figs["top_customers"] = fig_tc
else:
    print("Lewati Top Customers: 'customer_id' / 'after_discount' tidak ada.")

def fig_div(fig):
    return pio.to_html(fig, include_plotlyjs=False, full_html=False, config={"displaylogo": False})

kpi_html = f"""
<div class="kpi-grid">
  <div class="kpi"><div class="kpi-label">Total Sales</div><div class="kpi-value">${total_sales:,.0f}</div></div>
  <div class="kpi"><div class="kpi-label">Total Orders</div><div class="kpi-value">{total_orders:,}</div></div>
  <div class="kpi"><div class="kpi-label">Unique Customers</div><div class="kpi-value">{unique_customers:,}</div></div>
  <div class="kpi"><div class="kpi-label">AOV</div><div class="kpi-value">${aov:,.0f}</div></div>
</div>
"""

sections = [
    ("trend", "Tren Penjualan Bulanan"),
    ("category", "Top Kategori"),
    ("payment", "Porsi Metode Pembayaran"),
    ("top_products", "Top 10 Produk"),
    ("top_customers", "Top 10 Pelanggan"),
]

cards_html = ""
for key, title in sections:
    if key in figs:
        cards_html += f"""
        <section class="card">
          <h2>{title}</h2>
          {fig_div(figs[key])}
        </section>
        """

# ===================== FIX: escape {{ }} di blok .author =====================
html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>TOKOPEDIA</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
  :root {{
    --bg: #0f172a;
    --panel: #0b1229;
    --border: #1f2937;
    --text: #e5e7eb;
    --muted: #94a3b8;
    --accent: #93c5fd;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; background: var(--bg); color: var(--text);
    font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Ubuntu;
  }}
  .container {{ max-width: 1200px; margin: 0 auto; padding: 24px 16px; }}
  h1 {{ margin: 0 0 8px; font-size: 28px; }}
  .sub {{ color: var(--accent); margin-bottom: 20px; }}

  /* Blok author harus pakai kurung ganda di f-string */
  .author {{
    margin-top: 4px;
    margin-bottom: 12px;
    color: var(--accent);
    font-size: 14px;
    line-height: 1.5;
  }}
  .author p {{
    margin: 2px 0;
  }}

  .kpi-grid {{
    display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 18px;
  }}
  .kpi {{
    background: linear-gradient(135deg, #111827 0%, #0b1229 100%);
    border: 1px solid var(--border); border-radius: 16px; padding: 14px 16px;
    box-shadow: 0 10px 30px rgba(0,0,0,.25);
  }}
  .kpi-label {{ font-size: 12px; color: var(--accent); }}
  .kpi-value {{ font-size: 22px; font-weight: 700; color: #f8fafc; }}
  .card {{
    background: var(--panel); border: 1px solid var(--border); border-radius: 16px;
    padding: 16px; margin: 18px 0;
  }}
  h2 {{ margin:0 0 10px; font-size: 18px; }}
  footer {{ margin-top: 26px; color: var(--muted); font-size: 12px; }}
  @media (max-width: 900px) {{ .kpi-grid {{ grid-template-columns: repeat(2,1fr); }} }}
  @media (max-width: 600px) {{ .kpi-grid {{ grid-template-columns: 1fr; }} }}
</style>
</head>
<body>
  <div class="container">
    <h1>TOKOPEDIA</h1>
    <div class="author">
      <p><strong>Nama:</strong> Triana Revana Sirumpa</p>
      <p><strong>NPM:</strong> 230712628</p>
      <p><em>UTS Visualisasi dan Interpretasi Data</em></p>
    </div>
    <div class="sub">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
    {kpi_html}
    {cards_html}
    <footer>All charts are interactive (Plotly). Desain: clarity → hierarchy → storytelling.</footer>
  </div>
</body>
</html>"""
# ============================================================================

Path(OUT_HTML).write_text(html, encoding="utf-8")
print(f"Selesai ✅  1 halaman HTML dibuat: {Path(OUT_HTML).resolve()}")
