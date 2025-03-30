import os
import json
import pandas as pd
from 读取配置文件模块 import load_config

def load_config(config_path="配置.json"):
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"未找到配置文件：{config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # 获取期号与路径
    issue = cfg["issue"]
    raw_excel_path = cfg["excel_path"]
    output_html = cfg.get("output_html", f"output_{issue}.html")

    # 动态替换路径模板中的 {issue}
    excel_path = raw_excel_path.replace("{issue}", issue)
    output_html = output_html.replace("{issue}", issue)
    # 解析为绝对路径
    excel_path = os.path.abspath(os.path.join(os.getcwd(), excel_path))
    output_html = os.path.abspath(os.path.join(os.getcwd(), output_html))
    
    return excel_path, output_html, issue

def compute_analysis_fields(row):
    try:
        sp_win = float(row.get("初盘主胜赔率", "0") or "0")
        sp_draw = float(row.get("初盘平局赔率", "0") or "0")
        sp_lose = float(row.get("初盘客胜赔率", "0") or "0")
        k_win = float(row.get("初盘主凯利", "0") or "0")
        k_draw = float(row.get("初盘平凯利", "0") or "0")
        k_lose = float(row.get("初盘客凯利", "0") or "0")
        handicap = float(row.get("初盘盘口", "0") or "0")

        # 冷热评分（赔率总和 * 5）
        cold_score = round((sp_win + sp_draw + sp_lose) * 5, 2)

        # 凯利异常（差值超过0.1）
        k_list = [k_win, k_draw, k_lose]
        k_max = max(k_list)
        k_min = min(k_list)
        kelly_warning = "⚠️ 是" if (k_max - k_min) >= 0.1 else "正常"

        # 冷门信号判断
        cold_flag = False
        if (sp_win > 3 and k_win > 0.95) or (sp_lose > 3 and k_lose > 0.95):
            cold_flag = True
        cold_signal = "🔴 有" if cold_flag else "无"

        # 庄家策略判断
        if abs(handicap) >= 1.5:
            strategy = "深盘造热"
        elif abs(handicap) <= 0.25:
            strategy = "低盘防冷"
        else:
            strategy = "中庸博弈"

        # 投注倾向建议
        if cold_flag:
            tip = "防平局" if sp_draw < 3.5 else "防冷门"
        else:
            tip = "支持主胜" if sp_win < sp_lose else "倾向客胜"

        return cold_score, kelly_warning, cold_signal, strategy, tip

    except:
        return "-", "-", "-", "-", "-"

def render_dashboard_with_analysis(excel_path, output_path="智能雷达仪表盘.html"):
    df = pd.read_excel(excel_path).fillna("").astype(str)

    base_cols = ["场次", "联赛", "主队", "客队"]
    initial_cols = [col for col in df.columns if col.startswith("初盘")]
    middle_cols = [col for col in df.columns if col.startswith("中盘")]
    final_cols = [col for col in df.columns if col.startswith("临盘")]
    end_cols = [col for col in df.columns if col.startswith("封盘")]

    # 添加智能分析字段（运行时计算，不修改原Excel）
    analysis_fields = ["冷热评分", "凯利异常", "冷门信号", "庄家策略", "投注倾向"]
    for field in analysis_fields:
        df[field] = ""

    for i, row in df.iterrows():
        results = compute_analysis_fields(row)
        for j, field in enumerate(analysis_fields):
            df.at[i, field] = results[j]

    # 构建主表和详情
    html_rows = ""
    for _, row in df.iterrows():
        row_html = "<tr class='main-row'>"
        for col in base_cols + analysis_fields:
            row_html += f"<td>{row.get(col, '')}</td>"
        row_html += "<td><button class='expand-btn'>＋</button></td></tr>"

        def block(title, cols):
            if not cols: return ""
            header = "".join([f"<th>{c}</th>" for c in cols])
            values = "".join([f"<td>{row.get(c, '')}</td>" for c in cols])
            return f"<div><b>{title}</b><table class='inner'><tr>{header}</tr><tr>{values}</tr></table></div>"

        detail_html = (
            "<div style='padding:10px'>"
            + block("📊 初盘数据", initial_cols)
            + block("⏱️ 中盘数据", middle_cols)
            + block("⏳ 临盘数据", final_cols)
            + block("🔚 封盘数据", end_cols)
            + "</div>"
        )
        detail_row = f"<tr class='detail-row' style='display:none'><td colspan='{len(base_cols + analysis_fields) + 1}'>{detail_html}</td></tr>"
        html_rows += row_html + detail_row

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>足彩智能雷达仪表盘</title>
        <style>
            body {{ font-family: "Microsoft YaHei"; padding: 20px; }}
            table {{ border-collapse: collapse; width: 100%; }}
            th, td {{ border: 1px solid #ccc; padding: 5px; text-align: center; }}
            .expand-btn {{
                background: #3498db; color: white; border: none; padding: 5px 10px;
                border-radius: 3px; cursor: pointer;
            }}
            .detail-row td {{ background: #f9f9f9; }}
            .inner {{ margin: 10px 0; width: 100%; border: 1px solid #ddd; }}
        </style>
    </head>
    <body>
        <h2>足彩盘口智能雷达仪表盘（分析主表 + 展开三盘详情）</h2>
        <table>
            <thead><tr>{"".join([f"<th>{c}</th>" for c in base_cols + analysis_fields])}<th>更多</th></tr></thead>
            <tbody>{html_rows}</tbody>
        </table>
        <script>
            document.querySelectorAll('.expand-btn').forEach(btn => {{
                btn.addEventListener('click', () => {{
                    const detailRow = btn.parentElement.parentElement.nextElementSibling;
                    const expanded = detailRow.style.display === 'table-row';
                    detailRow.style.display = expanded ? 'none' : 'table-row';
                    btn.textContent = expanded ? '＋' : '－';
                }});
            }});
        </script>
    </body>
    </html>
    """

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ 页面生成成功：{output_path}")

# 示例执行（你可以传入任意足彩Excel）
if __name__ == "__main__":
    excel_path, output_html, issue = load_config()
    render_dashboard_with_analysis(excel_path, output_html)
