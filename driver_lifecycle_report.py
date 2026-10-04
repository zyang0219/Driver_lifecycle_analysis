"""
W33+ 司机周度得分 -> 司机生命周期分析报告

用法：
    python driver_lifecycle_report.py

- 读取「司机分数趋势底表」，只分析 TARGET_REGIONS 里的大区
- 生成「站点汇总」「待召回优质司机」「新司机」「流失司机」四张表
- 导出到本地 Excel
- 按大区拆分，写入对应的飞书 Wiki 文档

本脚本只定义逻辑、不在生成时自动跑——按你的习惯，改好配置后自己在本地执行。
"""

import os
import re

import numpy as np
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# ===================== 配置 =====================

INPUT_PATH = r"W33司机周度得分.xlsx"
SHEET_NAME = "司机分数趋势底表"

OUTPUT_DIR = r"E:\GOFO Driver Dashboard\History"
OUTPUT_NAME = "W33_Driver_Lifecycle_Report.xlsx"
GOOD_THRESHOLD = 90

TARGET_REGIONS = ["TX", "GL"]

WEEK_TAG = "W34"

FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")

# wiki 链接域名只是自定义文档域名，Open API 通常走官方域名。
# 国际版 Lark 用 open.larksuite.com；国内飞书租户改成 open.feishu.cn
FEISHU_BASE_URL = "https://open.larksuite.com/open-apis"

TX_WIKI_URL = "https://acnjh1thgeif.larkenterprise.com/wiki/Ow7JwKfYOiEKzbkN24hcXesFnfg"
GL_WIKI_URL = "https://acnjh1thgeif.larkenterprise.com/wiki/MXRowEDAhidcLFkUI0cc995bn1e"

PERCENT_COLS = {"总体司机流失率", "优质司机流失率", "新司机占比", "召回率"}

COLS_OUT = None  # 运行时赋值，见 load_and_prepare()


# ===================== 数据处理 =====================

def load_and_prepare(input_path: str, sheet_name: str, target_regions: list) -> pd.DataFrame:
    global COLS_OUT

    df = pd.read_excel(input_path, sheet_name=sheet_name)
    df.columns = [c.strip() for c in df.columns]

    expected_cols = ["大区", "站点", "dsp", "司机id", "w-1司机分数", "w-2司机分数", "w-3司机分数", "w-4司机分数"]
    missing = [c for c in expected_cols if c not in df.columns]
    assert not missing, f"缺少列: {missing}, 实际列名: {list(df.columns)}"

    w1, w2, w3, w4 = "w-1司机分数", "w-2司机分数", "w-3司机分数", "w-4司机分数"
    for c in [w1, w2, w3, w4]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df[df["大区"].isin(target_regions)].reset_index(drop=True)

    df["has_w1"] = df[w1].notna()
    df["has_w2"] = df[w2].notna()
    df["has_w3"] = df[w3].notna()
    df["has_w4"] = df[w4].notna()

    df["is_good_w1"] = df["has_w1"] & (df[w1] >= GOOD_THRESHOLD)
    df["is_good_w2"] = df["has_w2"] & (df[w2] >= GOOD_THRESHOLD)
    df["is_good_w3"] = df["has_w3"] & (df[w3] >= GOOD_THRESHOLD)
    df["is_good_w4"] = df["has_w4"] & (df[w4] >= GOOD_THRESHOLD)

    # 流失：上周在岗，本周不在岗
    df["churned"] = df["has_w2"] & ~df["has_w1"]

    # 优质流失（周环比口径）：上周是优质（w-2>=90），本周没有分数
    df["good_churned"] = df["is_good_w2"] & ~df["has_w1"]

    # 新司机：本周在岗，且过去三周从未出现过
    df["new_driver"] = df["has_w1"] & ~df["has_w2"] & ~df["has_w3"] & ~df["has_w4"]

    # 召回：本周在岗，上周不在岗，但更早（w-3或w-4）出现过
    df["recalled"] = df["has_w1"] & ~df["has_w2"] & (df["has_w3"] | df["has_w4"])

    # 上周"待召回优质司机"池子（召回率分母）：w-3是优质，w-2没有分数
    df["pending_recall_last_week"] = df["is_good_w3"] & ~df["has_w2"]

    COLS_OUT = ["大区", "站点", "dsp", "司机id", w1, w2, w3, w4]
    return df


def safe_div(numer: pd.Series, denom: pd.Series) -> pd.Series:
    return (numer / denom).replace([np.inf, -np.inf], np.nan)


def summarize_by(df_in: pd.DataFrame, group_cols: list) -> pd.DataFrame:
    agg = df_in.groupby(group_cols, as_index=False).agg(
        本周司机人数=("has_w1", "sum"),
        上周司机人数=("has_w2", "sum"),
        总体流失司机数=("churned", "sum"),
        本周优质司机人数=("is_good_w1", "sum"),
        上周优质司机人数=("is_good_w2", "sum"),
        优质司机流失人数=("good_churned", "sum"),
        新司机人数=("new_driver", "sum"),
        召回司机人数=("recalled", "sum"),
        上周待召回优质司机人数=("pending_recall_last_week", "sum"),
    )

    agg["总体司机流失率"] = safe_div(agg["总体流失司机数"], agg["上周司机人数"])
    agg["优质司机流失率"] = safe_div(agg["优质司机流失人数"], agg["上周优质司机人数"])
    agg["新司机占比"] = safe_div(agg["新司机人数"], agg["本周司机人数"])
    agg["召回率"] = safe_div(agg["召回司机人数"], agg["上周待召回优质司机人数"])

    agg = agg.drop(columns=["上周优质司机人数", "上周待召回优质司机人数"])
    return agg


def build_site_report(df: pd.DataFrame) -> pd.DataFrame:
    station_summary = summarize_by(df, ["大区", "站点"])
    region_summary = summarize_by(df, ["大区"])
    region_summary["站点"] = "大区汇总"

    col_order = ["大区", "站点", "本周司机人数", "上周司机人数", "总体流失司机数", "总体司机流失率",
                 "本周优质司机人数", "优质司机流失人数", "优质司机流失率",
                 "新司机人数", "新司机占比", "召回司机人数", "召回率"]
    station_summary = station_summary[col_order]
    region_summary = region_summary[col_order]

    final_rows = []
    for region in station_summary["大区"].unique():
        final_rows.append(station_summary[station_summary["大区"] == region])
        final_rows.append(region_summary[region_summary["大区"] == region])

    return pd.concat(final_rows, ignore_index=True)


def build_pending_recall(df: pd.DataFrame) -> pd.DataFrame:
    """待召回优质司机：过去三周（w-2/w-3/w-4）里只要有一周是优质，且本周没有分数。
    注意：这个口径比「优质司机流失人数」（只看 w-2）更宽，两张表的人数可能对不上，是预期行为。
    如果想跟流失率口径完全一致，改成: df["is_good_w2"] & ~df["has_w1"]
    """
    mask = (df["is_good_w2"] | df["is_good_w3"] | df["is_good_w4"]) & ~df["has_w1"]
    return df[mask][COLS_OUT].reset_index(drop=True)


def build_new_drivers(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["new_driver"]][COLS_OUT].reset_index(drop=True)


def build_churned(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["churned"]][COLS_OUT].reset_index(drop=True)


# ===================== 本地 Excel 导出 =====================

def export_to_excel(site_report, pending_recall, new_drivers, churned, output_dir, output_name):
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_name)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        site_report.to_excel(writer, sheet_name="站点汇总", index=False)
        pending_recall.to_excel(writer, sheet_name="待召回优质司机", index=False)
        new_drivers.to_excel(writer, sheet_name="新司机", index=False)
        churned.to_excel(writer, sheet_name="流失司机", index=False)

    wb = load_workbook(output_path)
    ws = wb["站点汇总"]

    header = [cell.value for cell in ws[1]]
    pct_col_idx = [header.index(c) + 1 for c in PERCENT_COLS if c in header]

    bold_font = Font(bold=True)
    subtotal_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    site_cell_idx = header.index("站点") + 1

    for row in range(2, ws.max_row + 1):
        for col_idx in pct_col_idx:
            cell = ws.cell(row=row, column=col_idx)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.0%"
        if ws.cell(row=row, column=site_cell_idx).value == "大区汇总":
            for col_idx in range(1, ws.max_column + 1):
                c = ws.cell(row=row, column=col_idx)
                c.font = bold_font
                c.fill = subtotal_fill

    for sheet in wb.worksheets:
        for col_cells in sheet.columns:
            length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            col_letter = get_column_letter(col_cells[0].column)
            sheet.column_dimensions[col_letter].width = min(max(length + 2, 10), 30)

    wb.save(output_path)
    print("已保存:", output_path)
    return output_path


# ===================== 飞书写入 =====================

def safe_json(resp: requests.Response, context: str) -> dict:
    """resp.json() 的安全版本：非 JSON 响应时打印状态码和原始内容前 500 字，
    方便定位是不是 base_url/路径不对（比如返回了 HTML 404/登录页）。"""
    try:
        return resp.json()
    except ValueError:
        raise RuntimeError(
            f"{context} 返回的不是 JSON，可能是 base_url 或路径不对。\n"
            f"HTTP 状态码: {resp.status_code}\n"
            f"请求地址: {resp.url}\n"
            f"响应内容前 500 字: {resp.text[:500]!r}"
        )


def get_tenant_access_token(app_id: str, app_secret: str, base_url: str) -> str:
    resp = requests.post(
        f"{base_url}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret},
        timeout=15,
    )
    data = safe_json(resp, "获取 tenant_access_token")
    if data.get("code") != 0:
        raise RuntimeError(f"获取 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


def extract_wiki_node_token(wiki_url: str) -> str:
    m = re.search(r"/wiki/([A-Za-z0-9]+)", wiki_url)
    if not m:
        raise ValueError(f"无法从链接中解析出 node token: {wiki_url}")
    return m.group(1)


def resolve_spreadsheet_token(base_url: str, token: str, node_token: str) -> str:
    resp = requests.get(
        f"{base_url}/wiki/v2/spaces/get_node",
        headers={"Authorization": f"Bearer {token}"},
        params={"token": node_token, "obj_type": "wiki"},
        timeout=15,
    )
    data = safe_json(resp, "解析 wiki 节点")
    if data.get("code") != 0:
        raise RuntimeError(f"解析 wiki 节点失败: {data}")
    return data["data"]["node"]["obj_token"]


def list_sheets(base_url: str, token: str, spreadsheet_token: str) -> list:
    resp = requests.get(
        f"{base_url}/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query",
        headers={"Authorization": f"Bearer {token}"},
        timeout=15,
    )
    data = safe_json(resp, "获取 sheet 列表")
    if data.get("code") != 0:
        raise RuntimeError(f"获取 sheet 列表失败: {data}")
    return data["data"]["sheets"]


def delete_sheet(base_url: str, token: str, spreadsheet_token: str, sheet_id: str):
    resp = requests.post(
        f"{base_url}/sheets/v3/spreadsheets/{spreadsheet_token}/sheets_batch_update",
        headers={"Authorization": f"Bearer {token}"},
        json={"requests": [{"deleteSheet": {"sheetId": sheet_id}}]},
        timeout=15,
    )
    return safe_json(resp, "删除 sheet")


def create_sheet(base_url: str, token: str, spreadsheet_token: str, title: str) -> str:
    resp = requests.post(
        f"{base_url}/sheets/v3/spreadsheets/{spreadsheet_token}/sheets_batch_update",
        headers={"Authorization": f"Bearer {token}"},
        json={"requests": [{"addSheet": {"properties": {"title": title}}}]},
        timeout=15,
    )
    data = safe_json(resp, f"创建 sheet「{title}」")
    if data.get("code") != 0:
        raise RuntimeError(f"创建 sheet「{title}」失败: {data}")
    return data["data"]["replies"][0]["addSheet"]["properties"]["sheetId"]


def ensure_sheet(base_url: str, token: str, spreadsheet_token: str, title: str) -> str:
    """如果同名 sheet 已存在（比如上周跑过一遍），先删掉再重建，保证内容是最新的。"""
    existing = list_sheets(base_url, token, spreadsheet_token)
    for s in existing:
        if s.get("title") == title:
            delete_sheet(base_url, token, spreadsheet_token, s.get("sheet_id") or s.get("sheetId"))
            break
    return create_sheet(base_url, token, spreadsheet_token, title)


def df_to_values(df_in: pd.DataFrame) -> list:
    """DataFrame -> 飞书 values 接口的二维数组：表头 + 数据行，NaN 转空字符串，百分比列转成 "xx.x%" 字符串。"""
    df_fmt = df_in.copy()
    for col in df_fmt.columns:
        if col in PERCENT_COLS:
            df_fmt[col] = df_fmt[col].apply(lambda v: "" if pd.isna(v) else f"{v:.1%}")
        else:
            df_fmt[col] = df_fmt[col].apply(lambda v: "" if pd.isna(v) else v)
    return [list(df_fmt.columns)] + df_fmt.values.tolist()


def write_values(base_url: str, token: str, spreadsheet_token: str, sheet_id: str, df_in: pd.DataFrame):
    values = df_to_values(df_in)
    n_rows = len(values)
    n_cols = len(values[0])
    end_col = chr(ord('A') + n_cols - 1) if n_cols <= 26 else "Z"
    value_range = f"{sheet_id}!A1:{end_col}{n_rows}"

    resp = requests.put(
        f"{base_url}/sheets/v2/spreadsheets/{spreadsheet_token}/values",
        headers={"Authorization": f"Bearer {token}"},
        json={"valueRange": {"range": value_range, "values": values}},
        timeout=30,
    )
    data = safe_json(resp, f"写入数据 ({value_range})")
    if data.get("code") != 0:
        raise RuntimeError(f"写入数据失败 ({value_range}): {data}")
    return data


def bold_header(base_url: str, token: str, spreadsheet_token: str, sheet_id: str, n_cols: int):
    end_col = chr(ord('A') + n_cols - 1) if n_cols <= 26 else "Z"
    try:
        resp = requests.put(
            f"{base_url}/sheet/v2/spreadsheets/{spreadsheet_token}/style",
            headers={"Authorization": f"Bearer {token}"},
            json={"appendStyle": {"range": f"{sheet_id}!A1:{end_col}1", "style": {"font": {"bold": True}}}},
            timeout=15,
        )
        data = safe_json(resp, "表头加粗")
        if data.get("code") != 0:
            print("表头加粗失败（不影响数据），返回:", data)
    except Exception as e:
        print("表头加粗请求异常（不影响数据）:", e)


def push_region_report(region: str, wiki_url: str, token: str,
                        site_report, pending_recall, new_drivers, churned):
    node_token = extract_wiki_node_token(wiki_url)
    spreadsheet_token = resolve_spreadsheet_token(FEISHU_BASE_URL, token, node_token)
    print(f"[{region}] spreadsheet_token = {spreadsheet_token}")

    sheets_to_push = {
        f"{WEEK_TAG}站点汇总": site_report[site_report["大区"] == region].reset_index(drop=True),
        f"{WEEK_TAG}待召回优质司机": pending_recall[pending_recall["大区"] == region].reset_index(drop=True),
        f"{WEEK_TAG}新司机": new_drivers[new_drivers["大区"] == region].reset_index(drop=True),
        f"{WEEK_TAG}流失司机": churned[churned["大区"] == region].reset_index(drop=True),
    }

    for title, data in sheets_to_push.items():
        sheet_id = ensure_sheet(FEISHU_BASE_URL, token, spreadsheet_token, title)
        write_values(FEISHU_BASE_URL, token, spreadsheet_token, sheet_id, data)
        bold_header(FEISHU_BASE_URL, token, spreadsheet_token, sheet_id, len(data.columns))
        print(f"[{region}] 「{title}」写入完成，{len(data)} 行")


# ===================== 主流程 =====================

def main():
    df = load_and_prepare(INPUT_PATH, SHEET_NAME, TARGET_REGIONS)

    site_report = build_site_report(df)
    pending_recall = build_pending_recall(df)
    new_drivers = build_new_drivers(df)
    churned = build_churned(df)

    export_to_excel(site_report, pending_recall, new_drivers, churned, OUTPUT_DIR, OUTPUT_NAME)

    assert FEISHU_APP_ID and FEISHU_APP_SECRET, "请先设置 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量"
    tenant_token = get_tenant_access_token(FEISHU_APP_ID, FEISHU_APP_SECRET, FEISHU_BASE_URL)
    push_region_report("TX", TX_WIKI_URL, tenant_token, site_report, pending_recall, new_drivers, churned)
    push_region_report("GL", GL_WIKI_URL, tenant_token, site_report, pending_recall, new_drivers, churned)
    print("全部写入完成 ✅")


if __name__ == "__main__":
    main()
