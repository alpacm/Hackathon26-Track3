# -*- coding: utf-8 -*-
"""Baseline diagnosis for the Track 3 simulated airport workbook.

The script compares three definitions of an abnormal month, checks whether
adverse weather still correlates after seasonality is removed, and writes:

  figures/fig1_baseline_definitions.png
  results/baseline_three_definitions.csv
  results/baseline_report.txt

Run from this directory: python baseline_diagnosis.py
"""

from __future__ import annotations

import os
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import font_manager

ROOT = Path(__file__).resolve().parent
DATA_PATH = ROOT / "data" / "Track 3 Data.xlsx"
FIG_PATH = ROOT / "figures" / "fig1_baseline_definitions.png"
CSV_PATH = ROOT / "results" / "baseline_three_definitions.csv"
REPORT_PATH = ROOT / "results" / "baseline_report.txt"

BASELINE_SHEET = "01_机场月度运行基线"
WEATHER_SHEET = "03_机场天气月度"
METRIC = "按计划完成航班的比例"
RELEASE = "机场放行正常比例"
PEER_COL = "同类机场平均比例"
OWN_AVG = "该机场过去12个月平均比例"
DEV_COL = "与自身平均差（百分点）"
AIRPORT = "机场名称"
YEAR = "年份"
MONTH = "月份"
CLASS = "机场分类"
ADVERSE_DAYS = "明显不利天气天数"
MISSING = "是否缺少数据"

# Definition A is the contest rule: 5 points below the airport's own history.
TH_A = -5.0
TH_B = -2.0
TH_C = -3.0

C_MAIN, C_SELF, C_PEER, C_WARN = "#2F5D8C", "#E08A1E", "#3E9C7C", "#C0392B"

FONT_CANDIDATES = (
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
)


def configure_fonts() -> None:
    families: list[str] = []
    for path in FONT_CANDIDATES:
        if os.path.isfile(path):
            font_manager.fontManager.addfont(path)
            families.append(font_manager.FontProperties(fname=path).get_name())
    plt.rcParams["font.sans-serif"] = families + ["DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False


def load_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    if not path.is_file():
        raise FileNotFoundError(f"Workbook not found: {path}")
    return pd.read_excel(path, sheet_name=sheet_name)


def prepare_baseline(df: pd.DataFrame) -> pd.DataFrame:
    df = df.sort_values([AIRPORT, YEAR, MONTH]).reset_index(drop=True)
    df["ym"] = df[YEAR].astype(str) + "-" + df[MONTH].astype(str).str.zfill(2)
    return df


def add_peer_excl_self(df: pd.DataFrame) -> pd.DataFrame:
    """Recompute the same-class mean after removing the airport itself.

    The workbook column includes the airport in its own peer average. A sharp
    drop then pulls the peer mean down and makes the gap look smaller.
    """
    grouped = df.groupby([YEAR, MONTH, CLASS])[METRIC]
    count = grouped.transform("size")
    total = grouped.transform("sum")
    excl = np.where(count > 1, (total - df[METRIC]) / (count - 1), np.nan)
    df["同类机场平均_剔除自身"] = excl
    df["同类家数"] = count
    df["含自身同类偏离pp"] = (df[METRIC] - df[PEER_COL]) * 100
    return df


def add_three_definitions(df: pd.DataFrame) -> pd.DataFrame:
    df["A_自身历史偏离pp"] = df[DEV_COL]
    df["hitA"] = df["A_自身历史偏离pp"] < TH_A

    center = df.groupby(AIRPORT)[METRIC].transform("mean")
    spread = df.groupby(AIRPORT)[METRIC].transform("std")
    df["B_自身z分数"] = (df[METRIC] - center) / spread
    df["hitB"] = df["B_自身z分数"] < TH_B

    df["C_同类偏离pp"] = (df[METRIC] - df["同类机场平均_剔除自身"]) * 100
    df["hitC"] = df["C_同类偏离pp"] < TH_C
    return df


def weather_evidence(df: pd.DataFrame, weather: pd.DataFrame) -> dict[str, float]:
    merged = df.merge(weather, on=[YEAR, MONTH, AIRPORT], how="left")
    days = merged[ADVERSE_DAYS]
    gap = merged[DEV_COL]
    month_mean = merged.groupby(MONTH)[DEV_COL].transform("mean")
    airport_mean = merged.groupby(AIRPORT)[DEV_COL].transform("mean")
    return {
        "raw": float(days.corr(gap)),
        "deseason": float(days.corr(gap - month_mean)),
        "deseason_plus_airport": float(days.corr(gap - month_mean - airport_mean)),
    }


def _airport_table(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for airport, group in df.groupby(AIRPORT):
        rows.append(
            {
                "机场": airport,
                "分类": group[CLASS].iloc[0],
                "均值": round(float(group[METRIC].mean()), 4),
                "标准差": round(float(group[METRIC].std()), 4),
                "A命中": int(group["hitA"].sum()),
                "B命中": int(group["hitB"].sum()),
                "C命中": int(group["hitC"].sum()),
            }
        )
    return pd.DataFrame(rows).sort_values("均值").reset_index(drop=True)


def _line(lines: list[str], *parts: object) -> None:
    lines.append(" ".join(str(part) for part in parts))


def build_report(df: pd.DataFrame, weather: pd.DataFrame) -> str:
    lines: list[str] = []
    table = _airport_table(df)
    weather_corr = weather_evidence(df, weather)
    release_corr = float(df[METRIC].corr(df[RELEASE]))
    peer_gap = float((df.groupby([YEAR, MONTH, CLASS])[METRIC].transform("mean") - df[PEER_COL]).abs().max())

    _line(lines, "=" * 78)
    _line(lines, "Track 3 · 基线诊断报告")
    _line(lines, "=" * 78)
    _line(
        lines,
        f"样本：{len(df)} 条 = {df[AIRPORT].nunique()} 个机场 × {df['ym'].nunique()} 个月",
    )
    _line(
        lines,
        f"指标：{METRIC}。它与「{RELEASE}」的相关系数为 {release_corr:.3f}，"
        "二者不能当作两个独立信源。",
    )
    _line(lines, "")

    _line(lines, "-" * 78)
    _line(lines, "【1】三种「异常」定义的命中率")
    _line(lines, "-" * 78)
    _line(lines, "  定义A  比自身过去 12 个月均值至少低 5 个百分点（RULE003）")
    _line(lines, "  定义B  该机场 24 个月 z 分数低于 -2")
    _line(lines, "  定义C  比剔除自身后的同类均值至少低 3 个百分点")
    _line(lines, "")
    for key, threshold in (("hitA", TH_A), ("hitB", TH_B), ("hitC", TH_C)):
        count = int(df[key].sum())
        _line(lines, f"  {key}  阈值={threshold:>5}   命中 {count:>3} 条  占 {100 * count / len(df):5.2f}%")
    only_a = int((df["hitA"] & ~df["hitB"]).sum())
    only_b = int((df["hitB"] & ~df["hitA"]).sum())
    both = int((df["hitA"] & df["hitB"]).sum())
    either = int((df["hitA"] | df["hitB"]).sum())
    _line(lines, "")
    _line(lines, "  定义A与定义B的重叠：")
    _line(lines, f"    A∩B = {both}   A∪B = {either}")
    _line(lines, f"    仅A命中 = {only_a}   仅B命中 = {only_b}")
    if either:
        _line(
            lines,
            f"  在 A 或 B 命中的 {either} 条里，有 {only_a + only_b} 条两种定义不一致"
            f"（{100 * (only_a + only_b) / either:.0f}%）。换一把尺子，结论就会变。",
        )
    _line(lines, "")

    _line(lines, "-" * 78)
    _line(lines, "【2】各机场画像")
    _line(lines, "-" * 78)
    _line(lines, table.to_string(index=False))
    _line(lines, "")
    widest = table.loc[table["标准差"].idxmax()]
    steadiest = table.loc[table["标准差"].idxmin()]
    ratio = widest["标准差"] / steadiest["标准差"]
    _line(
        lines,
        f"  {widest['机场']}的标准差 {widest['标准差']:.4f} 是"
        f"{steadiest['机场']} {steadiest['标准差']:.4f} 的 {ratio:.1f} 倍。"
        f"同样的 {abs(TH_A):.0f} 个百分点，对高波动机场更常见，对低波动机场更罕见。",
    )
    low_peer = table.sort_values("C命中", ascending=False).iloc[0]
    classmates = table[table["分类"] == low_peer["分类"]]
    others = classmates[classmates["机场"] != low_peer["机场"]]
    other_means = "、".join(f"{row['机场']} {row['均值']:.3f}" for _, row in others.iterrows())
    _line(
        lines,
        f"  {low_peer['机场']}均值 {low_peer['均值']:.3f}，同类为{other_means}。"
        f"定义C命中 {int(low_peer['C命中'])} 个月。只和同类比，会把长期偏低的机场常年标成异常。",
    )
    split = table.assign(差=(table["A命中"] - table["B命中"]).abs()).sort_values("差", ascending=False).iloc[0]
    _line(
        lines,
        f"  {split['机场']}的定义A命中 {int(split['A命中'])} 次、定义B命中 {int(split['B命中'])} 次。"
        "固定的 5 个百分点，和按该机场自己的标准差计算的 z 分数，不是同一把尺子。",
    )
    _line(lines, "")

    _line(lines, "-" * 78)
    _line(lines, "【3】含自身的同类均值会稀释异常")
    _line(lines, "-" * 78)
    _line(lines, f"  表内「{PEER_COL}」与含自身重算均值的最大绝对差为 {peer_gap:.6f}。")
    _line(lines, "  这列就是含自身的分类均值。本脚本另算了剔除自身后的同类均值。")
    dilution = (df["C_同类偏离pp"] - df["含自身同类偏离pp"]).abs()
    example = df.loc[dilution.idxmax()]
    _line(
        lines,
        f"  稀释最大的一条是 {example['ym']} {example[AIRPORT]}："
        f"完成率 {example[METRIC]:.4f}。",
    )
    _line(lines, f"    相对含自身同类均值：{example['含自身同类偏离pp']:+.2f} pp")
    _line(lines, f"    剔除自身后：        {example['C_同类偏离pp']:+.2f} pp")
    _line(lines, f"    相对自身过去 12 个月：{example[DEV_COL]:+.2f} pp")
    _line(lines, "")

    _line(lines, "-" * 78)
    _line(lines, "【4】天气是关联线索，不是因果证据")
    _line(lines, "-" * 78)
    _line(lines, f"  corr(明显不利天气天数, 与自身平均差)    原始 = {weather_corr['raw']:+.3f}")
    _line(lines, f"  扣除月份效应后                          = {weather_corr['deseason']:+.3f}")
    _line(lines, f"  再扣除机场自身均值                      = {weather_corr['deseason_plus_airport']:+.3f}")
    _line(lines, "")
    _line(lines, "  原始相关几乎全部来自季节性公共成分。去掉季节后接近 0。")
    _line(lines, "  月度数据只能说天气与运行波动同时出现，不能说天气导致了延误。")
    _line(lines, "  可用写法：天气是风险上升的关联线索；逐航班原因缺失，无法确认因果。")
    _line(lines, "")

    _line(lines, "-" * 78)
    _line(lines, "【5】数据核对")
    _line(lines, "-" * 78)
    cold = df[(df[YEAR] == 2024) & (df[MONTH] == 1)]
    cold_zero = int((cold[DEV_COL] == 0).sum())
    _line(lines, f"  2024-01 共 {len(cold)} 条，其中 {cold_zero} 条「与自身平均差」为 0。这是基线冷启动月。")
    merged = df.merge(weather, on=[YEAR, MONTH, AIRPORT], how="left")
    missing = merged[merged[MISSING] == "是"][["ym", AIRPORT, METRIC]]
    _line(lines, f"  标记为缺少天气数据的记录 {len(missing)} 条，这些月份不能把天气插值后当成事实：")
    for _, row in missing.iterrows():
        _line(lines, f"    · {row['ym']}  {row[AIRPORT]}  运行值 {row[METRIC]:.4f}")
    hail = df[(df[AIRPORT] == "海陵国际机场") & (df["ym"].isin(["2025-07", "2025-08"]))]
    hail_text = "、".join(f"{row['ym']} {row[DEV_COL]:+.2f} pp" for _, row in hail.iterrows())
    _line(lines, f"  海陵国际机场设计异常窗口：{hail_text}。")
    _line(lines, "")

    _line(lines, "-" * 78)
    _line(lines, "【6】结论")
    _line(lines, "-" * 78)
    _line(lines, "  一个绝对阈值不够。判断某个月是否异常，需要同时看三件事：")
    _line(lines, f"  1. 是否低于该机场自身历史至少 {abs(TH_A):.0f} 个百分点（定义A，RULE003）。")
    _line(lines, "  2. 这次偏离相对它自己的波动是否罕见（定义B）。")
    _line(lines, "  3. 是否也低于剔除自身之后的同类机场（定义C）。含自身的同类均值会把异常稀释掉。")
    _line(lines, "  天气只作为关联线索。完成率与放行正常率高度相关，不能算成两个独立证据。")
    return "\n".join(lines) + "\n"


def make_figure(df: pd.DataFrame, fig_path: Path) -> None:
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14.5, 5.6), gridspec_kw={"width_ratios": [1.35, 1]})

    airport = "海陵国际机场"
    series = df[df[AIRPORT] == airport].copy()
    x = np.arange(len(series))
    window = [i for i, ym in enumerate(series["ym"]) if ym in ("2025-07", "2025-08")]
    if window:
        ax1.axvspan(min(window) - 0.5, max(window) + 0.5, color=C_WARN, alpha=0.10, zorder=0)

    self_base = series[OWN_AVG].to_numpy()
    ax1.fill_between(
        x,
        self_base - 0.05,
        self_base + 0.05,
        color=C_SELF,
        alpha=0.13,
        lw=0,
        zorder=1,
        label="RULE003 容忍带（基线①±5pp）",
    )
    ax1.plot(x, series[METRIC], "-o", color=C_MAIN, lw=2.0, ms=4, label="实际值", zorder=4)
    ax1.plot(x, self_base, "--", color=C_SELF, lw=1.6, label="基线①：自身过去 12 个月均值", zorder=3)
    ax1.plot(
        x,
        series["同类机场平均_剔除自身"],
        ":",
        color=C_PEER,
        lw=1.8,
        label="基线②：同类机场均值（剔除自身）",
        zorder=3,
    )

    low_at = int(np.argmin(series["A_自身历史偏离pp"].to_numpy()))
    gap = float(series["A_自身历史偏离pp"].iloc[low_at])
    times = abs(gap / TH_A)
    ax1.annotate(
        f"{series['ym'].iloc[low_at]} 偏离 {gap:.2f} pp\n(容忍带的 {times:.1f} 倍)",
        xy=(low_at, series[METRIC].iloc[low_at]),
        xytext=(max(0, low_at - 8), float(series[METRIC].iloc[low_at]) + 0.03),
        fontsize=9,
        color=C_WARN,
        va="center",
        arrowprops={"arrowstyle": "->", "color": C_WARN, "lw": 1.2, "connectionstyle": "arc3,rad=-0.18"},
    )
    ax1.set_xticks(x[::4])
    ax1.set_xticklabels(series["ym"].iloc[::4], rotation=45, ha="right", fontsize=8)
    ax1.set_ylabel("按计划完成航班的比例")
    plotted = np.concatenate(
        [
            series[METRIC].to_numpy(),
            self_base,
            series["同类机场平均_剔除自身"].to_numpy(dtype=float),
        ]
    )
    low = float(np.nanmin(plotted))
    high = float(np.nanmax(plotted))
    pad = (high - low) * 0.22
    ax1.set_ylim(low - pad, high + pad * 0.45)
    ax1.yaxis.set_major_formatter(lambda value, _pos: f"{value * 100:.0f}%")
    ax1.set_title(f"(a) {airport}：设计异常窗口与两套基线", fontsize=11, loc="left")
    ax1.legend(fontsize=8.5, loc="lower left", framealpha=0.92)
    ax1.grid(alpha=0.25, ls=":")

    pivot = df.groupby(AIRPORT)[["hitA", "hitB", "hitC"]].sum()
    order = df.groupby(AIRPORT)[METRIC].mean().sort_values().index
    pivot = pivot.loc[order]
    positions = np.arange(len(pivot))
    width = 0.26
    groups = [
        ax2.bar(
            positions - width,
            pivot["hitA"],
            width,
            label=f"定义A 自身历史偏离 < {TH_A:.0f}pp",
            color=C_MAIN,
            edgecolor="white",
            lw=0.6,
        ),
        ax2.bar(
            positions,
            pivot["hitB"],
            width,
            label=f"定义B 自身 z < {TH_B:.0f}",
            color=C_SELF,
            edgecolor="white",
            lw=0.6,
        ),
        ax2.bar(
            positions + width,
            pivot["hitC"],
            width,
            label=f"定义C 同类偏离 < {TH_C:.0f}pp",
            color=C_PEER,
            edgecolor="white",
            lw=0.6,
        ),
    ]
    for group in groups:
        for bar in group:
            height = bar.get_height()
            if height >= 1:
                ax2.text(
                    bar.get_x() + bar.get_width() / 2,
                    height + 0.12,
                    f"{int(height)}",
                    ha="center",
                    va="bottom",
                    fontsize=7.5,
                    color="#333333",
                )
    ax2.set_xticks(positions)
    ax2.set_xticklabels([name.replace("国际机场", "\n国际机场") for name in pivot.index], fontsize=8)
    ax2.set_ylabel("被判为「异常」的月份数（共 24 个月）")
    ax2.set_title("(b) 三种定义互相不一致", fontsize=11, loc="left")
    ax2.legend(fontsize=8.5, loc="upper left", framealpha=0.92)
    ax2.grid(alpha=0.25, axis="y", ls=":")
    ax2.set_ylim(0, float(pivot.to_numpy().max()) * 1.34 + 0.6)
    ax2.text(
        0.99,
        0.98,
        f"A∩B={int((df.hitA & df.hitB).sum())}   A∪B={int((df.hitA | df.hitB).sum())}\n"
        f"仅A命中={int((df.hitA & ~df.hitB).sum())}   仅B命中={int((df.hitB & ~df.hitA).sum())}",
        transform=ax2.transAxes,
        ha="right",
        va="top",
        fontsize=8.5,
        bbox={"boxstyle": "round,pad=0.4", "fc": "white", "ec": "#CCCCCC", "alpha": 0.95},
    )

    fig.suptitle(
        "图1  基线诊断：同一个「异常」在三种参考系下不是同一件事",
        fontsize=12.5,
        y=0.99,
        x=0.02,
        ha="left",
    )
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    fig_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(fig_path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    configure_fonts()
    baseline = prepare_baseline(load_sheet(DATA_PATH, BASELINE_SHEET))
    weather = load_sheet(DATA_PATH, WEATHER_SHEET)
    diagnosed = add_three_definitions(add_peer_excl_self(baseline))

    report = build_report(diagnosed, weather)
    print(report, end="")

    CSV_PATH.parent.mkdir(parents=True, exist_ok=True)
    diagnosed.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    REPORT_PATH.write_text(report, encoding="utf-8")
    make_figure(diagnosed, FIG_PATH)

    print(f"[OK] 图   -> {FIG_PATH}")
    print(f"[OK] 表   -> {CSV_PATH}")
    print(f"[OK] 报告 -> {REPORT_PATH}")


if __name__ == "__main__":
    main()
