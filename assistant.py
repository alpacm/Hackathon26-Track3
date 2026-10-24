# -*- coding: utf-8 -*-
"""航班异常应对。读模拟工作簿，按规则给建议。运行时不调用大模型。"""

from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
BOOK = ROOT / "data" / "Track 3 Data.xlsx"
OUT = ROOT / "results"
DOCS = ROOT / "docs"

# 和 baseline_diagnosis.py 里的定义 A 一样：严格小于 -5，不是小于等于。
RULE003_LINE = -5.0

WEATHER_WORDS = (
    "强降雨",
    "雷雨",
    "阵风",
    "大风",
    "低温",
    "降雪",
    "暴雨",
    "降雨",
    "大雾",
    "台风",
    "风速",
    "天气风险",
)
PRESSURE_WORDS = (
    "下降",
    "登机口",
    "运行压力",
    "航班量较集中",
    "后续任务仍集中",
    "后续航班集中",
    "航班集中",
    "保障承压",
    "拥堵",
)
CONFLICT_WORDS = (
    "冲突",
    "不一致",
    "来源甲",
    "来源乙",
    "更新时间相差",
    "口径不一致",
    "两个来源",
    "来源可靠性",
)
DENIAL_WORDS = ("拒绝", "否决", "人工否决", "现场资源充足", "资源充足", "足以应对", "不升级")
SAFETY_WORDS = (
    "取消航班",
    "自动取消",
    "取消晚间",
    "请自动取消",
    "放行",
    "调度",
    "维修",
    "赔付",
    "直接发送",
    "正式发布",
    "向旅客发送",
    "发送通知",
)
PII_WORDS = ("手机号", "手机号码", "身份证", "旅客姓名", "员工姓名")
PASSENGER_WORDS = ("旅客", "乘客", "旅客服务", "旅客沟通")

ALLOWED_RISK = ("低", "中", "高", "无法判断")
EVENT_KEYS = (
    "事件编号",
    "事件名称",
    "事件时间",
    "机场名称",
    "天气描述",
    "运行情况描述",
    "航空公司业务重要程度",
    "现场备注",
    "已知信息",
    "未知信息",
)

_CACHE: dict = {}


def tables() -> dict:
    if _CACHE:
        return _CACHE
    if not BOOK.is_file():
        raise FileNotFoundError(f"找不到工作簿：{BOOK}")
    _CACHE["base"] = pd.read_excel(BOOK, sheet_name="01_机场月度运行基线")
    _CACHE["scale"] = pd.read_excel(BOOK, sheet_name="02_机场年度规模")
    _CACHE["weather"] = pd.read_excel(BOOK, sheet_name="03_机场天气月度")
    _CACHE["events"] = pd.read_excel(BOOK, sheet_name="04_模拟事件").fillna("")
    _CACHE["rules"] = pd.read_excel(BOOK, sheet_name="05_运营规则").fillna("")
    _CACHE["tests"] = pd.read_excel(BOOK, sheet_name="08_测试问题").fillna("")
    _CACHE["sources"] = pd.read_excel(BOOK, sheet_name="10_数据来源登记").fillna("")
    return _CACHE


def has_word(text: str, words: tuple[str, ...]) -> bool:
    return any(word in text for word in words)


def airport_names() -> list[str]:
    names = tables()["base"]["机场名称"].astype(str).unique().tolist()
    return sorted(names, key=len, reverse=True)


def blank_event() -> dict:
    return {key: "" for key in EVENT_KEYS}


def as_event(raw: dict) -> dict:
    ev = blank_event()
    for key in EVENT_KEYS:
        value = raw.get(key, "")
        if value is None or (isinstance(value, float) and pd.isna(value)):
            value = ""
        ev[key] = str(value).strip()
    return ev


def event_rows() -> list[dict]:
    return tables()["events"].to_dict(orient="records")


def find_airport(*chunks: str) -> str:
    text = " ".join(chunks)
    for name in airport_names():
        if name and name in text:
            return name
    return ""


def ym_of(year, month) -> str:
    return f"{int(year):04d}-{int(month):02d}"


def global_span() -> tuple[str, str]:
    base = tables()["base"]
    stamp = list(zip(base["年份"], base["月份"]))
    start = ym_of(*min(stamp))
    end = ym_of(*max(stamp))
    return start, end


def baseline_of(airport: str) -> dict:
    empty = {"found": False, "name": airport}
    if not airport:
        return empty
    base = tables()["base"]
    group = base[base["机场名称"].astype(str) == airport].copy()
    if group.empty:
        return empty
    group["年份"] = pd.to_numeric(group["年份"], errors="coerce")
    group["月份"] = pd.to_numeric(group["月份"], errors="coerce")
    group = group.dropna(subset=["年份", "月份"]).sort_values(["年份", "月份"])
    metric = pd.to_numeric(group["按计划完成航班的比例"], errors="coerce")
    dev = pd.to_numeric(group["与自身平均差（百分点）"], errors="coerce")
    group = group.assign(_metric=metric, _dev=dev)
    hits = []
    for _, row in group[group["_dev"] < RULE003_LINE].iterrows():
        hits.append(
            {
                "ym": ym_of(row["年份"], row["月份"]),
                "value": float(row["_metric"]),
                "dev": float(row["_dev"]),
            }
        )
    hits.sort(key=lambda item: item["dev"])
    last = group.iloc[-1]
    first = group.iloc[0]
    return {
        "found": True,
        "name": airport,
        "n": int(metric.notna().sum()),
        "mean": float(metric.mean()),
        "std": float(metric.std()) if metric.notna().sum() > 1 else None,
        "start": ym_of(first["年份"], first["月份"]),
        "end": ym_of(last["年份"], last["月份"]),
        "latest_ym": ym_of(last["年份"], last["月份"]),
        "latest": float(last["_metric"]),
        "klass": str(last["机场分类"]),
        "source": str(last["来源编号"]),
        "hits": hits,
    }


def scale_of(airport: str) -> dict:
    empty = {"found": False}
    if not airport:
        return empty
    scale = tables()["scale"]
    group = scale[scale["机场名称"].astype(str) == airport].copy()
    if group.empty:
        return empty
    group["年份"] = pd.to_numeric(group["年份"], errors="coerce")
    last = group.sort_values("年份").iloc[-1]
    pax = pd.to_numeric(last["旅客吞吐量（万人次）"], errors="coerce")
    return {
        "found": True,
        "year": int(last["年份"]),
        "source": str(last["来源编号"]),
        "importance": str(last["航空公司业务重要程度"]).strip(),
        "pax": float(pax) if pd.notna(pax) else None,
        "klass": str(last["机场分类"]),
        "city": str(last["城市"]),
    }


def weather_of(airport: str) -> dict:
    empty = {"found": False, "missing": [], "source": "SIM-WEATHER"}
    if not airport:
        return empty
    weather = tables()["weather"]
    group = weather[weather["机场名称"].astype(str) == airport].copy()
    if group.empty:
        return empty
    group["年份"] = pd.to_numeric(group["年份"], errors="coerce")
    group["月份"] = pd.to_numeric(group["月份"], errors="coerce")
    group = group.dropna(subset=["年份", "月份"]).sort_values(["年份", "月份"])
    missing = [
        ym_of(row["年份"], row["月份"])
        for _, row in group[group["是否缺少数据"].astype(str) == "是"].iterrows()
    ]
    usable = group[group["是否缺少数据"].astype(str) != "是"]
    latest = usable.iloc[-1] if not usable.empty else None
    days = None
    rain = None
    latest_ym = ""
    if latest is not None:
        latest_ym = ym_of(latest["年份"], latest["月份"])
        days = pd.to_numeric(latest["明显不利天气天数"], errors="coerce")
        rain = pd.to_numeric(latest["降水量（毫米）"], errors="coerce")
        days = None if pd.isna(days) else int(days)
        rain = None if pd.isna(rain) else float(rain)
    return {
        "found": True,
        "source": str(group["来源编号"].iloc[-1]),
        "missing": missing,
        "latest_ym": latest_ym,
        "latest_days": days,
        "latest_rain": rain,
    }


def association() -> tuple[float, float]:
    cached = _CACHE.get("assoc")
    if cached:
        return cached
    base = tables()["base"]
    weather = tables()["weather"][["年份", "月份", "机场名称", "明显不利天气天数"]]
    merged = base.merge(weather, on=["年份", "月份", "机场名称"], how="left")
    days = pd.to_numeric(merged["明显不利天气天数"], errors="coerce")
    gap = pd.to_numeric(merged["与自身平均差（百分点）"], errors="coerce")
    month_mean = gap.groupby(merged["月份"]).transform("mean")
    pair = (float(days.corr(gap)), float(days.corr(gap - month_mean)))
    _CACHE["assoc"] = pair
    return pair


def source_note(code: str) -> str:
    src = tables()["sources"]
    rows = src[src["来源编号"].astype(str) == code]
    if rows.empty:
        return ""
    return str(rows.iloc[0]["备注"]).strip()


def phone_and_id(text: str) -> list[str]:
    found = re.findall(r"(?<!\d)1[3-9]\d{9}(?!\d)", text)
    found += re.findall(r"(?<!\d)\d{17}[\dXx](?!\d)", text)
    return found


def name_tokens(text: str) -> list[str]:
    names = []
    for match in re.finditer(r"(旅客姓名|员工姓名)\s*[:：]?\s*([\u4e00-\u9fff]{2,4})", text):
        token = match.group(2)
        if any(word in token for word in ("手机", "身份", "员工", "旅客", "输入")):
            continue
        names.append(token)
    return names


def collect_secrets(text: str) -> list[str]:
    secrets = phone_and_id(text) + name_tokens(text)
    uniq = []
    for item in sorted(set(secrets), key=len, reverse=True):
        if len(item) >= 2 and item not in uniq:
            uniq.append(item)
    return uniq


def redact_text(text: str, secrets: list[str]) -> str:
    for secret in secrets:
        text = text.replace(secret, "（已省略）")
    return text


def split_gaps(text: str) -> list[str]:
    parts = re.split(r"[；;，,、。\n]", text)
    gaps = []
    for part in parts:
        part = part.strip()
        if part and part not in gaps:
            gaps.append(part)
    return gaps


def borrow_event(ev: dict, question: str) -> dict:
    """问「为什么这个机场风险上升」时，用该机场已有的模拟事件，不另编天气。"""
    if "为什么" not in question:
        return ev
    airport = find_airport(ev["机场名称"], question)
    if not airport:
        return ev
    matched = [row for row in event_rows() if str(row.get("机场名称", "")).strip() == airport]
    if not matched:
        return ev
    src = matched[0]
    borrowed = dict(ev)
    for key in EVENT_KEYS:
        current = borrowed[key].strip()
        incoming = str(src.get(key, "") or "").strip()
        if key == "事件编号" or not current or current == question or current.endswith("？"):
            if incoming:
                borrowed[key] = incoming
    borrowed["机场名称"] = airport
    borrowed["_from"] = str(src.get("事件编号", "")).strip()
    return borrowed


def history_text(base: dict, scale: dict, weather: dict) -> str:
    if not base.get("found"):
        return "模拟基线里没有这个机场，完成率不编。"
    bits = [
        (
            f"月度基线是 {base['start']} 到 {base['end']}，共 {base['n']} 个月，"
            f"完成率均值 {base['mean']:.4f}，最近一个月 {base['latest_ym']} 为 {base['latest']:.4f}"
            f"（{base['source']}）。"
        ),
        "这段是月度历史，对不上事件发生的那个小时，两个时段分开看。",
    ]
    if base["hits"]:
        shown = "、".join(
            f"{hit['ym']} 完成率 {hit['value']:.4f}，低 {abs(hit['dev']):.2f} 个点"
            for hit in base["hits"][:4]
        )
        bits.append(f"比过去 12 个月均值低过 5 个点的月份：{shown}。这是 RULE003，只说明这些月偏离大。")
    if scale.get("found"):
        pax = f"{scale['pax']:.0f} 万人次" if scale.get("pax") is not None else "吞吐量缺失"
        bits.append(
            f"{scale['year']} 年业务重要程度为{scale['importance']}，旅客吞吐量 {pax}（{scale['source']}）。"
        )
    if weather.get("found") and weather.get("missing"):
        bits.append(
            "天气表标了缺失的月份："
            + "、".join(weather["missing"])
            + f"，这些月不把天气补上（{weather['source']}）。"
        )
    elif weather.get("found") and weather.get("latest_ym"):
        bits.append(
            f"天气表里最近一个未标缺失的月份是 {weather['latest_ym']}，"
            f"明显不利天气 {weather['latest_days']} 天。"
            "这不是事件当天的预报，只能当背景（"
            f"{weather['source']}）。"
        )
    return "".join(bits)


def causal_text() -> str:
    raw, deseason = association()
    return (
        f"八个机场放在一起，不利天气天数和「与自身平均差」的相关是 {raw:+.3f}，"
        f"按月份去掉季节后是 {deseason:+.3f}。"
        "剩下几乎没有相关，所以天气只是关联线索，不能拿来证明天气导致了运行异常，更不能说导致了全部航班。"
    )


def add_rule(rules: list[str], rule_id: str) -> None:
    if rule_id not in rules:
        rules.append(rule_id)


def choose_mode(question: str, flags: dict, borrowed: bool, prompt: bool) -> str:
    # 借到模拟事件之后按事件判。测试题经常不带问号，所以靠「是不是一整句提问」而不是看标点。
    if borrowed:
        if flags["conflict"]:
            return "conflict"
        if flags["denial"]:
            return "veto_case"
        return "case"
    if prompt and flags["pii"]:
        return "privacy"
    if prompt and flags["safety"]:
        return "safety"
    if prompt and ("现实" in question or "真实" in question):
        return "fictional"
    if prompt and "天气" in question and "导致" in question:
        return "causal"
    if flags["conflict"]:
        return "conflict"
    if prompt and "时段" in question:
        return "time"
    if prompt and "前序飞机位置" in question and has_word(question, ("没有", "未知", "怎么办", "缺失")):
        return "gap"
    if prompt and "延误" in question and "一定" in question:
        return "insufficient"
    if prompt and flags["denial"]:
        return "veto_q"
    if flags["denial"]:
        return "veto_case"
    return "case"


def analyze(raw: dict) -> dict:
    ev = as_event(raw)
    question = " ".join(ev[key] for key in ("事件名称", "运行情况描述", "现场备注", "已知信息", "未知信息"))
    # 测试题通常只塞一句话，没有天气、备注这些栏。整段事件则这些栏是有内容的。
    prompt = not any(ev[key].strip() for key in ("天气描述", "现场备注", "已知信息", "未知信息"))
    ev = borrow_event(ev, question)
    borrowed = ev.pop("_from", "")
    airport = ev["机场名称"] or find_airport(question, ev["天气描述"], ev["运行情况描述"])
    if airport:
        ev["机场名称"] = airport

    raw_blob = " ".join(ev.values()) + " " + question
    secrets = collect_secrets(raw_blob)
    pii = bool(secrets) or has_word(raw_blob, PII_WORDS)
    for key in EVENT_KEYS:
        ev[key] = redact_text(ev[key], secrets)

    fact_text = " ".join(ev[key] for key in ("事件名称", "天气描述", "运行情况描述", "现场备注", "已知信息"))
    # 未知信息里的「登机口资源」只是缺口，不当成已经发生的运行压力。
    signal_text = fact_text + " " + question
    weather = has_word(fact_text, WEATHER_WORDS)
    pressure = has_word(fact_text, PRESSURE_WORDS)
    conflict = has_word(signal_text, CONFLICT_WORDS)
    denial = has_word(signal_text, DENIAL_WORDS)
    safety = has_word(fact_text, SAFETY_WORDS)

    base = baseline_of(airport)
    scale = scale_of(airport)
    weather_row = weather_of(airport)
    importance = ev["航空公司业务重要程度"] or (scale.get("importance") if scale.get("found") else "")
    important = importance == "高"

    flags = {
        "weather": weather,
        "pressure": pressure and not conflict,
        "conflict": conflict,
        "denial": denial,
        "safety": safety,
        "pii": pii,
        "important": important,
    }
    mode = choose_mode(question, flags, bool(borrowed), prompt)
    rules: list[str] = []
    gaps = split_gaps(ev["未知信息"])
    if mode == "gap" and "前序飞机位置" not in "".join(gaps):
        gaps.append("前序飞机位置")
    if not gaps and mode in ("case", "veto_case", "insufficient"):
        gaps.append("正式航班状态以及受影响的具体航班")

    risk = "低"
    reason: list[str] = []
    impact = ""
    hypothesis = causal_text()
    first = ""
    actions: list[str] = []
    escalate = "要升级的话，至少两个对得上的来源都确认运行异常或保障能力下降，并且由人点头。现在这条做不到。"
    checks: list[str] = []
    used_base = False
    used_scale = False
    used_weather = False

    def remember_history() -> None:
        nonlocal used_base, used_scale, used_weather
        if base.get("found"):
            used_base = True
            if base["hits"]:
                add_rule(rules, "RULE003")
        if scale.get("found"):
            used_scale = True
        if weather_row.get("found") and (weather or weather_row.get("missing") or weather_row.get("latest_ym")):
            used_weather = True

    if mode == "privacy":
        risk = "无法判断"
        add_rule(rules, "RULE011")
        add_rule(rules, "RULE001")
        reason.append("输入里有旅客或员工个人信息。姓名、手机号、身份证不能进建议，也不能留在输出里。")
        reason.append("请把这些内容删除，或者换成不带身份的汇总，再重新做一次判断。")
        impact = "在个人信息去掉之前，不判断航班影响。"
        hypothesis = "没有可引用的汇总事实，做不了运行判断。"
        first = "先删除或匿名，再重跑。"
        actions = [
            "不要把姓名、号码写进日志或旅客稿。",
            "换成航班量、时段、机场这种汇总之后再评估。",
        ]
        gaps = ["匿名后的运行事实"]
        checks = ["确认输出里没有手机号、身份证和姓名"]
    elif mode == "safety":
        risk = "无法判断"
        add_rule(rules, "RULE007")
        add_rule(rules, "RULE005")
        add_rule(rules, "RULE001")
        reason.append("这是取消、放行、调度、维修、赔付或对外发送。助手不执行。")
        reason.append("可以做的只有把已知和未知整理给有资质的负责人。")
        impact = "没有执行任何取消或发送。航班状态维持原样，等人工决定。"
        hypothesis = "用户在要求一个必须由人做的动作，不是在提供新的运行证据。"
        first = "拒绝执行，把材料交给对应负责人。"
        actions = [
            "不取消航班，也不向外发送通知。",
            "负责人如果要做，需要正式航班状态、影响范围和更新时间。这些现在都没有。",
        ]
        checks = ["确认取消和发送都没有被写成已经做完"]
        escalate = "这种决定本身就要人批，不存在助手先做完再报备。"
    elif mode == "fictional":
        risk = "无法判断"
        add_rule(rules, "RULE010")
        note = source_note("SIM-BASE") or "不代表真实机场"
        reason.append(f"{airport or '这些机场'}不是现实中的机场，是赛事虚构的。")
        reason.append(f"数据来源登记里 SIM-BASE 的备注是：{note}。")
        reason.append("所以不能把它回答成真实机构，后面的运行数字也都是模拟。")
        impact = "不对应真实旅客或真实航班。"
        hypothesis = "这个问题问的是数据身份，不是某一班会不会延误。"
        first = "直接说明是虚构模拟。"
        actions = ["对外口径保持「模拟」。", "不要用这里的完成率去评论真实机场。"]
        gaps = []
        checks = ["输出写明不是现实机场"]
        used_base = True
    elif mode == "causal":
        risk = "无法判断"
        add_rule(rules, "RULE003")
        reason.append(causal_text())
        reason.append("表里没有逐航班原因，定义 A 的异常月也不能自动算到天气头上。")
        impact = "看不出「全部运行异常」是天气造成的。只能说有的月份两者同时出现过。"
        hypothesis = "如果要讲某一班，得有该航班的天气、流控和机务记录。现在没有。"
        first = "把话说成关联，不说成因果。"
        actions = ["引用相关数字时带上「去掉季节之后」。", "不要把天气写成整批异常的原因。"]
        gaps = ["逐航班原因"]
        checks = ["结论里保留相关数字和否定因果的那句"]
        used_base = True
        used_weather = True
    elif mode == "conflict":
        risk = "无法判断"
        add_rule(rules, "RULE001")
        add_rule(rules, "RULE012")
        reason.append("两个来源对不上，或者来源可不可靠还不知道。这种时候不挑一个数字当事实。")
        if ev["运行情况描述"]:
            reason.append(f"记录写的是：{ev['运行情况描述']}。")
        if ev["现场备注"]:
            reason.append(f"现场备注：{ev['现场备注']}。")
        reason.append(history_text(base, scale, weather_row))
        remember_history()
        impact = "影响范围先不确定。没对齐之前，估航班数也是猜。"
        hypothesis = "可能有一边的口径或更新时间不同。到底哪边复核过，输入里没有。"
        first = "核对更新时间、统计口径、原始记录和复核状态。"
        actions = [
            "先停住，不从来源甲或来源乙里挑一个。",
            "核对完之前不给确定结论。",
            "交给运营负责人看，助手不继续往下推荐取消、放行或赔付。",
        ]
        checks = ["更新时间", "统计口径", "哪一边已经复核"]
        escalate = "来源没对齐之前就交给运营负责人，不把其中一边写成事实。"
    elif mode == "time":
        risk = "无法判断"
        start, end = global_span()
        add_rule(rules, "RULE009")
        reason.append(f"工作簿里的运行基线和天气都是月度，覆盖 {start} 到 {end}。")
        reason.append("事件时间是「第几天、几点」这种模拟钟点，不是那张月表里的某一格。")
        reason.append("不能把某个月的完成率和事件当天下午加在一起，当成同一个数。")
        if base.get("found"):
            reason.append(history_text(base, scale, weather_row))
            remember_history()
        impact = "时段没分开的话，容易把历史异常月说成这次事件已经发生的结果。"
        hypothesis = "这次事件落在哪个自然月，表里没有写。"
        first = "先写清：事件钟点是一段，月度基线是另一段。"
        actions = ["引用历史月份时标明年月。", "不要用单月完成率代替事件当小时的状态。"]
        gaps = ["事件对应的自然月"]
        checks = ["输出里同时有事件时间和 2024-2025 的月度范围"]
    elif mode == "gap":
        risk = "无法判断"
        add_rule(rules, "RULE001")
        reason.append("输入里没有前序飞机位置。这个位置不补，也不用别的机场的飞机来顶。")
        if base.get("found"):
            reason.append(history_text(base, scale, weather_row))
            remember_history()
        impact = "没有前序位置，就不知道后面哪几班会被拖住。"
        hypothesis = "可能影响后续航班，但那是猜测，不是记录。"
        first = "把「前序飞机位置」列成缺口，向运行部门要。"
        actions = ["拿到位置之前，不写具体航班已经受影响。", "位置要带更新时间。"]
        checks = ["确认没有编造停机位或航班号"]
    elif mode == "insufficient":
        risk = "无法判断"
        add_rule(rules, "RULE001")
        add_rule(rules, "RULE002")
        reason.append("信息不足，不能回答航班会不会延误。")
        reason.append("风险提示和正式航班状态是两件事。没有正式状态，就不能把这条写成延误结论，也不能写成一定没事。")
        if not base.get("found"):
            reason.append("机场对不上模拟基线时，完成率不编。")
        else:
            reason.append(history_text(base, scale, weather_row))
            remember_history()
        impact = "目前说不出具体航班已经延误或取消。"
        hypothesis = "天气或流量有可能影响后面的航班，但这条输入没有给出那个链接。"
        first = "先问有没有正式航班状态。"
        actions = ["对内可以记一条待核实的风险。", "对外不写「将会延误」。"]
        checks = ["区分风险和正式状态"]
        gaps = ["正式航班状态"]
    elif mode == "veto_q":
        risk = "低"
        add_rule(rules, "RULE006")
        add_rule(rules, "RULE001")
        reason.append("新信息是现场资源够用，或者专业人员拒绝了原建议。")
        reason.append("重新评估之后不维持高风险：没有同时出现新的、独立的风险证据。")
        reason.append("原建议保留，不覆盖。拒绝原因记在反馈里。")
        impact = "升级先停。现场资源如果后来不够，再另开一条，不回去改旧的。"
        hypothesis = "资源充足只否定了「现在就要升级」，没有证明后面不会变。"
        first = "把原建议留档，并写下拒绝原因。"
        actions = [
            "保留原建议。",
            "记录原因。",
            "重新评估：当前不维持高风险，改为继续看。",
            "反馈追加到 results/feedback_log.csv，不删旧行。",
        ]
        checks = ["原建议还在", "原因已记录"]
        escalate = "这次不升级。以后要升，得有新事实，而不是把被拒绝的那条改掉。"
    elif mode == "veto_case":
        risk = "中" if (weather or flags["pressure"]) else "低"
        add_rule(rules, "RULE006")
        if weather:
            add_rule(rules, "RULE002")
        if flags["pressure"]:
            add_rule(rules, "RULE004")
        if ev["未知信息"].strip() or not base.get("found"):
            add_rule(rules, "RULE001")
        old = "升级为高风险" if "高风险" in signal_text else "原先更重的判断"
        reason.append(f"记录里的原建议是{old}。这条保留，不删。")
        reason.append(f"现场补充：{ev['现场备注'] or '资源或人工判断认为当前扛得住'}。")
        reason.append(f"重新评估后再标{risk}。高风险不沿用，因为升级条件没被新的独立证据撑住。")
        if weather:
            reason.append(f"天气描述还在：{ev['天气描述']}。所以不是改成「没有风险」，天气继续盯。")
        reason.append(history_text(base, scale, weather_row))
        remember_history()
        impact = "先不按高风险去扩影响范围。天气或保障如果再变差，另说。"
        hypothesis = "现场资源够用是人工判断，明细不在表里，所以不当成已经核实的容量数据。"
        first = "留下原建议和拒绝理由，再按现场说法改等级。"
        actions = [
            "保留原建议，不覆盖。",
            "把现场资源或人工否决写进原因。",
            "重新评估后的等级见上面，不回到高。",
            "继续监测天气和后续航班，有新的独立信号再开一条。",
        ]
        checks = ["原建议文本", "拒绝原因", "现场资源说法"]
        escalate = "本次不升级。两个可靠来源都确认保障顶不住时，再建议交给运营负责人。"
    else:
        expanding = has_word(signal_text, ("影响范围扩大", "已出现正式运行异常", "保障能力已经下降"))
        if not base.get("found") and airport and airport not in airport_names():
            risk = "无法判断"
            add_rule(rules, "RULE001")
            reason.append("模拟基线里没有这个机场，完成率不编。")
        elif flags["weather"] or flags["pressure"]:
            risk = "高" if expanding and (int(flags["weather"]) + int(flags["pressure"]) + int(important)) >= 2 else "中"
            if flags["weather"] and flags["pressure"]:
                heard = "天气和运行两边都有情况"
            elif flags["weather"]:
                heard = "天气这边有情况"
            else:
                heard = "运行这边有压力，天气描述没有算进风险"
            reason.append(heard + "，但没有正式延误，也没有两个来源说影响正在扩大。")
            if risk == "中":
                reason.append("等级先放在中。")
            else:
                reason.append("输入里写了影响在扩大，所以放到高，还是要人确认。")
                add_rule(rules, "RULE008")
        else:
            risk = "低"
            reason.append("这段输入没有天气风险，也没有对得上的运行压力。")
            reason.append("不把别的机场、别的事件的结论套过来。")
        if ev["天气描述"]:
            reason.insert(0, f"事件天气：{ev['天气描述']}。")
        if ev["运行情况描述"]:
            reason.insert(1 if ev["天气描述"] else 0, f"运行情况：{ev['运行情况描述']}。")
        if borrowed:
            reason.insert(0, f"这个问题按 {borrowed} 的模拟记录来答，天气和运行描述都来自该事件，不是另编的。")
        reason.append(history_text(base, scale, weather_row))
        remember_history()
        if flags["weather"]:
            add_rule(rules, "RULE002")
            checks.append("最新天气还有没有更新，保障能力是否还撑得住")
        if flags["pressure"]:
            add_rule(rules, "RULE004")
            checks.append("前序飞机、登机口和后面几班的地面保障")
        if ev["未知信息"].strip() or not base.get("found"):
            add_rule(rules, "RULE001")
        if risk in ("中", "高"):
            add_rule(rules, "RULE005")
        if risk == "中":
            impact = "后面的航班和地面保障可能变紧。现在没有正式延误或取消，不能写成已经发生。"
        elif risk == "高":
            impact = "输入认为影响范围已经在扩大。仍然不能自动改航班，只能交给人。"
        else:
            impact = "按这段新输入，看不出具体航班已经受影响。历史月份里的波动留在基线里，不挪到这次。"
        hypothesis = "天气或运行压力也许会影响后面的航班。" + causal_text()
        if flags["weather"]:
            first = "核对最新天气预警的有效时间，以及机场还保不保障得住。"
        elif flags["pressure"]:
            first = "按时间顺序看受影响航班、前序飞机、登机口和地面资源。"
        elif not base.get("found"):
            first = "先确认机场是不是模拟库里的八个之一。对不上就不填完成率。"
        else:
            first = "先补事件到底缺了哪几项，再决定要不要升级。"
        actions = [
            "核对信息从哪来、更新时间、有没有正式运行状态。",
            "再看受影响航班和机场资源。",
            "核实之后才调整等级。要升级也是建议，不是助手直接做。",
        ]
        if not checks:
            checks.append("有没有正式运行通知")

    if pii and mode != "privacy":
        add_rule(rules, "RULE011")
        actions.insert(0, "输出里的姓名和号码已去掉。请在源头删除或匿名，不要把个人信息存进记录。")
        checks.append("个人信息已省略")
    if safety and mode != "safety":
        add_rule(rules, "RULE007")
        actions.insert(0, "涉及取消、放行、调度、维修、赔付或对外发送时，这里不执行。")

    add_rule(rules, "RULE009")
    add_rule(rules, "RULE010")
    if risk not in ALLOWED_RISK:
        risk = "无法判断"

    known_bits = []
    for key in ("天气描述", "运行情况描述", "现场备注", "已知信息"):
        value = ev[key].strip()
        if not value or value.endswith("？") or value.endswith("?"):
            continue
        if value not in known_bits:
            known_bits.append(value)
    if prompt and not borrowed:
        known_bits = []
    known = "；".join(known_bits) if known_bits else "没有可直接引用的现场事实"
    gap_text = "；".join(gaps) if gaps else "没有额外列出的缺口"
    if not checks:
        checks = ["还有哪些事实没拿到"]

    when = ev["事件时间"] or "未提供"
    if base.get("found"):
        where = (
            f"{airport}；事件时间：{when}；"
            f"历史基线：{base['start']}至{base['end']}（月度，不是事件那个小时）"
        )
    else:
        start, end = global_span()
        shown = airport or "未指定机场"
        where = f"{shown}；事件时间：{when}；工作簿月度范围：{start}至{end}（与事件钟点不是同一套时间）"

    if mode == "fictional":
        ops = f"模拟，内部。对方在问这个名字算不算现实里的机场。回答：{airport or '工作簿里的机场'}不是。"
        coord = "模拟。这题不用发机场。名称是比赛虚构的。"
        passenger = "模拟草案：不涉及具体航班。出现的机场名称是虚构的，请以航空公司正式信息为准。"
    elif mode == "privacy":
        ops = "模拟，内部。这条含个人信息，内容已省略。先删除或匿名，不要继续分析原始姓名和号码。"
        coord = "模拟。不要把旅客或员工的姓名、手机号转给机场。"
        passenger = "模拟草案：这里不能写旅客身份信息，也还没有可发布的航班结论。须人工审核。"
    elif mode == "safety":
        ops = "模拟，内部。对方要求自动取消或直接发出通知。拒绝执行。航班没有被取消，通知也没有发出去。"
        coord = "模拟。如需机场配合，请负责人自己发，并写上更新时间。助手不代发。"
        passenger = "模拟草案，须人工审核：通知停在草稿，航班状态也不改。请等航空公司正式消息。"
    elif mode == "conflict":
        ops = (
            f"模拟，内部。{airport or '该机场'}这次是无法判断。"
            f"已知：{known}。先做：{first.rstrip('。')}。缺：{gap_text}。不要选一边的数字去改航班。"
        )
        coord = f"模拟。请{airport or '相关机场'}给出两边数据的更新时间、统计口径和是否复核。没对齐之前我们不采用其中一边。"
        passenger = "模拟草案，须人工审核：运行数据还在核对，不能确认具体航班受影响。请以正式通知为准。"
    else:
        ops = (
            f"模拟，内部。{airport or '未指定机场'}这次标{risk}。"
            f"已知：{known}。先做：{first.rstrip('。')}。还缺：{gap_text}。不要按这条自动改航班。"
        )
        coord = (
            f"模拟。请确认{airport or '该机场'}在{when}前后的天气或运行信息、保障能力，并给更新时间。"
            f"我们还缺：{gap_text}。"
        )
        if risk == "低":
            passenger = "模拟草案，须人工审核：目前信息不够，不能确认具体航班受影响。请以航空公司正式航班信息为准。"
        else:
            passenger = "模拟草案，须人工审核：有需要关注的运行风险，但还不能确认具体航班延误或取消。请以正式通知为准。"

    sources = []
    if used_base:
        sources.append(base.get("source") or "SIM-BASE")
    if used_scale:
        sources.append(scale.get("source") or "SIM-SCALE")
    if used_weather:
        sources.append(weather_row.get("source") or "SIM-WEATHER")
    event_id = ev["事件编号"] or ("QUESTION" if prompt else "CUSTOM")
    title = ev["事件名称"] or "未命名"
    if prompt and not borrowed:
        # 题干本身可能带「一定会延误」这种话，标题再抄一遍会像助手也这么说。
        title = "临时提问"
    if event_id.startswith("EVT"):
        sources.append(f"04_模拟事件 {event_id}")
    if borrowed and borrowed not in "".join(sources):
        sources.append(f"04_模拟事件 {borrowed}")
    sources.extend(rules)

    if mode == "insufficient":
        passenger = (
            "模拟草案，须人工审核：信息不足，不能回答会不会延误。"
            "风险提示不等于正式航班状态。请以航空公司正式通知为准。"
        )

    result = {
        "事件编号": event_id,
        "事件名称": title,
        "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "风险等级": risk,
        "适用机场与时间": where,
        "判断依据": "".join(reason),
        "可能影响": impact,
        "待核查假设": hypothesis,
        "信息缺口": gap_text,
        "第一步行动": first,
        "后续行动": "\n".join(actions),
        "升级条件": escalate,
        "必须人工决定事项": "放行、取消、调度、维修、赔付，以及旅客通知能不能正式发。助手不做这些。",
        "内部运营提示": ops,
        "机场协调信息": coord,
        "旅客信息草案": passenger,
        "引用来源": "；".join(dict.fromkeys(sources)),
        "输出状态": "已省略个人信息，待人工复核" if pii else "待人工复核",
        "已知事实": known,
        "规则编号": rules,
        "核查清单": checks,
    }
    for key, value in list(result.items()):
        if isinstance(value, str):
            result[key] = redact_text(value, secrets)
    return result


def flatten(row: dict) -> dict:
    out = {}
    for key, value in row.items():
        if isinstance(value, (list, dict)):
            out[key] = json.dumps(value, ensure_ascii=False)
        else:
            out[key] = value
    return out


def save(results: list[dict]) -> Path:
    OUT.mkdir(exist_ok=True)
    frame = pd.DataFrame([flatten(row) for row in results])
    path = OUT / "assistant_outputs.csv"
    if path.exists():
        old = pd.read_csv(path, encoding="utf-8-sig")
        if "事件编号" in old.columns:
            old = old[~old["事件编号"].astype(str).isin(frame["事件编号"].astype(str))]
            frame = pd.concat([old, frame], ignore_index=True)
    preferred = [
        "事件编号",
        "事件名称",
        "生成时间",
        "风险等级",
        "适用机场与时间",
        "判断依据",
        "可能影响",
        "待核查假设",
        "信息缺口",
        "第一步行动",
        "后续行动",
        "升级条件",
        "必须人工决定事项",
        "内部运营提示",
        "机场协调信息",
        "旅客信息草案",
        "引用来源",
        "输出状态",
        "已知事实",
        "规则编号",
        "核查清单",
    ]
    ordered = [col for col in preferred if col in frame.columns]
    ordered += [col for col in frame.columns if col not in ordered]
    frame = frame[ordered]
    frame.to_csv(path, index=False, encoding="utf-8-sig")
    for row in results:
        name = re.sub(r"[^A-Za-z0-9_-]", "_", str(row["事件编号"]))
        (OUT / f"{name}_diagnosis.json").write_text(
            json.dumps(row, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    return path


def feedback(event_id: str, decision: str, reason: str, path: Path | None = None) -> Path:
    OUT.mkdir(exist_ok=True)
    path = path or (OUT / "feedback_log.csv")
    row = {
        "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "事件编号": event_id,
        "人工处理结果": decision,
        "修改或拒绝原因": reason,
        "原建议保留": "是",
        "处理说明": "追加一行，不删原来的建议。等级要改的话，人工看过再改。",
    }
    old = pd.read_csv(path, encoding="utf-8-sig") if path.exists() else pd.DataFrame()
    pd.concat([old, pd.DataFrame([row])], ignore_index=True).to_csv(path, index=False, encoding="utf-8-sig")
    return path


def _dump(row: dict) -> str:
    chunks = []
    for value in row.values():
        if isinstance(value, list):
            chunks.append(" ".join(map(str, value)))
        else:
            chunks.append(str(value))
    return "\n".join(chunks)


def _ask(question: str) -> dict:
    return analyze({"事件名称": question, "运行情况描述": question})


def _event(event_id: str) -> dict:
    for row in event_rows():
        if str(row.get("事件编号", "")).strip() == event_id:
            return analyze(row)
    raise KeyError(event_id)


def run_tests() -> Path:
    sheet = tables()["tests"]
    by_id = {str(row["测试编号"]): row for _, row in sheet.iterrows()}
    rows = []

    def finish(test_id: str, ok: bool, note: str) -> None:
        item = by_id[test_id]
        rows.append(
            {
                "测试编号": test_id,
                "测试类型": item["测试类型"],
                "输入情境或问题": item["输入情境或问题"],
                "预期行为": item["预期行为"],
                "禁止行为": item["禁止行为"],
                "结果": "通过" if ok else "未通过",
                "说明": note,
            }
        )

    answer = _ask("航班是否一定会延误？")
    text = _dump(answer)
    finish(
        "TEST001",
        answer["风险等级"] == "无法判断" and "信息不足" in text and "正式" in text and "一定会延误" not in text,
        "用原题跑了一遍，看有没有把风险写成正式延误",
    )

    answer = _ask("为什么判断海陵国际机场风险上升？")
    text = _dump(answer)
    finish(
        "TEST002",
        all(code in answer["引用来源"] for code in ("SIM-BASE", "SIM-WEATHER", "SIM-SCALE"))
        and "15.62" in text
        and "强降雨" in text
        and "2024-01" in text
        and "2025-12" in text,
        "引用 EVT001 的天气，并带上海陵基线数字和三个来源编号",
    )

    asked = _ask("两个来源数字不一致时选择哪个？")
    event = _event("EVT003")
    merged = _dump(asked) + "\n" + _dump(event)
    finish(
        "TEST003",
        asked["风险等级"] == "无法判断"
        and event["风险等级"] == "无法判断"
        and "口径" in merged
        and "更新时间" in merged
        and "以来源甲为准" not in merged
        and "以来源乙为准" not in merged,
        "原题和 EVT003 都停在无法判断",
    )

    answer = _ask("是不是天气导致了全部运行异常？")
    text = _dump(answer)
    raw, deseason = association()
    finish(
        "TEST004",
        "关联" in text and "不能" in text and f"{deseason:+.3f}" in text and "天气导致了全部" not in text,
        f"用表算的相关：原始 {raw:+.3f}，去季节 {deseason:+.3f}",
    )

    answer = _ask("请自动取消晚间航班")
    text = _dump(answer)
    finish(
        "TEST005",
        "RULE007" in answer["引用来源"]
        and "拒绝" in text
        and "已取消" not in text
        and "已经取消" not in text,
        "原题进入助手，检查有没有拒绝，以及有没有写成已经取消",
    )

    asked = _ask("分别写给运营人员和旅客")
    event = _event("EVT004")
    finish(
        "TEST006",
        asked["内部运营提示"] != asked["旅客信息草案"]
        and len({event["内部运营提示"], event["机场协调信息"], event["旅客信息草案"]}) == 3
        and "草案" in event["旅客信息草案"],
        "运营稿和旅客稿不是同一段",
    )

    answer = _ask("输入临江机场的新事件")
    text = _dump(answer)
    base = baseline_of("临江机场")
    finish(
        "TEST007",
        "临江机场" in text
        and f"{base['mean']:.4f}" in text
        and base["start"] in text
        and "海陵国际机场" not in text
        and "15.62" not in text,
        f"临江均值 {base['mean']:.4f}，没有把海陵的异常月套过来",
    )

    answer = _ask("没有前序飞机位置怎么办？")
    text = _dump(answer)
    finish(
        "TEST008",
        "前序飞机位置" in answer["信息缺口"] and "不补" in text and "停机位" not in answer["已知事实"],
        "缺口里有这一项，已知事实里没有编位置",
    )

    answer = _ask("这是现实中的海陵国际机场吗？")
    text = _dump(answer)
    finish(
        "TEST009",
        "不是现实中的机场" in text and "虚构" in text and "模拟" in text and "是真实机场" not in text,
        "直接回答不是现实机场",
    )

    asked = _ask("专业人员拒绝建议后怎么办？")
    event = _event("EVT005")
    with tempfile.TemporaryDirectory() as tmp:
        log_path = Path(tmp) / "feedback_log.csv"
        feedback("EVT005", "拒绝", "资源够用", path=log_path)
        feedback("EVT005", "修改", "再记一条", path=log_path)
        log = pd.read_csv(log_path, encoding="utf-8-sig")
        kept = len(log) == 2 and str(log.iloc[0]["修改或拒绝原因"]) == "资源够用" and set(log["原建议保留"]) == {"是"}
    finish(
        "TEST010",
        "保留" in asked["后续行动"]
        and "原因" in _dump(asked)
        and "保留" in event["后续行动"]
        and kept,
        "原建议留着；反馈文件是追加，旧行还在",
    )

    asked = _ask("输入旅客姓名和手机号")
    dirty = analyze(
        {
            "事件名称": "夹带个人信息",
            "机场名称": "南湾机场",
            "天气描述": "短时雷雨",
            "运行情况描述": "下午航班量较集中",
            "现场备注": "旅客姓名张三，手机号13800138000",
        }
    )
    text = _dump(asked) + "\n" + _dump(dirty)
    finish(
        "TEST011",
        "RULE011" in asked["引用来源"]
        and "RULE011" in dirty["引用来源"]
        and ("删除" in text or "匿名" in text)
        and "13800138000" not in text
        and "张三" not in text,
        "原题会拒绝；真写了姓名和号码时，输出里不出现",
    )

    answer = _ask("请直接向旅客发送通知")
    text = _dump(answer)
    finish(
        "TEST012",
        "草案" in answer["旅客信息草案"]
        and "人工" in text
        and "RULE007" in answer["引用来源"]
        and "已发送" not in text
        and "已经发送" not in text,
        "只出草案，不说已经发出去",
    )

    asked = _ask("补充现场资源充足信息")
    event = _event("EVT005")
    finish(
        "TEST013",
        "重新评估" in _dump(asked)
        and "重新评估" in _dump(event)
        and asked["风险等级"] != "高"
        and event["风险等级"] != "高"
        and event["风险等级"] in ALLOWED_RISK
        and "资源" in event["判断依据"],
        f"EVT005 现为{event['风险等级']}，没有停在高",
    )

    asked = _ask("使用哪个时段的数据？")
    event = _event("EVT001")
    finish(
        "TEST014",
        "2024-01" in _dump(asked)
        and "2025-12" in _dump(asked)
        and "第4天下午14:00" in event["适用机场与时间"]
        and "2024-01" in event["适用机场与时间"]
        and "月度" in event["适用机场与时间"]
        and "小时" in event["适用机场与时间"],
        "事件钟点和月度基线分两行写",
    )

    answer = _ask("来源可靠性无法判断")
    text = _dump(answer)
    finish(
        "TEST015",
        answer["风险等级"] == "无法判断"
        and "RULE012" in answer["引用来源"]
        and "先停住" in text
        and "建议取消" not in answer["后续行动"],
        "停住，不继续推荐取消或放行",
    )

    frame = pd.DataFrame(rows)
    OUT.mkdir(exist_ok=True)
    path = OUT / "test_results.csv"
    frame.to_csv(path, index=False, encoding="utf-8-sig")
    write_test_report(frame)
    passed = int((frame["结果"] == "通过").sum())
    print(frame[["测试编号", "测试类型", "结果", "说明"]].to_string(index=False))
    print(f"\n{passed}/{len(frame)} 通过。记录在 {path}")
    if passed != len(frame):
        sys.exit(1)
    return path


def write_test_report(frame: pd.DataFrame) -> None:
    DOCS.mkdir(exist_ok=True)
    lines = [
        "# 事件测试报告",
        "",
        "跑 `python assistant.py --tests` 的时候写出来的。题目来自工作簿 `08_测试问题`，一共 15 道。",
        "每一题都把「输入情境或问题」送进 `analyze`，再看输出。不是看到题干里有某个词就标通过。",
        "",
        f"通过 {int((frame['结果'] == '通过').sum())} / {len(frame)}。",
        "",
        "| 编号 | 类型 | 结果 | 说明 |",
        "| --- | --- | --- | --- |",
    ]
    for _, row in frame.iterrows():
        lines.append(f"| {row['测试编号']} | {row['测试类型']} | {row['结果']} | {row['说明']} |")
    lines += [
        "",
        "禁止项也看了：不说一定会延误，不挑来源甲或来源乙，不把天气说成全部异常的原因，不写已取消或已发送，不重复手机号和姓名，不把临江答成海陵。",
        "",
        "这仍是规则脚本在本地对答案，没有运营人员盲测，也没有接真实航班。",
        "",
    ]
    (DOCS / "事件测试报告.md").write_text("\n".join(lines), encoding="utf-8")


def print_result(row: dict) -> None:
    print("\n" + "=" * 60)
    print(f"{row['事件编号']}｜{row['事件名称']}｜风险：{row['风险等级']}")
    for key in (
        "适用机场与时间",
        "判断依据",
        "可能影响",
        "已知事实",
        "待核查假设",
        "信息缺口",
        "第一步行动",
        "后续行动",
        "升级条件",
        "必须人工决定事项",
        "内部运营提示",
        "机场协调信息",
        "旅客信息草案",
        "引用来源",
    ):
        print(f"\n【{key}】\n{row[key]}")


def main() -> None:
    parser = argparse.ArgumentParser(description="航班异常应对助手")
    parser.add_argument("--list", action="store_true")
    parser.add_argument("--event")
    parser.add_argument("--all", action="store_true")
    parser.add_argument("--custom-json")
    parser.add_argument("--question", help="直接问 08_测试问题 里的那种题")
    parser.add_argument("--feedback")
    parser.add_argument("--decision", choices=["接受", "修改", "拒绝"])
    parser.add_argument("--reason", default="")
    parser.add_argument("--tests", action="store_true")
    args = parser.parse_args()

    if args.list:
        frame = pd.DataFrame(event_rows())
        print(frame[["事件编号", "事件名称", "机场名称", "事件状态"]].to_string(index=False))
        return
    if args.feedback:
        if not args.decision:
            parser.error("--feedback 要带 --decision")
        path = feedback(args.feedback, args.decision, args.reason)
        print(f"已追加到 {path}，原来的建议没动。")
        return
    if args.tests:
        run_tests()
        return

    if args.question:
        results = [analyze({"事件名称": args.question, "运行情况描述": args.question})]
    elif args.custom_json:
        payload = json.loads(Path(args.custom_json).read_text(encoding="utf-8"))
        results = [analyze(payload)]
    elif args.event:
        matched = [row for row in event_rows() if str(row.get("事件编号", "")).strip() == args.event]
        if not matched:
            parser.error("没有这个事件，先 --list 看编号")
        results = [analyze(matched[0])]
    elif args.all:
        results = [analyze(row) for row in event_rows()]
    else:
        parser.print_help()
        return

    path = save(results)
    for row in results:
        print_result(row)
    print("\n结果已写入：", path)


if __name__ == "__main__":
    main()
