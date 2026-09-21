"""
parsers/train_parser.py
专职解析 12306 铁路电子客票 / 新版数电铁路发票
"""
import re
from pathlib import Path

PY_MAP = {
    "nanjingnan": "南京南",
    "nanjing": "南京",
    "lianyungang": "连云港",
    "lianyungangdong": "连云港东",
    "shanghai": "上海",
    "shanghaihongqiao": "上海虹桥",
    "beijing": "北京",
    "beijingnan": "北京南",
    "xuzhou": "徐州",
    "xuzhoudong": "徐州东",
    "suzhou": "苏州",
    "wuxi": "无锡",
    "changzhou": "常州",
    "zhenjiang": "镇江",
    "hangzhou": "杭州",
    "hangzhoudong": "杭州东",
}


def parse_train_ticket(text: str, file_path: str) -> dict:
    result = {
        "valid": False,
        "type": "高铁",
        "date": "",
        "amount": 0.0,
        "desc": "",
        "raw_path": file_path
    }

    if not text:
        return result

    # 排除典型的打车发票词，防止误入
    if "客运服务费" in text or "领行" in text:
        return result

    # 1. 提取金额
    amt_match = re.search(r'[¥￥]\s*([0-9]+\.[0-9]{2})', text)
    if amt_match:
        result["amount"] = float(amt_match.group(1))
    else:
        amt_fallback = re.search(r'([0-9]+\.[0-9]{2})\s*元', text)
        if amt_fallback:
            result["amount"] = float(amt_fallback.group(1))

    # 2. 提取乘车日期 (匹配 '2026 08 29 18 19'、'2026-08-29'、'2026年08月29日')
    date_matches = re.findall(r'(\d{4})[\s年\-\/\.](\d{2})[\s月\-\/\.](\d{2})', text)
    if date_matches:
        y, m, d = date_matches[0]
        result["date"] = f"{y}-{m}-{d}"

    # 3. 提取正规车次（必须以字母 G/D/C/Z/T/K/Y/S 开头，杜绝抓取纯数字干扰项）
    train_num = ""
    train_match = re.search(r'\b([GDCZTKYS]\d{1,4})\b', text, re.IGNORECASE)
    if train_match:
        train_num = train_match.group(1).upper()

    # 4. 提取站名区间 (针对拼音行清洗空格匹配)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    found_cities = []
    dep_city = ""

    for idx, raw_line in enumerate(lines):
        line_clean = re.sub(r'[^a-zA-Z]', '', raw_line).lower()
        if line_clean in PY_MAP:
            city_name = PY_MAP[line_clean]
            if city_name not in found_cities:
                found_cities.append(city_name)
            if idx > 0 and train_num and train_num in lines[idx - 1]:
                dep_city = city_name

    route_str = ""
    if len(found_cities) >= 2:
        if dep_city and dep_city in found_cities:
            arr_city = [c for c in found_cities if c != dep_city][0]
            route_str = f"{dep_city}-{arr_city}"
        else:
            route_str = f"{found_cities[1]}-{found_cities[0]}"
    elif len(found_cities) == 1:
        route_str = found_cities[0]
    else:
        cn_match = re.findall(r'([\u4e00-\u9fa5]{2,6})\s*(?:站)?\s*[-—至到]\s*([\u4e00-\u9fa5]{2,6})\s*(?:站)?', text)
        if cn_match:
            route_str = f"{cn_match[0][0]}-{cn_match[0][1]}"

    # 组合行程说明
    parts = []
    if route_str:
        parts.append(route_str)
    if train_num:
        parts.append(train_num)

    result["desc"] = "_".join(parts) if parts else (train_num or "高铁票")

    # 严格判定：必须同时具有明确铁路特征（12306/95306/正规字母车次），且金额与日期有效
    has_train_proof = bool(train_num or "12306" in text or "95306" in text)
    if has_train_proof and result["amount"] > 0 and result["date"]:
        result["valid"] = True

    return result