"""
parsers/taxi_parser.py
出租车/网约车专用解析模块：
支持「电子行程单（逐单明细拆解）」与「数电发票」
"""
import re


def clean_str(s: str) -> str:
    """清除空白字符"""
    return re.sub(r'\s+', '', s)


def _parse_itinerary(text: str, pdf_path: str = "") -> dict:
    """专职解析行程单：提取每日上车日期与对应单笔明细"""
    result = {
        "valid": False,
        "type": "打车",
        "is_itinerary": True,
        "date": "",
        "amount": 0.0,
        "desc": "打车行程单",
        "sub_items": [],
        "raw_path": pdf_path
    }

    # 1. 提取总金额
    amt_match = re.search(r'合计\s*[¥￥]?\s*([0-9]+\.[0-9]{2})\s*元?', text)
    if not amt_match:
        amt_match = re.search(r'[¥￥]\s*([0-9]+\.[0-9]{2})', text)
    if amt_match:
        try:
            result["amount"] = float(amt_match.group(1))
        except ValueError:
            result["amount"] = 0.0

    # 2. 提取截止/申请日期
    date_match = re.search(r'至\s*(\d{4})[\s年\-\/\.](\d{1,2})[\s月\-\/\.](\d{1,2})', text)
    if not date_match:
        date_match = re.search(r'申请时间[:：\s]*(\d{4})[\s年\-\/\.](\d{1,2})[\s月\-\/\.](\d{1,2})', text)
    if date_match:
        y, m, d = date_match.groups()
        result["date"] = f"{y}-{int(m):02d}-{int(d):02d}"

    # 3. 提取服务商
    provider = "高德打车" if "高德" in text else ("T3出行" if ("T3" in text.upper() or "领行" in text) else ("滴滴打车" if "滴滴" in text else "打车"))

    # 4. 逐项抓取单笔行程：日期 (YYYY-MM-DD) ... 金额 (XX.XX元)
    item_pattern = re.compile(r'(\d{4}-\d{2}-\d{2})\s+\d{2}:\d{2}[\s\S]*?([0-9]+\.[0-9]{2})\s*元')

    for m in item_pattern.finditer(text):
        sub_date, sub_amt = m.groups()
        amt_val = float(sub_amt)
        
        if result["amount"] > 0 and abs(amt_val - result["amount"]) < 0.001:
            continue

        result["sub_items"].append({
            "valid": True,
            "type": "打车",
            "date": sub_date.strip(),
            "amount": amt_val,
            "desc": f"{provider}(有票)",
            "raw_path": pdf_path,
            "is_sub_item": True
        })

    # 优化点：若没有抓到总额，则通过单笔小计自动求和
    if result["amount"] == 0.0 and result["sub_items"]:
        result["amount"] = round(sum(item["amount"] for item in result["sub_items"]), 2)
        if not result["date"]:
            result["date"] = max(item["date"] for item in result["sub_items"])

    sub_count = len(result["sub_items"])
    count_match = re.search(r'共计\s*(\d+)\s*单', text)
    count_str = f"({count_match.group(1)}单)" if count_match else (f"({sub_count}单)" if sub_count else "")
    result["desc"] = f"{provider}行程单{count_str}"

    # 优化点：只要提取到子单项，或者总金额与日期齐全，即为有效
    if (result["amount"] > 0 and result["date"]) or sub_count > 0:
        result["valid"] = True

    return result


def parse_taxi_invoice(text: str, pdf_path: str = "") -> dict:
    """打车发票与行程单分流主入口"""
    if not text:
        return {"valid": False, "type": "打车", "date": "", "amount": 0.0, "desc": "客运服务费", "raw_path": pdf_path}

    if "行程单" in text or "ITINERARY" in text.upper():
        return _parse_itinerary(text, pdf_path)

    result = {
        "valid": False,
        "type": "打车",
        "is_itinerary": False,
        "date": "",
        "amount": 0.0,
        "desc": "客运服务费",
        "sub_items": [],
        "raw_path": pdf_path
    }

    # 抓取开票日期
    date_match = re.search(r'开票日期[：:\s]*(\d{4})[\s年\-](\d{1,2})[\s月\-](\d{1,2})', text)
    if not date_match:
        date_match = re.search(r'(\d{4})[\s年\-](\d{1,2})[\s月\-](\d{1,2})(?:日)?', text)
    if date_match:
        y, m, d = date_match.groups()
        result["date"] = f"{y}-{int(m):02d}-{int(d):02d}"

    # 抓取发票金额
    amt_match = re.search(r'（小写）[¥￥\s]*(\d+[\s\.]*\d{2})', text)
    if not amt_match:
        amt_match = re.search(r'\(小写\)[¥￥\s]*(\d+[\s\.]*\d{2})', text)
    if not amt_match:
        amt_match = re.search(r'价税合计.*?([¥￥]?\s*\d+\.\d{2})', text)
    if amt_match:
        raw_val = clean_str(amt_match.group(1)).replace("¥", "").replace("￥", "")
        try:
            result["amount"] = float(raw_val)
        except ValueError:
            result["amount"] = 0.0

    if "领行" in text or "T3" in text.upper():
        result["desc"] = "客运服务费(T3出行)"
    elif "滴滴" in text:
        result["desc"] = "客运服务费(滴滴出行)"
    elif "高德" in text:
        result["desc"] = "客运服务费(高德打车)"
    else:
        seller = re.search(r'名\s*称\s*[:：]\s*([\u4e00-\u9fa5]{4,20}(?:公司|出行|客运))', text)
        result["desc"] = f"客运服务费({seller.group(1)[:4]})" if seller else "客运服务费"

    if result["date"] and result["amount"] > 0:
        result["valid"] = True

    return result