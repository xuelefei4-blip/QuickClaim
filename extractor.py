"""
extractor.py
调度中心：严格区分打车与高铁特征，防止误判混淆
"""
import re
from pathlib import Path
from pypdf import PdfReader

from parsers.train_parser import parse_train_ticket
from parsers.taxi_parser import parse_taxi_invoice


def sanitize_pdf_font_descriptors(page):
    """修复 12306 数电票 DengXian 字体双重 FontFile 冲突"""
    seen = set()

    def _walk_and_clean(obj):
        if id(obj) in seen:
            return
        seen.add(id(obj))

        resolved = obj.get_object() if hasattr(obj, "get_object") else obj
        if isinstance(resolved, dict):
            ff_keys = [k for k in list(resolved.keys()) if str(k).startswith("/FontFile")]
            if len(ff_keys) > 1:
                for k in ff_keys[1:]:
                    try:
                        del resolved[k]
                    except Exception:
                        pass
            for v in resolved.values():
                _walk_and_clean(v)
        elif isinstance(resolved, list):
            for item in resolved:
                _walk_and_clean(item)

    _walk_and_clean(page)


def fallback_extract_stream(page) -> str:
    """PDF 指令流底层纯文本提取兜底"""
    try:
        contents = page.get_contents()
        if not contents:
            return ""
        raw_data = contents.get_data().decode('latin1', errors='ignore')
        tokens = re.findall(r'\(([^)]+)\)', raw_data)
        return "\n".join(tokens)
    except Exception:
        return ""


def extract_pdf_text_robust(pdf_path: str) -> str:
    full_text = []
    try:
        reader = PdfReader(pdf_path, strict=False)
        for page in reader.pages:
            sanitize_pdf_font_descriptors(page)
            page_text = ""
            try:
                page_text = page.extract_text() or ""
            except Exception:
                page_text = ""

            if not page_text.strip():
                page_text = fallback_extract_stream(page)

            full_text.append(page_text)
    except Exception as e:
        print(f"读取 PDF 文件异常 {pdf_path}: {e}")
        return ""

    return "\n".join(full_text)


def extract_ticket(pdf_path: str) -> dict:
    path_obj = Path(pdf_path)
    default_result = {
        "valid": False,
        "type": "未知",
        "date": "",
        "amount": 0.0,
        "desc": "",
        "raw_path": str(path_obj.resolve())
    }

    text = extract_pdf_text_robust(pdf_path)
    if not text.strip():
        default_result["desc"] = "无法提取文本"
        return default_result

    # 1. 优先判定：网约车/打车数电发票特征
    taxi_strong_keywords = [
        "客运服务费", "领行汽车", "滴滴", "嘀嘀", "高德打车", 
        "阳光出行", "曹操", "首汽", "出租汽车", "网约车"
    ]
    if any(k in text for k in taxi_strong_keywords):
        res_taxi = parse_taxi_invoice(text, pdf_path)
        if res_taxi.get("valid"):
            return res_taxi

    # 2. 判定：12306 铁路数电票 / 高铁票特征
    train_strong_keywords = ["12306", "95306", "铁路", "客票", "车次", "席别", "检票口"]
    if any(k in text for k in train_strong_keywords):
        res_train = parse_train_ticket(text, pdf_path)
        if res_train.get("valid"):
            return res_train

    # 3. 兜底判定：无强关键字时依次尝试
    res_taxi = parse_taxi_invoice(text, pdf_path)
    if res_taxi.get("valid"):
        return res_taxi

    res_train = parse_train_ticket(text, pdf_path)
    if res_train.get("valid"):
        return res_train

    default_result["desc"] = "非受支持的票据格式"
    return default_result