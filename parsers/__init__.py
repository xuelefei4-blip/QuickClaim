"""
parsers/__init__.py
票据解析器导出入口
"""

from .train_parser import parse_train_ticket
from .taxi_parser import parse_taxi_invoice

__all__ = [
    "parse_train_ticket",
    "parse_taxi_invoice",
]