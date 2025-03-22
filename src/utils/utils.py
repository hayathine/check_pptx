import os
from datetime import datetime

def get_datetime_str() -> str:
    """現在の日時を表す文字列を取得する

    Returns:
        str: 現在の日時を表す文字列
    """
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def flatten(obj):
    if isinstance(obj, dict):  # 辞書の場合
        for key, value in obj.items():
            if isinstance(value, (dict, list)):  # 値が辞書やリストの場合、再帰的に展開
                yield from flatten(value)
            else:
                yield key, value
    elif isinstance(obj, list):  # リストの場合
        for item in obj:
            if isinstance(item, (dict, list)):  # 要素が辞書やリストの場合、再帰的に展開
                yield from flatten(item)
            else:
                yield item
    else:
        yield obj