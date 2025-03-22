import os
from datetime import datetime

def get_datetime_str() -> str:
    """現在の日時を表す文字列を取得する

    Returns:
        str: 現在の日時を表す文字列
    """
    return datetime.now().strftime("%Y%m%d_%H%M%S")
