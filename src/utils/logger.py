import logging
import os
from src.utils import utils
from datetime import datetime
from logging.handlers import RotatingFileHandler



def setup_logger(name: str = __name__) -> logging.Logger:
    """ロガーの設定を行う

    Args:
        name (str): ロガー名（デフォルトは呼び出し元のモジュール名）

    Returns:
        logging.Logger: 設定済みのロガーインスタンス
    """
    # ログ保存用のディレクトリを作成
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)

    # 日時の取得
    date = utils.get_datetime_str()
    
    # ログファイル名を現在の日付で生成
    log_file = os.path.join(log_dir, f"{date}.log")

    # ロガーの作成
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # ログのフォーマット設定
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # ファイルハンドラの設定（ローテーション付き）
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=1024*1024,  # 1MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # コンソールハンドラの設定
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # ハンドラの追加（既存のハンドラを削除してから追加）
    if logger.handlers:
        logger.handlers.clear()
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

