from pptx import Presentation
from src.check import Checker
import os
import traceback
import streamlit as st
from dotenv import load_dotenv
from src.utils import logger, utils


# 環境変数の読み込み
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
MODEL_NAME = os.getenv("GEMINI_API_NAME")
debug_mode = os.getenv("DEBUG")
# ロガーの設定
logger = logger.setup_logger()
def main():
    # Streamlitアプリのタイトル設定
    st.title("パワーポイントファイル分析ツール")
    
    # ファイルアップローダーの表示
    uploaded_file = st.file_uploader("パワーポイントファイルをアップロードしてください", type=["pptx"])
    
    # TODO:くくり出す
    # テンプレートの定義
    templates = {
        "構成の評価": "スライド全体の流れが論理的に整理されているか？",
        "内容の一貫性": "各スライドの情報が矛盾なくまとまっているか？",
        "デザインの評価": "フォント、カラー、レイアウトが統一されているか？",
        "初出用語のチェック": "専門用語が適切に説明されているか？",
        "スライドごとのポイント": "各スライドの要点が明確か？"
    }
    formatted_template = "\n\n".join([f"{key}: {value}" for key, value in templates.items()])
    
    # 各テンプレートを表示（コードブロックにすることでコピーしやすく）
    st.write("分析の観点のテンプレート")
    st.code(formatted_template, language="plaintext")

    # テキストエリア
    prompt = st.text_area(
        "分析の観点を入力してください",
        # value=st.session_state["prompt_text"],
        placeholder="例：プレゼンテーションの構成、内容の一貫性、デザインの評価など"
    )
    
    # チェックツールの初期化
    checker = Checker(GEMINI_API_KEY)

    # ボタンの初期状態の設定
    st.session_state.confilm = False

    st.button("内容確認",on_click=checker.confilm_pptx(uploaded_file, prompt))

    # 分析開始
    st.button("分析開始", on_click=checker.llm_pptx)

if __name__ == "__main__":
    main()
