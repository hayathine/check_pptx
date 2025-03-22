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

# assets ディレクトリのパス
ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")

# ロガーの設定
logger = logger.setup_logger()
# tempfileディレクトリのパス
temp_dir = os.path.join(ASSETS_DIR,"temp")
def main():
    # Streamlitアプリのタイトル設定
    st.title("パワーポイントファイル分析ツール")
    
    # ファイルアップローダーの表示
    uploaded_file = st.file_uploader("パワーポイントファイルをアップロードしてください", type=["pptx"])
    
    # テンプレートの定義
    templates = {
        "構成の評価": "スライド全体の流れが論理的に整理されているか？",
        "内容の一貫性": "各スライドの情報が矛盾なくまとまっているか？",
        "デザインの評価": "フォント、カラー、レイアウトが統一されているか？",
        "初出用語のチェック": "専門用語が適切に説明されているか？",
        "スライドごとのポイント": "各スライドの要点が明確か？"
    }

    # # セッションステートを使ってプロンプトを保持
    # if "prompt_text" not in st.session_state:
    #     st.session_state["prompt_text"] = ""

    # # ボタンを配置
    # for name, text in templates.items():
    #     if st.button(name):
    #         st.session_state["prompt_text"] = text

    formatted_template = "\n\n".join([f"{key}: {value}" for key, value in templates.items()])
    # 各テンプレートを表示（コードブロックにすることでコピーしやすく）
    with st.expander(f"📌 テンプレート"):
        st.code(formatted_template, language="plaintext")

    # テキストエリア
    prompt = st.text_area(
        "分析の観点を入力してください",
        # value=st.session_state["prompt_text"],
        placeholder="例：プレゼンテーションの構成、内容の一貫性、デザインの評価など"
    )
    
    # チェックツールの初期化
    checker = Checker(GEMINI_API_KEY)

    if st.button("分析開始"):
        if not uploaded_file:
            st.warning("PowerPointファイルをアップロードしてください。")
            return
        
        if not prompt:
            st.warning("分析の観点を入力してください。")
            return

        try:
            # 一時ファイルとして保存
            filename = f"{uploaded_file.name}"
            temp_path = os.path.join(temp_dir, filename)
            
            # 一時ディレクトリの作成
            os.makedirs(temp_dir, exist_ok=True)
            
            # ファイルの保存
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            logger.info(f"ファイル '{uploaded_file.name}' がアップロードされました")
            print(temp_path)
            # PowerPointの内容を抽出
            content = checker.extract_pptx(temp_path)
            
            # プログレスバーの表示
            with st.spinner("PowerPointの内容を分析中..."):
                # LLMによる分析
                analysis_result = checker.check_pptx(
                    model=MODEL_NAME,
                    content=content, 
                    prompt=prompt)
                
                # 結果の表示
                st.subheader("分析結果")
                st.write(analysis_result)
            
            # 一時ファイルの削除
            os.remove(temp_path)
            logger.info("一時ファイルを削除しました")
            
        except Exception as e:
            st.error(f"エラーが発生しました: {str(e)}")
            logger.error(f"予期せぬエラーが発生: {str(e)}", exc_info=True)
        finally:
            # 一時ファイルの削除（エラー時も確実に削除）
            if 'temp_path' in locals() and os.path.exists(temp_path):
                os.remove(temp_path)

if __name__ == "__main__":
    main()
