from pptx import Presentation
from src import check
import os
import traceback
import streamlit as st
from dotenv import load_dotenv
from src.utils import logger, utils


# 環境変数の読み込み
load_dotenv()
gemini_api_key = os.getenv("GEMINI_API_KEY")
debug_mode = os.getenv("DEBUG")

# assets ディレクトリのパス
ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")
# 日時の取得
date = utils.get_datetime_str()
# ロガーの設定
logger = logger.setup_logger()
# tempfileディレクトリのパス
temp_dir = os.path.join(ASSETS_DIR,"temp")
def extract_pptx_content(pptx_path):
    logger.info(f"PowerPointファイルの読み込みを開始: {pptx_path}")
    try:
        prs = Presentation(pptx_path)
        content = []
        
        for i, slide in enumerate(prs.slides, 1):
            logger.debug(f"スライド {i} の処理を開始")
            slide_content = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    if shape.text.strip():
                        slide_content.append(shape.text)
            content.append(slide_content)
            logger.debug(f"スライド {i} の処理が完了")
        
        logger.info(f"PowerPointファイルの読み込みが完了: {pptx_path}")
        return content
    except Exception as e:
        logger.error(f"PowerPointファイルの読み込み中にエラーが発生: {str(e)}", exc_info=True)
        raise

def main():

    # Streamlitアプリのタイトル設定
    st.title("パワーポイントファイル読み込みツール")
    
    # ファイルアップローダーの表示
    uploaded_file = st.file_uploader("パワーポイントファイルをアップロードしてください", type=["pptx"])
    
    try:
        if uploaded_file is not None:
            # 一時ファイルとして保存
            filename = f"{date}_{uploaded_file.name}"
            temp_path = os.path.join(temp_dir, filename)
            print(temp_path)
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success(f"ファイル '{uploaded_file.name}' がアップロードされました")
            # ログ出力
            print(f"INFO: ファイル '{uploaded_file.name}' が正常にアップロードされました")
    except Exception as e:
        # エラーメッセージの表示
        st.error(f"ファイルのアップロード中にエラーが発生しました: {str(e)}")
        # ログ出力
        print(f"ERROR: ファイルアップロード中にエラーが発生: {str(e)}")
        print(f"ERROR: 詳細なエラー情報: {traceback.format_exc()}")
        logger.warning(f'pptxが存在しません:{e}')

    
    logger.info("プログラムを開始")
    
    try:
        # 入力ディレクトリの存在確認
        if not os.path.exists(temp_dir):
            logger.warning(f"入力ディレクトリが存在しません: {temp_dir}")
            os.makedirs(temp_dir)
            logger.info(f"入力ディレクトリを作成しました: {temp_dir}")
        
        # 入力ディレクトリ内のすべての.pptxファイルを処理
        pptx_files = [f for f in os.listdir(temp_dir) if f.endswith(".pptx")]
        
        if not pptx_files:
            logger.warning(f"PowerPointファイルが見つかりません: {temp_dir}")
            return
        
        for filename in pptx_files:
            logger.info(f"ファイル処理開始: {filename}")
            
            content = extract_pptx_content(temp_path)
            
            # 各スライドの内容を表示
            for i, slide_content in enumerate(content, 1):
                logger.info(f"スライド {i} の内容:")
                for text in slide_content:
                    logger.info(f"- {text}")
        
        logger.info("すべてのファイルの処理が完了")
    except Exception as e:
        logger.error("予期せぬエラーが発生", exc_info=True)
        raise

if __name__ == "__main__":
    main()
