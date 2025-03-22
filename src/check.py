from pptx import Presentation
from google import genai
from google.genai import types
from src.utils import logger
import streamlit as st
import os
from dotenv import load_dotenv

# 環境変数の読み込み
load_dotenv()
MODEL_NAME = os.getenv("GEMINI_API_NAME")

class Checker:
    def __init__(self, api_key: str):
        self.client = genai.Client(api_key=api_key)
        self.logger = logger.setup_logger()
        # assets ディレクトリのパス
        self.assets_dir = os.path.join(os.path.dirname(__file__), "assets")
        # tempfileディレクトリのパス
        self.temp_dir = os.path.join(self.assets_dir,"temp")
        self.slides_text = ""
        self.prompt = ""
        self.temp_path = ""

    def import_pptx(file_path):
        """
        PowerPointファイル（.pptx）を取り込む関数
        
        Args:
            file_path (str): 取り込むPowerPointファイルのパス
            
        Returns:
            Presentation: python-pptxのPresentationオブジェクト
        """
        
        try:
            presentation = Presentation(file_path)
            return presentation
        except Exception as e:
            print(f"ファイルの取り込み中にエラーが発生しました: {e}")
            return None


    def extract_pptx(self):
        """
        PowerPointファイル（.pptx）からテキスト、フォント、フォントサイズ、フォントカラー、テキストボックスの位置を抽出する関数
        
            
        Returns:
            list: 各スライドの情報を含む辞書のリスト
        """
        # PowerPointファイルの取り込み
        try:
            presentation = Presentation(self.temp_path)
        except Exception as e:
            self.logger.error(f"PowerPointファイルの取り込み中にエラーが発生: {str(e)}", exc_info=True)
            raise
        
        slides_data = []
        
        for slide_index, slide in enumerate(presentation.slides):
            texts = []
            for p in slide.placeholders:
                texts.append(p.text[:-1]) # 最後の要素はページ番号なので除外
            slides_data.append(f"スライド {slide_index+1}:タイトル:"+",".join(texts))
                
        return slides_data

    def check_pptx(self) -> str:
        """PowerPointの内容をLLMでチェックする

        Args:
            model : 使用するモデル

        Returns:
            str: LLMからの分析結果
        """
        # プロンプトの作成
        full_prompt = f"""
                        以下のPowerPointの内容を分析し、{self.prompt}の観点から評価してください。

                        PowerPointの内容:
                        {self.slides_text}

                        分析結果について以下の内容を日本語で出力してください。
                        1. 分析結果の出力
                        2. 分析結果を元にslides_textを修正してpython-pptxライブラリのpresenttation.slidesの形式に合わせて出力
                        """
        
        self.logger.info(full_prompt)
        try:
            # LLMによる分析の実行
            response = self.client.models.generate_content(
                model=MODEL_NAME,
                contents=full_prompt
                )
            result = response.text
            
            self.logger.info("PowerPointの内容の分析が完了しました")
            return result
            
        except Exception as e:
            self.logger.error(f"LLMによる分析中にエラーが発生: {str(e)}", exc_info=True)
            raise

    def confilm_pptx(self, uploaded_file, prompt):
        """
        PowerPointファイルの内容を確認する関数
        """
        self.prompt = prompt
        if not uploaded_file:
            st.warning("PowerPointファイルをアップロードしてください。")
            return
        
        if not prompt:
            st.warning("分析の観点を入力してください。")
            return

        # 一時ファイルとして保存
        if not uploaded_file:
            st.error("ファイルがアップロードされていません")
        filename = f"{uploaded_file.name}"
        self.temp_path = os.path.join(self.temp_dir, filename)
        
        # 一時ディレクトリの作成
        os.makedirs(self.temp_dir, exist_ok=True)

        try:
            # ファイルの保存
            with open(self.temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            self.logger.info(f"ファイル '{uploaded_file.name}' がアップロードされました")
        except Exception as e:
            st.error(f"エラーが発生しました: {str(e)}")
        
        # PowerPointの内容を抽出
        content = list(self.extract_pptx()) 
        self.slides_text = ", ".join(content)
        # contentの中身を表示
        st.write(content)
        self.logger.info("内容の確認が完了しました")

    def llm_pptx(self):
        """
        PowerPointファイルの内容を分析する関数
        """
        # プログレスバーの表示
        with st.spinner("PowerPointの内容を分析中..."):
            try:
                # LLMによる分析
                analysis_result = self.check_pptx()
                
                # 結果の表示
                st.subheader("分析結果")
                st.write(analysis_result)
                self.logger.info("分析結果の表示が完了しました")
            except Exception as e:
                st.error(f"エラーが発生しました: {str(e)}")
                self.logger.error(f"予期せぬエラーが発生: {str(e)}", exc_info=True)
            finally:
                # 一時ファイルの削除
                os.remove(self.temp_path)
                self.logger.info("一時ファイルを削除しました")