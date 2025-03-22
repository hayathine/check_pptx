from pptx import Presentation
from google import genai
from google.genai import types
from src.utils import logger

class Checker:
    def __init__(self, api_key: str):
        self.client = genai.Client(api_key=api_key)
        self.logger = logger.setup_logger()

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


    def extract_pptx(self, temp_path: str):
        """
        PowerPointファイル（.pptx）からテキスト、フォント、フォントサイズ、フォントカラー、テキストボックスの位置を抽出する関数
        
        Args:
            temp_path: 取り込むPowerPointファイルのパス
            
        Returns:
            list: 各スライドの情報を含む辞書のリスト
        """
        # PowerPointファイルの取り込み
        try:
            presentation = Presentation(temp_path)
        except Exception as e:
            self.logger.error(f"PowerPointファイルの取り込み中にエラーが発生: {str(e)}", exc_info=True)
            raise
        
        slides_data = []
        
        for slide_index, slide in enumerate(presentation.slides):
            slides_data.append(f"スライド {slide_index+1}")
            for shape in slide.shapes:
                # テキストを含む形状のみを処理
                if not shape.has_text_frame:
                    continue
                    
                shape_data = {
                    'runs': [],
                    'position': {
                        'left': shape.left,
                        'top': shape.top,
                        'width': shape.width,
                        'height': shape.height
                    }
                }
                
                for paragraph in shape.text_frame.paragraphs:
                    
                    paragraph_info = []
                    
                    for run in paragraph.runs:
                        font = run.font
                        color = 'なし'
                        
                        # フォントカラーの取得
                        if font.color.type is not None:
                            if hasattr(font.color, 'rgb') and font.color.rgb is not None:
                                rgb = font.color.rgb
                                color = f'RGB({rgb[0]}, {rgb[1]}, {rgb[2]})'
                        
                        run_info = {
                            'text': run.text,
                            'font_name': font.name,
                            'font_size': font.size.pt if font.size is not None else 'デフォルト',
                            'font_color': color,
                            'bold': font.bold,
                            'italic': font.italic,
                            'underline': font.underline
                        }
                        
                        paragraph_info.append(str(run_info.items()))
                    
                    shape_data['runs'].append(",".join(paragraph_info))
                
            slides_data.append(str(shape_data.items()))
            
        
        return slides_data



    def check_pptx(self, model: str, slides_text: str, prompt: str) -> str:
        """PowerPointの内容をLLMでチェックする

        Args:
            model : 使用するモデル
            slides_text : PowerPointの内容
            prompt : チェックのためのプロンプト

        Returns:
            str: LLMからの分析結果
        """
        # プロンプトの作成
        full_prompt = f"""
                        以下のPowerPointの内容を分析し、{prompt}の観点から評価してください。

                        PowerPointの内容:
                        {slides_text}

                        分析結果について以下の内容を日本語で出力してください。
                        1. 分析結果の出力
                        2. 分析結果を元にslides_textを修正してpython-pptxライブラリのpresenttation.slidesの形式に合わせて出力
                        """
        
        self.logger.info(full_prompt)
        try:
            # LLMによる分析の実行
            response = self.client.models.generate_content(
                model=model,
                contents=full_prompt
                )
            result = response.text
            
            self.logger.info("PowerPointの内容の分析が完了しました")
            return result
            
        except Exception as e:
            self.logger.error(f"LLMによる分析中にエラーが発生: {str(e)}", exc_info=True)
            raise
