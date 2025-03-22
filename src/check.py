from pptx import Presentation

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

def extract_pptx_content(presentation):
    """
    PowerPointファイル（.pptx）からテキスト、フォント、フォントサイズ、フォントカラー、テキストボックスの位置を抽出する関数
    
    Args:
        presentation (Presentation): python-pptxのPresentationオブジェクト
        
    Returns:
        list: 各スライドの情報を含む辞書のリスト
    """
    
    if presentation is None:
        return []
    
    slides_data = []
    
    for slide_index, slide in enumerate(presentation.slides):
        slide_data = {
            'slide_number': slide_index + 1,
            'shapes': []
        }
        
        for shape in slide.shapes:
            # テキストを含む形状のみを処理
            if not shape.has_text_frame:
                continue
                
            shape_data = {
                'text': '',
                'paragraphs': [],
                'position': {
                    'left': shape.left,
                    'top': shape.top,
                    'width': shape.width,
                    'height': shape.height
                }
            }
            
            for paragraph in shape.text_frame.paragraphs:
                paragraph_text = paragraph.text
                shape_data['text'] += paragraph_text + '\n'
                
                paragraph_info = {
                    'text': paragraph_text,
                    'runs': []
                }
                
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
                    
                    paragraph_info['runs'].append(run_info)
                
                shape_data['paragraphs'].append(paragraph_info)
            
            slide_data['shapes'].append(shape_data)
        
        slides_data.append(slide_data)
    
    return slides_data

def print_pptx_content(slides_data):
    """
    抽出したPowerPointの内容を表示する関数
    
    Args:
        slides_data (list): extract_pptx_content関数で抽出したデータ
    """
    
    if not slides_data:
        print("データがありません。")
        return
    
    for slide in slides_data:
        print(f"\n===== スライド {slide['slide_number']} =====")
        
        for shape_index, shape in enumerate(slide['shapes']):
            print(f"\n-- テキストボックス {shape_index + 1} --")
            print(f"位置: 左={shape['position']['left']}, 上={shape['position']['top']}, "
                    f"幅={shape['position']['width']}, 高さ={shape['position']['height']}")
            
            for para_index, paragraph in enumerate(shape['paragraphs']):
                print(f"\n段落 {para_index + 1}:")
                
                for run_index, run in enumerate(paragraph['runs']):
                    print(f"  テキスト: {run['text']}")
                    print(f"  フォント: {run['font_name']}, サイズ: {run['font_size']}, 色: {run['font_color']}")
                    
                    style = []
                    if run['bold']:
                        style.append('太字')
                    if run['italic']:
                        style.append('斜体')
                    if run['underline']:
                        style.append('下線')
                    
                    if style:
                        print(f"  スタイル: {', '.join(style)}")
                    print("")
