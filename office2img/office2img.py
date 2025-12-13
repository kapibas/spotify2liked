import fitz
from pathlib import Path
import win32com.client
import pythoncom
import time

DPI = 200
IMG_FORMAT = "png"
OUTPUT_DIR = Path("screenshots")
MODE = "images"

WD_EXPORT_FORMAT_PDF = 17

def get_image_name(page_num: int, prefix: str = "page") -> str:
    """Возвращает порядковое имя файла"""
    return f"{prefix}_{page_num:03d}.{IMG_FORMAT}"

def convert_pdf(pdf_path: Path, output_dir: Path) -> None:
    """Конвертирует PDF в изображения"""
    if not pdf_path.exists():
        print(f"  [WARNING] Файл не найден: {pdf_path}")
        return
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    doc = None
    try:
        doc = fitz.open(pdf_path)
        scale = DPI / 72.0
        matrix = fitz.Matrix(scale, scale)
        
        print(f"  Страниц: {len(doc)}")
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            
            img_path = output_dir / get_image_name(page_num + 1, "page")
            pix.save(str(img_path))
            print(f"    [OK] {img_path.name}")
            pix = None
    finally:
        if doc:
            doc.close()

def convert_docx(docx_path: Path, output_dir: Path, word_app) -> None:
    """Конвертирует DOCX через Microsoft Word"""
    if not docx_path.exists():
        print(f"  [WARNING] Файл не найден: {docx_path}")
        return
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    doc = None
    pdf_doc = None
    temp_pdf = output_dir / "temp_export.pdf"
    
    try:
        doc_path = str(docx_path.resolve())
        doc = word_app.Documents.Open(doc_path, ReadOnly=True)
        
        page_count = doc.ComputeStatistics(2)
        print(f"  Страниц: {page_count}")
        
        doc.ExportAsFixedFormat(
            OutputFileName=str(temp_pdf.resolve()),
            ExportFormat=WD_EXPORT_FORMAT_PDF
        )
        
        pdf_doc = fitz.open(str(temp_pdf))
        scale = DPI / 72.0
        matrix = fitz.Matrix(scale, scale)
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            
            img_path = output_dir / get_image_name(page_num + 1, "page")
            pix.save(str(img_path))
            print(f"    [OK] {img_path.name}")
            pix = None
        
    except Exception as e:
        print(f"  [ERROR] Ошибка Word: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if pdf_doc:
            pdf_doc.close()
        
        if temp_pdf.exists():
            try:
                temp_pdf.unlink()
            except:
                pass
        
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass

def extract_pptx_text(pptx_path: Path, all_text_list: list, ppt_app) -> None:
    """Извлекает весь текст из PPTX и добавляет в общий список"""
    if not pptx_path.exists():
        print(f"  [WARNING] Файл не найден: {pptx_path}")
        return
    
    presentation = None
    
    try:
        prs_path = str(pptx_path.resolve())
        presentation = ppt_app.Presentations.Open(prs_path, ReadOnly=True, WithWindow=False)
        
        slide_count = presentation.Slides.Count
        print(f"  Слайдов: {slide_count}")
        
        all_text_list.append(f"\n\n{'='*80}")
        all_text_list.append(f"ПРЕЗЕНТАЦИЯ: {pptx_path.name}")
        all_text_list.append(f"Всего слайдов: {slide_count}")
        all_text_list.append(f"{'='*80}\n")
        
        for slide_num in range(1, slide_count + 1):
            slide = presentation.Slides(slide_num)
            
            all_text_list.append(f"\n{'─'*60}")
            all_text_list.append(f"СЛАЙД {slide_num}")
            all_text_list.append(f"{'─'*60}\n")
            
            slide_text = []
            for shape in slide.Shapes:
                if hasattr(shape, "TextFrame") and shape.HasTextFrame:
                    text_frame = shape.TextFrame
                    if hasattr(text_frame, "TextRange"):
                        text = text_frame.TextRange.Text.strip()
                        if text:
                            slide_text.append(text)
            
            if slide_text:
                all_text_list.append("\n".join(slide_text))
            else:
                all_text_list.append("[Нет текста на слайде]")
            
            print(f"    [OK] Слайд {slide_num} обработан")
        
    except Exception as e:
        print(f"  [ERROR] Ошибка PowerPoint: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if presentation:
            try:
                presentation.Close()
            except:
                pass

def convert_pptx(pptx_path: Path, output_dir: Path, ppt_app) -> None:
    if not pptx_path.exists():
        print(f"  [WARNING] Файл не найден: {pptx_path}")
        return
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    presentation = None
    
    try:
        prs_path = str(pptx_path.resolve())
        presentation = ppt_app.Presentations.Open(prs_path, ReadOnly=True, WithWindow=False)
        
        slide_count = presentation.Slides.Count
        print(f"  Слайдов: {slide_count}")
        
        for slide_num in range(1, slide_count + 1):
            slide = presentation.Slides(slide_num)
            img_path = output_dir / get_image_name(slide_num, "slide")
            
            try:
                width = int(presentation.PageSetup.SlideWidth * DPI / 72)
                height = int(presentation.PageSetup.SlideHeight * DPI / 72)
                
                slide.Export(
                    str(img_path.resolve()),
                    "PNG" if IMG_FORMAT == "png" else "JPG",
                    width,
                    height
                )
                
                print(f"    [OK] {img_path.name}")
            
            except Exception as e:
                print(f"    [WARNING] Ошибка экспорта слайда {slide_num}: {e}")
        
    except Exception as e:
        print(f"  [ERROR] Ошибка PowerPoint: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if presentation:
            try:
                presentation.Close()
            except:
                pass

def process_file(file_path: Path, all_text_list: list = None, word_app=None, ppt_app=None) -> None:
    """Обрабатывает один файл"""
    if not file_path.exists():
        print(f"  [WARNING] Файл не найден: {file_path}")
        return
    
    ext = file_path.suffix.lower()
    output_dir = OUTPUT_DIR / file_path.stem
    
    print(f"\n[FILE] {file_path.name}")
    
    try:
        if MODE == "text":
            if ext in (".pptx", ".ppt"):
                if ppt_app:
                    extract_pptx_text(file_path, all_text_list, ppt_app)
                else:
                    print("  [ERROR] PowerPoint не доступен")
            else:
                print(f"  [WARNING] Режим 'text' работает только с PPTX файлами")
            return
        
        if ext == ".pdf":
            convert_pdf(file_path, output_dir)
        elif ext in (".docx", ".doc"):
            if word_app:
                convert_docx(file_path, output_dir, word_app)
            else:
                print("  [ERROR] Word не доступен")
        elif ext in (".pptx", ".ppt"):
            if ppt_app:
                convert_pptx(file_path, output_dir, ppt_app)
            else:
                print("  [ERROR] PowerPoint не доступен")
        else:
            print(f"  [WARNING] Формат {ext} не поддерживается")
            return
        
        print(f"  [DONE] Готово: {output_dir}")
    
    except Exception as e:
        print(f"  [ERROR] Ошибка: {e}")
        import traceback
        traceback.print_exc()

def check_office_installed(need_word: bool = True, need_ppt: bool = True) -> bool:
    """Проверяет наличие Microsoft Office"""
    try:
        try:
            pythoncom.CoInitialize()
        except pythoncom.com_error:
            pass
        
        if need_word:
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Quit()
                print("[OK] Microsoft Word найден")
            except Exception as e:
                print(f"[ERROR] Microsoft Word не найден: {e}")
                return False
        
        if need_ppt:
            try:
                ppt = win32com.client.Dispatch("PowerPoint.Application")
                ppt.Quit()
                print("[OK] Microsoft PowerPoint найден")
            except Exception as e:
                print(f"[ERROR] Microsoft PowerPoint не найден: {e}")
                return False
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        
        return True
    
    except Exception as e:
        print(f"[ERROR] Ошибка проверки Office: {e}")
        return False

def main():
    """Основная логика"""
    print("=" * 60)
    if MODE == "text":
        print("  Извлечение текста из PPTX")
    else:
        print("  Конвертер документов через Microsoft Office")
    print("=" * 60)
    
    script_dir = Path(__file__).parent
    
    if MODE == "text":
        supported_formats = {".pptx", ".ppt"}
    else:
        supported_formats = {".pdf", ".docx", ".doc", ".pptx", ".ppt"}
    
    files = [f for f in script_dir.iterdir() 
             if f.is_file() and f.suffix.lower() in supported_formats]
    
    if not files:
        print("\n[WARNING] Не найдено файлов для обработки")
        if MODE == "text":
            print("Поддерживаемые форматы: PPTX, PPT")
        else:
            print("Поддерживаемые форматы: PDF, DOCX, DOC, PPTX, PPT")
        return
    
    print(f"\nНайдено файлов: {len(files)}")
    
    if MODE == "text":
        need_word, need_ppt = False, True
    else:
        need_word = any(f.suffix.lower() in (".docx", ".doc") for f in files)
        need_ppt = any(f.suffix.lower() in (".pptx", ".ppt") for f in files)
    
    if need_word or need_ppt:
        print("\nПроверка Microsoft Office...")
        if not check_office_installed(need_word, need_ppt):
            print("\n[WARNING] Microsoft Office не установлен или недоступен")
            print("Для работы скрипта требуется установленный Office")
            return
    
    try:
        pythoncom.CoInitialize()
    except pythoncom.com_error:
        pass
    
    word_app = None
    ppt_app = None
    
    try:
        if need_word:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
        
        if need_ppt:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Visible = 1
        
        if MODE == "text":
            all_text = []
            all_text.append(f"{'#'*80}")
            all_text.append(f"  ОБЪЕДИНЕННЫЙ ТЕКСТ ВСЕХ ПРЕЗЕНТАЦИЙ")
            all_text.append(f"  Дата создания: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            all_text.append(f"  Всего файлов: {len(files)}")
            all_text.append(f"{'#'*80}")
            
            for file in files:
                process_file(file, all_text, word_app, ppt_app)
            
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_path = OUTPUT_DIR / "all_presentations.txt"
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(all_text))
            
            print(f"\n  [DONE] Общий файл сохранен: {output_path}")
        else:
            for file in files:
                process_file(file, None, word_app, ppt_app)
    
    finally:
        if word_app:
            try:
                word_app.Quit()
            except:
                pass
        
        if ppt_app:
            try:
                ppt_app.Quit()
            except:
                pass
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    
    print("\n" + "=" * 60)
    if MODE == "text":
        print(f"[DONE] Общий текстовый файл сохранен в: {OUTPUT_DIR.resolve()}/all_presentations.txt")
    else:
        print(f"[DONE] Все изображения сохранены в: {OUTPUT_DIR.resolve()}")
    print("=" * 60)

if __name__ == "__main__":
    main()
