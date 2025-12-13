#!/usr/bin/env python3
"""
Конвертер документов в изображения через Microsoft Office (Windows)
Использует реальный Word и PowerPoint для 100% точности
Поддерживает: PDF, DOCX, PPTX

Требования: установленный Microsoft Office
"""

import fitz  # PyMuPDF для PDF
from pathlib import Path
import win32com.client
import pythoncom
import time

# ============ НАСТРОЙКИ ============
DPI = 200  # Качество изображений (150-300)
IMG_FORMAT = "png"  # png или jpg
OUTPUT_DIR = Path("screenshots")
MODE = "images"  # "images" - конвертация в изображения, "text" - извлечение текста из PPTX

# Константы для Word
WD_EXPORT_FORMAT_PDF = 17  # wdExportFormatPDF

# ============ УТИЛИТЫ ============
def get_image_name(page_num: int, prefix: str = "page") -> str:
    """Возвращает порядковое имя файла"""
    return f"{prefix}_{page_num:03d}.{IMG_FORMAT}"

# ============ PDF → PNG (через PyMuPDF) ============
def convert_pdf(pdf_path: Path, output_dir: Path) -> None:
    """Конвертирует PDF в изображения"""
    if not pdf_path.exists():
        print(f"  ⚠ Файл не найден: {pdf_path}")
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
            print(f"    ✓ {img_path.name}")
            
            # Освобождаем память
            pix = None
    finally:
        if doc:
            doc.close()

# ============ DOCX → PNG (через Word) ============
def convert_docx(docx_path: Path, output_dir: Path, word_app) -> None:
    """Конвертирует DOCX через Microsoft Word"""
    if not docx_path.exists():
        print(f"  ⚠ Файл не найден: {docx_path}")
        return
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    doc = None
    pdf_doc = None
    temp_pdf = output_dir / "temp_export.pdf"
    
    try:
        # Открываем документ
        doc_path = str(docx_path.resolve())
        doc = word_app.Documents.Open(doc_path, ReadOnly=True)
        
        page_count = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
        print(f"  Страниц: {page_count}")
        
        # Сохраняем весь документ как PDF
        doc.ExportAsFixedFormat(
            OutputFileName=str(temp_pdf.resolve()),
            ExportFormat=WD_EXPORT_FORMAT_PDF
        )
        
        # Конвертируем PDF в изображения
        pdf_doc = fitz.open(str(temp_pdf))
        scale = DPI / 72.0
        matrix = fitz.Matrix(scale, scale)
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            
            img_path = output_dir / get_image_name(page_num + 1, "page")
            pix.save(str(img_path))
            print(f"    ✓ {img_path.name}")
            
            # Освобождаем память
            pix = None
        
    except Exception as e:
        print(f"  ❌ Ошибка Word: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Закрываем PDF документ
        if pdf_doc:
            pdf_doc.close()
        
        # Удаляем временный PDF
        if temp_pdf.exists():
            try:
                temp_pdf.unlink()
            except:
                pass
        
        # Закрываем Word документ
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass

# ============ PPTX → TXT (извлечение текста) ============
def extract_pptx_text(pptx_path: Path, all_text_list: list, ppt_app) -> None:
    """Извлекает весь текст из PPTX и добавляет в общий список"""
    if not pptx_path.exists():
        print(f"  ⚠ Файл не найден: {pptx_path}")
        return
    
    presentation = None
    
    try:
        # Открываем презентацию
        prs_path = str(pptx_path.resolve())
        presentation = ppt_app.Presentations.Open(prs_path, ReadOnly=True, WithWindow=False)
        
        slide_count = presentation.Slides.Count
        print(f"  Слайдов: {slide_count}")
        
        # Добавляем заголовок презентации
        all_text_list.append(f"\n\n{'='*80}")
        all_text_list.append(f"ПРЕЗЕНТАЦИЯ: {pptx_path.name}")
        all_text_list.append(f"Всего слайдов: {slide_count}")
        all_text_list.append(f"{'='*80}\n")
        
        for slide_num in range(1, slide_count + 1):
            slide = presentation.Slides(slide_num)
            
            all_text_list.append(f"\n{'─'*60}")
            all_text_list.append(f"СЛАЙД {slide_num}")
            all_text_list.append(f"{'─'*60}\n")
            
            # Извлекаем текст из всех фигур на слайде
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
            
            print(f"    ✓ Слайд {slide_num} обработан")
        
    except Exception as e:
        print(f"  ❌ Ошибка PowerPoint: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Закрываем презентацию
        if presentation:
            try:
                presentation.Close()
            except:
                pass

# ============ PPTX → PNG (через PowerPoint) ============
def convert_pptx(pptx_path: Path, output_dir: Path, ppt_app) -> None:
    """Конвертирует PPTX через Microsoft PowerPoint"""
    if not pptx_path.exists():
        print(f"  ⚠ Файл не найден: {pptx_path}")
        return
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    presentation = None
    
    try:
        # Открываем презентацию
        prs_path = str(pptx_path.resolve())
        presentation = ppt_app.Presentations.Open(prs_path, ReadOnly=True, WithWindow=False)
        
        slide_count = presentation.Slides.Count
        print(f"  Слайдов: {slide_count}")
        
        # Экспортируем каждый слайд
        for slide_num in range(1, slide_count + 1):
            slide = presentation.Slides(slide_num)
            
            img_path = output_dir / get_image_name(slide_num, "slide")
            
            try:
                # Экспорт слайда как изображение
                width = int(presentation.PageSetup.SlideWidth * DPI / 72)
                height = int(presentation.PageSetup.SlideHeight * DPI / 72)
                
                slide.Export(
                    str(img_path.resolve()),
                    "PNG" if IMG_FORMAT == "png" else "JPG",
                    width,
                    height
                )
                
                print(f"    ✓ {img_path.name}")
            
            except Exception as e:
                print(f"    ⚠ Ошибка экспорта слайда {slide_num}: {e}")
        
    except Exception as e:
        print(f"  ❌ Ошибка PowerPoint: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Закрываем презентацию
        if presentation:
            try:
                presentation.Close()
            except:
                pass

# ============ ГЛАВНАЯ ФУНКЦИЯ ============
def process_file(file_path: Path, all_text_list: list = None, word_app=None, ppt_app=None) -> None:
    """Обрабатывает один файл"""
    if not file_path.exists():
        print(f"  ⚠ Файл не найден: {file_path}")
        return
    
    ext = file_path.suffix.lower()
    output_dir = OUTPUT_DIR / file_path.stem
    
    print(f"\n📄 {file_path.name}")
    
    try:
        # Режим извлечения текста (только для PPTX)
        if MODE == "text":
            if ext in (".pptx", ".ppt"):
                if ppt_app:
                    extract_pptx_text(file_path, all_text_list, ppt_app)
                else:
                    print("  ❌ PowerPoint не доступен")
            else:
                print(f"  ⚠ Режим 'text' работает только с PPTX файлами")
            return
        
        # Режим конвертации в изображения
        if ext == ".pdf":
            convert_pdf(file_path, output_dir)
        elif ext in (".docx", ".doc"):
            if word_app:
                convert_docx(file_path, output_dir, word_app)
            else:
                print("  ❌ Word не доступен")
        elif ext in (".pptx", ".ppt"):
            if ppt_app:
                convert_pptx(file_path, output_dir, ppt_app)
            else:
                print("  ❌ PowerPoint не доступен")
        else:
            print(f"  ⚠ Формат {ext} не поддерживается")
            return
        
        print(f"  ✅ Готово: {output_dir}")
    
    except Exception as e:
        print(f"  ❌ Ошибка: {e}")
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
                print("✓ Microsoft Word найден")
            except Exception as e:
                print(f"✗ Microsoft Word не найден: {e}")
                return False
        
        if need_ppt:
            try:
                ppt = win32com.client.Dispatch("PowerPoint.Application")
                ppt.Quit()
                print("✓ Microsoft PowerPoint найден")
            except Exception as e:
                print(f"✗ Microsoft PowerPoint не найден: {e}")
                return False
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        
        return True
    
    except Exception as e:
        print(f"✗ Ошибка проверки Office: {e}")
        return False

def main():
    """Основная логика"""
    print("=" * 60)
    if MODE == "text":
        print("  Извлечение текста из PPTX")
    else:
        print("  Конвертер документов через Microsoft Office")
    print("=" * 60)
    
    # Ищем файлы в текущей папке
    script_dir = Path(__file__).parent
    
    if MODE == "text":
        supported_formats = {".pptx", ".ppt"}
    else:
        supported_formats = {".pdf", ".docx", ".doc", ".pptx", ".ppt"}
    
    files = [f for f in script_dir.iterdir() 
             if f.is_file() and f.suffix.lower() in supported_formats]
    
    if not files:
        print("\n⚠ Не найдено файлов для обработки")
        if MODE == "text":
            print("Поддерживаемые форматы: PPTX, PPT")
        else:
            print("Поддерживаемые форматы: PDF, DOCX, DOC, PPTX, PPT")
        return
    
    print(f"\nНайдено файлов: {len(files)}")
    
    # Определяем нужные компоненты Office
    if MODE == "text":
        need_word, need_ppt = False, True
    else:
        need_word = any(f.suffix.lower() in (".docx", ".doc") for f in files)
        need_ppt = any(f.suffix.lower() in (".pptx", ".ppt") for f in files)
    
    # Проверяем наличие Office
    if need_word or need_ppt:
        print("\nПроверка Microsoft Office...")
        if not check_office_installed(need_word, need_ppt):
            print("\n⚠ Microsoft Office не установлен или недоступен")
            print("Для работы скрипта требуется установленный Office")
            return
    
    # Инициализация COM и запуск Office приложений
    try:
        pythoncom.CoInitialize()
    except pythoncom.com_error:
        pass
    
    word_app = None
    ppt_app = None
    
    try:
        # Запускаем нужные приложения один раз
        if need_word:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
        
        if need_ppt:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Visible = 1
        
        # Для режима text создаем общий список текста
        if MODE == "text":
            all_text = []
            all_text.append(f"{'#'*80}")
            all_text.append(f"  ОБЪЕДИНЕННЫЙ ТЕКСТ ВСЕХ ПРЕЗЕНТАЦИЙ")
            all_text.append(f"  Дата создания: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            all_text.append(f"  Всего файлов: {len(files)}")
            all_text.append(f"{'#'*80}")
            
            # Обрабатываем каждый файл
            for file in files:
                process_file(file, all_text, word_app, ppt_app)
            
            # Сохраняем общий файл
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_path = OUTPUT_DIR / "all_presentations.txt"
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(all_text))
            
            print(f"\n  ✅ Общий файл сохранен: {output_path}")
        else:
            # Обрабатываем каждый файл
            for file in files:
                process_file(file, None, word_app, ppt_app)
    
    finally:
        # Закрываем приложения
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
        
        # Деинициализация COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    
    print("\n" + "=" * 60)
    if MODE == "text":
        print(f"✅ Общий текстовый файл сохранен в: {OUTPUT_DIR.resolve()}/all_presentations.txt")
    else:
        print(f"✅ Все изображения сохранены в: {OUTPUT_DIR.resolve()}")
    print("=" * 60)

if __name__ == "__main__":
    main()
