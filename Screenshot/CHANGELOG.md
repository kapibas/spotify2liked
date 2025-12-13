# Changelog

Все значимые изменения в этом проекте будут документироваться в этом файле.

Формат основан на [Keep a Changelog](https://keepachangelog.com/ru/1.0.0/),
и этот проект придерживается [Semantic Versioning](https://semver.org/lang/ru/).

## [1.0.0] - 2024-XX-XX

### Добавлено
- Поддержка конвертации PDF в PNG/JPG через PyMuPDF
- Поддержка конвертации DOCX/DOC в PNG/JPG через Microsoft Word
- Поддержка конвертации PPTX/PPT в PNG/JPG через Microsoft PowerPoint
- Режим извлечения текста из презентаций (PPTX/PPT → TXT)
- Настраиваемое качество изображений (DPI: 150-300)
- Выбор формата выходных изображений (PNG или JPG)
- Автоматическое управление памятью
- Оптимизированная производительность (Office приложения запускаются один раз)
- Подробная документация в README.md
- Шаблоны для Issues и Pull Requests
- Лицензия MIT

### Технические детали
- Использование COM интерфейса для работы с Microsoft Office
- Поддержка Python 3.7+
- Требования: Windows, Microsoft Office
