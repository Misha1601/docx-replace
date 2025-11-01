# docx-replace

`docx-replace` - это простая утилита на Python для замены текста в документах `docx` (Microsoft Word) с сохранением форматирования.
Она имитирует поведение поиска и замены, подобное Microsoft Word, работая с `run`'ами параграфа.

## Установка

```bash
pip install python-docx
pip install docx-replace
```

## Использование

Пример использования:

```python
from docx import Document
from docx_replace import docx_replace

# Загружаем документ
doc = Document("my_template.docx")

# Определяем значения для замены
replacements = {
    "{{имя_клиента}}": "Иван Иванов",
    "{{дата}}": "15.03.2024",
    "{{город}}": "Москва"
}

# Выполняем замену
docx_replace(doc, **replacements)

# Сохраняем измененный документ
doc.save("my_document_final.docx")
```

## Зависимости

- `python-docx`

## Лицензия

Этот проект распространяется под лицензией MIT. Подробности смотрите в файле `LICENSE` (если применимо).