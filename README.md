# docx-replace

`docx-replace` - это простая утилита на Python для замены текста в документах `docx` (Microsoft Word) с сохранением форматирования.
Она имитирует поведение поиска и замены, подобное Microsoft Word, работая с `run`'ами параграфа.

Данная библиотека выполняет только 1 единственную функцию - замену текста в docx документе.
Вы даже можене не использовать данную библиотеку, а скопировать код с [GitHub](https://github.com/Misha1601/docx-replace), и вставить его в свой код.

## Отличие от других библиотек

- решение 1 проблемы.
- не использует сторонние библиотеки
- минимум кода и его простота

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
docx_replace(doc, **replacements) # Рекомендуется
# или
docx_replace(doc, word1='new_word1', word2='new_word2')

# Сохраняем измененный документ
doc.save("my_document_final.docx")
```

## Зависимости

- `python-docx`

## Лицензия

Этот проект распространяется под лицензией MIT. Подробности смотрите в файле `LICENSE`.