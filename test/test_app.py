import sys
import os

import re
from collections import Counter

# Добавляем корень проекта в PYTHONPATH
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx_replace import docx_replace

document1 = "Pushkin_AS.docx"
document2 = "replaced.docx"


# Создаем документ с помощью docx
doc = Document(document1)
# производим замену в документе указав пары ключ значение
docx_replace(doc, как="ккк", Лебедь="word3", гости="word3")
# производим замену в документе передав словарь
# my_dict = {"word5":"word6", "word7": "word8"}
# docx_replace(doc, **my_dict)
# сохраняем полученный документ
doc.save(document2)

def get_top_words(docx_path: str, top_n: int = 10) -> list:
    """Возвращает топ-N самых частых слов в документе."""
    doc = Document(docx_path)
    full_text = '\n'.join(para.text for para in doc.paragraphs)
    words = re.findall(r'\b\w+\b', full_text.lower())
    word_counts = Counter(words)
    return word_counts.most_common(top_n)

top1 = get_top_words(document1, top_n=10)

print("Топ-10 слов в документе 1:")
for word, count in top1:
    print(f"{word}: {count}")


def compare_docx_word_counts(docx_path1: str, docx_path2: str) -> dict:
    """
    Сравнивает два .docx документа по количеству вхождений слов.
    Возвращает словарь: {слово: разница_в_количестве},
    где разница = (частота в docx_path1) - (частота в docx_path2).
    """
    def extract_text(docx_path: str) -> str:
        doc = Document(docx_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)

    def get_word_counts(text: str) -> Counter:
        # Удаляем пунктуацию и приводим к нижнему регистру
        words = re.findall(r'\b\w+\b', text)
        return Counter(words)

    text1 = extract_text(docx_path1)
    text2 = extract_text(docx_path2)

    count1 = get_word_counts(text1)
    count2 = get_word_counts(text2)

    # Получаем все уникальные слова из обоих документов
    all_words = set(count1.keys()) | set(count2.keys())

    # Считаем разницу
    diff = {word: count1.get(word, 0) - count2.get(word, 0) for word in all_words}

    return diff

diff = compare_docx_word_counts(document1, document2)
for word, delta in sorted(diff.items(), key=lambda x: abs(x[1]), reverse=True):
    if delta != 0:
        print(f"{word}: {delta}")
