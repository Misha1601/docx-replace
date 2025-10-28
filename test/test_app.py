import sys
import os

# Добавляем корень проекта в PYTHONPATH
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx_replace import docx_replace



# Создаем документ с помощью docx
doc = Document("Pushkin_AS2.docx")
# производим замену в документе указав пары ключ значение
docx_replace(doc, лебедь="word2", Лебедь="word3")
# производим замену в документе передав словарь
# my_dict = {"word5":"word6", "word7": "word8"}
# docx_replace(doc, **my_dict)
# сохраняем полученный документ
doc.save("replaced.docx")