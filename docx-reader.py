from docx2txt import process

# extract text to html

url = "./REMUNERATION.docx"

with open("output.html", "w", encoding="utf-8") as f:
    f.write(process(url))

