import xlwings as xw
from docx import Document

if __name__ == '__main__':
    文件= Document('1.docx')
    print(文件.paragraphs[0])
    print(1)