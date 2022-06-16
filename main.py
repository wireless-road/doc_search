import os
import pathlib
import re
import PyPDF2
import warnings

import docx.opc.exceptions
import pywintypes
from PyPDF2.errors import PdfReadError

warnings.filterwarnings("ignore")

import win32com.client
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table, _Row
from docx.text.paragraph import Paragraph

import logging
logging.basicConfig(filename='errors.log', level=logging.DEBUG,
                    format='%(asctime)s %(message)s')
logger=logging.getLogger(__name__)

pattern = "substring_to_find"
directory = pathlib.Path(r"C:\Users\User\root_folder_for_recursive_search")


def scan_pdf(directory, filename):
    filepath = os.path.join(directory, filename)
    try:
        object = PyPDF2.PdfFileReader(filepath)
    except FileNotFoundError:
        logger.error(f"! can't open PDF file {filepath}")
        return
    except PdfReadError as e:
        logger.error(f"! error reading PDF file: {filepath} - {e}")
        return

    NumPages = object.getNumPages()

    is_filename_not_printed = True
    for i in range(0, NumPages):
        PageObj = object.getPage(i)
        try:
            Text = PageObj.extractText()
        except AssertionError:
            continue
        except UnicodeDecodeError:
            logger.error(f"! can't read PDF {filepath} page {i+1}")
            continue
        res = re.search(pattern, Text)
        if res is not None and is_filename_not_printed:
            print(f'PDF {filepath}')
            is_filename_not_printed = False
        if res is not None:
            print(f' \t\tpage {i}. Match  {res}')

def iter_block_items(parent):
    """
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif isinstance(parent, _Row):
        parent_elm = parent._tr
    else:
        raise ValueError("something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def scan_docx(directory, filename):
    filepath = os.path.join(directory, filename)
    try:
        document = Document(filepath)
    except docx.opc.exceptions.PackageNotFoundError:
        logger.error(f"! can't open file {filepath}")
        return
    is_filename_not_printed = True
    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            if pattern in block.text:
                if is_filename_not_printed:
                    print(f'DOCX \t{filepath}. Match {pattern}')
                    is_filename_not_printed = False
                    return
        elif isinstance(block, Table):
            for row in block.rows:
                row_data = []
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        row_data.append(paragraph.text)
                        if pattern in paragraph.text:
                            if is_filename_not_printed:
                                print(f'DOCX \t{filepath}. Match {pattern}')
                                is_filename_not_printed = False
                                return

def scan_doc(directory, filename):
    filepath = os.path.join(directory, filename)

    root_folder = directory
    doc_ex = root_folder + "/" + filename
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False
    try:
        wb = word.Documents.Open(doc_ex)
    except pywintypes.com_error:
        logger.error(f'! DOC file not found {doc_ex}')
        return
    doc = word.ActiveDocument
    docText = doc.Content.Text

    if pattern in docText:
        print(f'DOC \t{filepath}. Match {pattern}')
        wb.Close()
        word.Quit()
        return
    wb.Close()
    word.Quit()

cnt = 0
def print_dot():
    global cnt
    if cnt < 100:
        print('.', end='')
    else:
        print('.')
        cnt = 0

for root, dir, files in os.walk(directory):
    for file in files:
        filename = os.fsdecode(file)
        if filename.startswith('~$'):
            continue
        if filename.endswith(".pdf"):
            cnt += 1
            print_dot()
            scan_pdf(root, filename)
        elif filename.endswith(".docx"):
            cnt += 1
            print_dot()
            scan_docx(root, filename)
        elif filename.endswith(".doc"):
            cnt += 1
            print_dot()
            scan_doc(root, filename)
        else:
            continue
