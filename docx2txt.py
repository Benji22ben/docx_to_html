#! /usr/bin/env python

import argparse
from itertools import count
import re
import xml.etree.ElementTree as ET
from xmlrpc.client import Boolean
import zipfile
import os
import sys


nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def process_args():
    parser = argparse.ArgumentParser(description='A pure python-based utility '
                                                 'to extract text and images '
                                                 'from docx files.')
    parser.add_argument("docx", help="path of the docx file")
    parser.add_argument('-i', '--img_dir', help='path of directory '
                                                'to extract images')

    args = parser.parse_args()

    if not os.path.exists(args.docx):
        print('File {} does not exist.'.format(args.docx))
        sys.exit(1)

    if args.img_dir is not None:
        if not os.path.exists(args.img_dir):
            try:
                os.makedirs(args.img_dir)
            except OSError:
                print("Unable to create img_dir {}".format(args.img_dir))
                sys.exit(1)
    return args


def qn(tag):
    """
    Stands for 'qualified name', a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    Source: https://github.com/python-openxml/python-docx/
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)


def xml2text(xml):
    """
    A string representing the textual content of this run, with content
    child elements like ``<w:tab/>`` translated to their Python
    equivalent.
    Adapted from: https://github.com/python-openxml/python-docx/
    """
    text = u''
    root = ET.fromstring(xml)
    before = 4
    sz_moy = 0
    bold = 0
    i = 0
    attrib = ""
    # b = ""
    for child in root.iter():
        if child.tag == qn('w:sz'):
            sz_moy += int(child.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"))
            i += 1
    sz_moy = sz_moy / i
    for child in root.iter():
        # Trying to see how is made a file with all the child.tag
        # b += "\n" + child.tag
        # with open("outputdocxtagtry3.txt", "w", encoding="utf-8") as f:
        #     f.write(b)

        if child.tag == qn('w:b') and bold == 0:
            bold += 1
            print(bold)
        elif child.tag == qn('w:t'):
            t_text = child.text
            if bold == 1:
                text += "\n" + "<b>" + t_text + "</b>" if t_text is not None else ''
            else:
                text += "\n" + t_text if t_text is not None else ''
            before = 0
            bold = 0
        elif child.tag in (qn('w:br'), qn('w:cr'), qn("w:p"), qn('w:tab')) and before < 2:
            text += '\n<br>'
            before += 1
    return text


def process(docx, img_dir=None):
    text = u''

    # unzip the docx in memory
    zipf = zipfile.ZipFile(docx)
    filelist = zipf.namelist()

    # get header text
    # there can be 3 header files in the zip
    
        # header_xmls = 'word/header[0-9]*.xml'
        # for fname in filelist:
        #     if re.match(header_xmls, fname):
        #         text += xml2text(zipf.read(fname))

    # get main text
    doc_xml = 'word/document.xml'
    text += xml2text(zipf.read(doc_xml))

    # # get footer text
    # # there can be 3 footer files in the zip
    
        # footer_xmls = 'word/footer[0-9]*.xml'
        # for fname in filelist:
        #     if re.match(footer_xmls, fname):
        #         text += xml2text(zipf.read(fname))

        # if img_dir is not None:
        #     # extract images
        #     for fname in filelist:
        #         _, extension = os.path.splitext(fname)
        #         if extension in [".jpg", ".jpeg", ".png", ".bmp"]:
        #             dst_fname = os.path.join(img_dir, os.path.basename(fname))
        #             with open(dst_fname, "wb") as dst_f:
        #                 dst_f.write(zipf.read(fname))

    zipf.close()
    return text.strip()


if __name__ == '__main__':
    args = process_args()
    text = process(args.docx, args.img_dir)
    sys.stdout.write(text.encode('utf-8'))
