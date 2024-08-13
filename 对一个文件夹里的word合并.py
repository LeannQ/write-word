# -*- coding: utf-8 -*-
"""
Created on Tue May 28 17:08:35 2024
@author: dbk
"""
import sys
sys.path.append('G:\ECPH_LY\MyPythonProject(python3)')
import Head
import docx
import os
import re
import win32com.client as wc
from docxcompose.composer import Composer

def get_head_digit(path):
    name = os.path.split(path)[1]
    pat = re.compile('^[\d]{1,5}')
    m = re.match(pat, name)
    if m:
        return int(m.group())
    return 1e10

def sort_func(lst):    
    r = sorted(lst, key=get_head_digit)
    return r
    
def merge_docx(outpath, paths):
    print('要合并的文档数量：', len(paths))
    master = docx.Document(paths[0])
    composer = Composer(master)
    for p in paths[1:]:
        try:
            doc = docx.Document(p)
            # doc.add_page_break()
            composer.append(doc)
        except:
            print(p)
    composer.save(outpath)

def get_path_suffix(p):
        _, suffix = os.path.splitext(p)
        return suffix

def doc2docx(doc_path, docx_path):
    word = wc.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(docx_path, 12)
    doc.Close()
    word.Quit()
    
def must_docx(p):
    suffix = get_path_suffix(p)
    if suffix == '.docx':
        return p
    elif suffix == '.doc':
        docxp = p + 'x'
        if os.path.exists(docxp):
            doc2docx(p, docxp)
        return docxp
    else:
        return None
        
if __name__ == "__main__":
    
    data_dir = 'G:/ECPH_LY/Data/协助同事/！三版内容中心/批量重命名文件/《版张上的艺术——邮政概览》2024.05.28'
    file_list = Head.GetFileList(data_dir, ['.docx'])
    file_list = [x.replace('\\','/') for x in file_list]
    docxpaths = list()
    metas = list()
    # 收集docx路径和文件信息
    for p in file_list:
        meta = dict()
        meta['file_path'] = p
        suffix = get_path_suffix(p)
        meta['suffix'] = suffix
        p = must_docx(p)
        if p and p not in docxpaths:
            docxpaths.append(p)
            meta['remark'] = '合并'
        else:
            meta['remark'] = ''
        metas.append(meta)
    # 合并
    docxpaths = sort_func(docxpaths)  # 这里发生了变化
    merge_docx(data_dir + '/combined.docx', docxpaths)
    
    
    