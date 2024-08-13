# -*- coding: utf-8 -*-
"""
Created on Mon Sep 02 08:58:20 2019

@author: SJB-083
"""

#按作者拆分条目
import Head
import re

def Process_author_list(author_list):
    
    new_list = []
    for author in author_list:
        new_list.append(Process_2author(author))
    return new_list 


def Process_2author(author):
    return author.split(' ')[0]

if __name__ == "__main__":
    
     data_dir = u'G:/ECPH_LY/Data/协助同事/王瑜/唐诗PDF转WORD（20210823）'
     xml_file = data_dir + '/' + u'全部xml.xml'
     xml_data = Head.ReadFile(xml_file)
     entry_f = r'<entry>[\s\S]*?</entry>'
     author_f = r'<AUTHOR>(.*?)</AUTHOR>'
     itemcn_f = r'<itemcn>(.*?)</itemcn>'
     entry_r = re.compile(entry_f)
     author_r = re.compile(author_f)
     itemcn_r = re.compile(itemcn_f)
     #得到不重复的作者的取值集合
     author_list_xml = re.findall(author_r,xml_data)
     author_list_processed = Process_author_list(author_list_xml)
     author_set = list(set(author_list_processed))
     author_set.append('')
     #得到xml中作者与entry一一对应的列表
     entry_list = re.findall(entry_r,xml_data)
     author_list_perEntry = []
     for entry in entry_list:
         author_list = re.findall(author_r,entry)
         itemcn = re.findall(itemcn_r,entry)
         itemcn = itemcn[0]
         author = re.findall(author_r,entry)
         if len(author) == 0:
             author_list_perEntry.append('')
         else:
             author_list_perEntry.append(Process_2author(author[0]))
             
    #遍历不同的作者，把属于作者的词条放入该作者的口袋中
     for author in author_set:         
         xml_per_author = ''
         for i in range(len(entry_list)):         
             entry = entry_list[i]
             author_entry = author_list_perEntry[i]
             if author_entry == author:
                xml_per_author = xml_per_author + '\n' + entry
         if author != '':
             file_dst = open(data_dir+'/' + author + '.xml','w',encoding='utf-8')  
         else:
             file_dst = open(data_dir+u'/无作者.xml','w',encoding='utf-8')  
         file_dst.write(xml_per_author)
         file_dst.close()   
         
        
    
    
    
    
    
    
    