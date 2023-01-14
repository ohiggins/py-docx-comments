# MIT License

# Copyright (c) 2023 Oliver Higgins

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.


# Requires Python-docx
# pip install python-docx

# #Usage
# import comments as cm

# #optional pattern
# pattern=['*',
# ['text','text','number'],
# ['CategoryText','InterpretationOfText','AuthenticityMarking']
# ]
# To use the pattern above, indicate the charcters that will be used to break 
# the fields apart, in this case it is an *. Next is the pattern it expects to
# find, in this case 2 text fields and then a number field. All numbers are
# converted to a float (decimal). The next set of fields denotes what the 
# columns will be called when exported to csv file
# Needs to be read and added to the json file.

# #single file
# commentdata=cm.getcomments('docx/LoremIpsum - Copy.docx',pattern)
# print(commentdata)

# #directory
# data=cm.getdirComments('docx/',pattern) #pattern is optional
# print(data)

# #save to a json file
# cm.jsonoutput(data,"comments.json")

# #save to a csv file
# cm.csvoutput(data,"comments.csv",pattern)#pattern is optional

from docx import Document
from lxml import etree
import zipfile
import json
import csv
import os

ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_document_comments(docxFileName,option=[]):
    pattern=option
    comments_dict={}
    docxZip = zipfile.ZipFile(docxFileName)
    commentsXML = docxZip.read('word/comments.xml')
    et = etree.XML(commentsXML)
    comments = et.xpath('//w:comment',namespaces=ooXMLns)
    for c in comments:
        comment=c.xpath('string(.)',namespaces=ooXMLns)
        comment_id=c.xpath('@w:id',namespaces=ooXMLns)[0]
        comment_author=c.xpath('@w:author',namespaces=ooXMLns)[0]
        comment_date=c.xpath('@w:date',namespaces=ooXMLns)[0]
        comments_dict[comment_id]=[int(comment_id),comment,comment_author,comment_date,docxFileName]
        if len(pattern)>1:
            data = [x.strip() for x in comment.split(pattern[0])]
            if len(data)>1:
                for i,e in enumerate(data):
                    if pattern[1][i]=='text':
                        comments_dict[comment_id].append(data[i])
                    elif pattern[1][i]=='number':
                        comments_dict[comment_id].append(float(data[i]))
            else:
                for i,e in enumerate(pattern[1]):
                    if pattern[1][i]=='text':
                        comments_dict[comment_id].append("")
                    elif pattern[1][i]=='number':
                        comments_dict[comment_id].append(0)
    return comments_dict

def paragraph_comments(paragraph,comments_dict):
    comments=[]
    for run in paragraph.runs:
        comment_reference=run._r.xpath("./w:commentReference")
        if comment_reference:
            comment_id=comment_reference[0].xpath('@w:id',namespaces=ooXMLns)[0]
            comment=comments_dict[comment_id]
            comments.append(comment)
    return comments

def comments_with_reference_paragraph(docxFileName,option=[]):
    document = Document(docxFileName)
    comments_dict=get_document_comments(docxFileName,option)
    comments_with_their_reference_paragraph=[]
    for paragraph in document.paragraphs:  
        if comments_dict: 
            comments=paragraph_comments(paragraph,comments_dict)  
            if comments:
                comments_with_their_reference_paragraph.append({paragraph.text: comments})
    return comments_with_their_reference_paragraph

def getcommentscontent(doc):
    ooXMLns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    docxFilePath = doc
    docxZip = zipfile.ZipFile(docxFilePath)
    documentXML = docxZip.read('word/document.xml')
    et = etree.XML(documentXML)
    commentlist=[]
    flg=0
    current=[]
    for tag in et.iter():
        if not len(tag):
            if tag.tag=="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart":
                current.append(int(tag.attrib.items()[0][1]))
                commentlist.append([int(tag.attrib.items()[0][1]), ""])
                flg=1
            if tag.tag=="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd":
                current.remove(int(tag.attrib.items()[0][1]))
                if len(current)==0:
                    flg=0
            if tag.tag=="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t":
                if flg==1:
                    for item in current:
                        for i in range(len(commentlist)):
                            if commentlist[i][0] == item:
                                commentlist[i][1] = commentlist[i][1]+tag.text
    return commentlist

def getcomments(doc,option=[]):
    x1=comments_with_reference_paragraph(doc,option)
    z1=getcommentscontent(doc)
    for obj in z1:
        for obj1 in x1:
            for key, value in obj1.items():
                for item in value:
                    if item[0]==obj[0]:
                        obj.append(key)
                        for i1 in item:
                            obj.append(i1)
    z1= [[row[i] for i in range(len(row)) if i != 3] for row in z1]
    return z1

def csvoutput(data,filename,option=[]):
    column_names = ['comment_id', 'CommentText', 'Paragraph','Comment','Author','DateTime','FileName']
    if len(option)>1:
        for i,e in enumerate(option[1]):
            column_names.append(option[2][i])
    with open(filename, 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(column_names)  # adding column names
        writer.writerows(data)
    print(f'Data exported to {filename}.')

def getdirComments(dir,option=[]):
    list=getFileList(dir)
    masterlist=[]
    for file in list:
        masterlist=masterlist+getcomments(dir+file,option)
    return masterlist

def getFileList(dir):
    path = dir
    files = os.listdir(path)
    docx_files = [f for f in files if f.endswith('.docx')]
    return docx_files

def jsonoutput(data,filename,option=[]):
    json_data = json.dumps(data, indent=0, separators=(',', ':'), ensure_ascii=True, sort_keys=False, default=str)
    json_data = json_data.replace("\n", "").replace("\r", "")
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(json_data)
