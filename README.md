# py-docx-comments

```import comments as cm

#optional pattern
pattern=['*',['text','text','number'],['CategoryText','InterpretationOfText','AuthenticityMarking']]

#single file
commentdata=cm.getcomments('docx/LoremIpsum - Copy.docx',pattern)
print(commentdata)

#directory
data=cm.getdirComments('docx/',pattern) #pattern is optional
# print(data)

#save to a json file
cm.jsonoutput(data,"testheader.fred",pattern,True)# add True as well to add the headers to the end of the file

#save to a csv file
cm.csvoutput(data,"comments.csv",pattern)#pattern is optional```
