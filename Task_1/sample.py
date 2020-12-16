import os

os.mkdir('folder_1')
os.chdir('folder_1')
with open('something','w') as f:
    f.write('something!\n')
    
'''
中文这个我还没有研究好heihei 
'''
#with open('something','a',encoding='utf-8') as f:
#    f.write('中文')

