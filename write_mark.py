#coding: utf-8

import zipfile
import os
import shutil

name_list = [
	['A',10,10,10,18,15,13,8,4],
	['B',10,10,10,18,19,12,10,5],
	['C',10,10,10,17,18,14,8,4],
	['D',10,10,10,18,15,13,10,4],
	['E',10,10,10,19,20,13,10,3],
	['F生',10,10,10,18,15,14,10,4],
	['G',10,10,10,19,15,13,10,5],
	['H',10,10,10,18,15,13,9,4],
	['I',10,10,10,18,15,13,9,4],
	['J',10,10,10,18,19,13,9,4],
	['K',9,10,10,17,15,12,8,4],
	['L',10,10,10,19,19,14,9,4],
	['M',10,10,10,17,18,13,9,3],
	['N',10,10,10,17,15,13,9,4],
	['O',10,10,10,18,15,13,9,4],
	['P',10,10,10,18,20,13,9,4],
	['Q',10,10,10,17,19,12,9,4],
]




for name in name_list:

    print('正在写入 '+ name[0] , end='  ')
    file_dir = 'your_dir' + name[0]

    # 确定内容文件
    f = open('your_dir\\document.xml','rb')
    text = f.readlines()[1].decode('utf-8')
    f.close()

    fullmark = 0
    for mark in name:
        if type(mark) == int:
            fullmark += mark
    print('总分'+str(fullmark), end=' ')

    content = text.replace('NAME@', str(name[0]))
    content = content.replace('1@', str(name[1]))
    content = content.replace('2@', str(name[2]))
    content = content.replace('3@', str(name[3]))
    content = content.replace('4@', str(name[4]))
    content = content.replace('5@', str(name[5]))
    content = content.replace('6@', str(name[6]))
    content = content.replace('7@', str(name[7]))
    content = content.replace('8@', str(name[8]))
    content = content.replace('LAST@', str(fullmark))
    content = content.replace('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>','')
    content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + content

    with open('your_dir\\document.xml', 'w', encoding='utf-8') as p:
        p.write(content)
        p.close()



    #复制同样文件并命名
    with open('your_moudle_word','rb') as mod:
        r = mod.read()
        with open(file_dir+'.docx','wb') as new:
            new.write(r)
            new.close()
        mod.close()


    #zipfile不支持直接删除，所以先解压
    z = zipfile.ZipFile(file_dir+'.docx', 'r')
    for file in z.namelist():
        z.extract(file,file_dir)

    z.close()

    #替换document.xml文件
    os.remove(file_dir+'\\word\\document.xml')
    shutil.copyfile('your_dir\\document.xml',file_dir+str('\\word\\document.xml'))

    #重新压缩打包
    zp = zipfile.ZipFile(file_dir+'.docx','w')
    for dirpath, dirnames, filenames in os.walk(file_dir):
        for filename in filenames:
            path = os.path.join(dirpath,filename)
            abs_path = file_dir + '\\'
            zp.write(path,path.strip(abs_path))

    #清除垃圾
    os.remove('your_dir\\document.xml')
    shutil._rmtree_unsafe(file_dir,onerror=None)

    print('写入成功！')

input('\n全部写入完成，请按任意键退出')

