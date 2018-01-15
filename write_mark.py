#coding: utf-8

import zipfile
import os
import shutil

name_list = [
	['林凯妍',10,10,10,18,15,13,8,4],
	['黄恩童',10,10,10,18,19,12,10,5],
	['朱彦科',10,10,10,17,18,14,8,4],
	['李星乐',10,10,10,18,15,13,10,4],
	['黎  悦',10,10,10,19,20,13,10,3],
	['陈伟生',10,10,10,18,15,14,10,4],
	['刘鸿杰',10,10,10,19,15,13,10,5],
	['吴  磊',10,10,10,18,15,13,9,4],
	['张仕萌',10,10,10,18,15,13,9,4],
	['陈达朗',10,10,10,18,19,13,9,4],
	['林晓晴',9,10,10,17,15,12,8,4],
	['邹  丽',10,10,10,19,19,14,9,4],
	['许泽钿',10,10,10,17,18,13,9,3],
	['杨锦鹏',10,10,10,17,15,13,9,4],
	['陈晓琪',10,10,10,18,15,13,9,4],
	['黄  信',10,10,10,18,20,13,9,4],
	['邓雨晴',10,10,10,17,19,12,9,4],
]




for name in name_list:

    print('正在写入 '+ name[0] , end='  ')
    file_dir = 'E:\\SAU-MSC\\2017-18人员考核\\12月\\秘书处\\' + name[0]

    # 确定内容文件
    f = open('E:\\SAU-MSC\\2017-18人员考核\\document.xml','rb')
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

    with open('E:\\SAU-MSC\\2017-18人员考核\\12月\\秘书处\\document.xml', 'w', encoding='utf-8') as p:
        p.write(content)
        p.close()



    #复制同样文件并命名
    with open('E:\\SAU-MSC\\2017-18人员考核\\广东工业大学校社联人员考核制度干事考核表.docx','rb') as mod:
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
    shutil.copyfile('E:\\SAU-MSC\\2017-18人员考核\\12月\\秘书处\\document.xml',file_dir+str('\\word\\document.xml'))

    #重新压缩打包
    zp = zipfile.ZipFile(file_dir+'.docx','w')
    for dirpath, dirnames, filenames in os.walk(file_dir):
        for filename in filenames:
            path = os.path.join(dirpath,filename)
            abs_path = file_dir + '\\'
            zp.write(path,path.strip(abs_path))

    #清除垃圾
    os.remove('E:\\SAU-MSC\\2017-18人员考核\\12月\\秘书处\\document.xml')
    shutil._rmtree_unsafe(file_dir,onerror=None)

    print('写入成功！')

input('\n全部写入完成，请按任意键退出')

