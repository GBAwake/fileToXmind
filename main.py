# -*- coding: utf-8 -*-
import os
import xmind
import zipfile
from xmind.core import workbook,saver
from xmind.core.topic import TopicElement


def add_node(w, father_node, name, is_dir=False):
    new_node = w.createTopic()
    try:
        name = name.decode('gbk').encode('utf-8')
    except:
        name = name.encode('utf-8')
    new_node.setTitle(name)
    if is_dir:
        new_node.setAttribute('style-id', '0n5nqvfpgenbpoq4304pjj4mku')
    else:
        new_node.setAttribute('style-id', '4847ntml8ncjh5980d9fptkc14')
    father_node.addSubTopic(new_node)
    return new_node


def list_all_file(w, root_dir, father_node):
    list = os.listdir(root_dir)
    for file in list:
        path = os.path.join(root_dir, file)
        if os.path.isdir(path):
            new_node = add_node(w, father_node, file, True)
            list_all_file(w, path, new_node)
        if os.path.isfile(path):
            new_node = add_node(w, father_node, file, False)

def xmind_add_style(style_file_path, xmind_path):
    xmind_file = zipfile.ZipFile(xmind_path, 'a')
    xmind_file.write(style_file_path, 'styles.xml')
    xmind_file.close()


if __name__ == '__main__':
    # root_dir = 'd:\\mytst\\test\\'
    root_dir = 'D:\\' u'文档'
    xmind_base_dir = '.\\base\\base.xmind'
    xmind_file_view = '.\\file_view\\test.xmind'
    w = xmind.load(xmind_base_dir)
    sheet = w.getPrimarySheet()
    sheet.setTitle('manage')
    root = sheet.getRootTopic()
    list_all_file(w=w, root_dir=root_dir, father_node=root)
    xmind.save(w, xmind_file_view)
    xmind_add_style('.\\base\\styles.xml', xmind_file_view)
