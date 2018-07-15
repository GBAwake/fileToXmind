# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os
import xmind
import zipfile
from xmind.core import workbook,saver
from xmind.core.topic import TopicElement


# def add_node(w, father_node, name, is_dir=False):
#     new_node = w.createTopic()
#     try:
#         name = name.decode('gbk').encode('utf-8')
#     except:
#         name = name.encode('utf-8')
#     new_node.setTitle(name)
#     if is_dir:
#         new_node.setAttribute('style-id', '0n5nqvfpgenbpoq4304pjj4mku')
#         new_node.setFileHyperlink()
#     else:
#         new_node.setAttribute('style-id', '4847ntml8ncjh5980d9fptkc14')
#     father_node.addSubTopic(new_node)
#     return new_node


# def list_all_file(w, root_dir, father_node):
#     list = os.listdir(root_dir)
#     for file in list:
#         path = os.path.join(root_dir, file)
#         if os.path.isdir(path):
#             new_node = add_node(w, father_node, file, True)
#             list_all_file(w, path, new_node)
#         if os.path.isfile(path):
#             new_node = add_node(w, father_node, file, False)

def add_file_to_xmind(xmind_info, file_info):
    list = os.listdir(file_info.now_dir)
    for file in list:
        path = os.path.join(file_info.now_dir, file)
        f = FileInfo(
            root_dir=file_info.root_dir,
            now_dir=path,
            relative_path=file_info.relative_path + '/' + file,
            name=file,
        )
        if os.path.isdir(path):
            f.is_dir = True
            new_node = add_node(xmind_info, f)
            x = XMindInfo(xmind_info.xmind_file, new_node)
            add_file_to_xmind(x, f)
        if os.path.isfile(path):
            f.is_dir = False
            new_node = add_node(xmind_info, f)

def add_node(xmind_info, file_info):
    new_node = deal_special_file(xmind_info, file_info)
    if new_node is not None:
        return new_node
    new_node = xmind_info.xmind_file.createTopic()
    name = encode_name(file_info.name)
    new_node.setTitle(name)
    if file_info.is_dir:
        new_node.setAttribute('style-id', '0n5nqvfpgenbpoq4304pjj4mku')
    else:
        new_node.setAttribute('style-id', '4847ntml8ncjh5980d9fptkc14')
    #  增加超链接
    new_node.setFileHyperlink(encode_name('file://' + file_info.relative_path))
    xmind_info.father_node.addSubTopic(new_node)
    return new_node

def encode_name(name):
    name = name.decode('gbk')
    return name


def deal_special_file(xmind_info, file_info):
    new_node = deal_special_file_ignore(xmind_info, file_info)
    if new_node is not None:
        return new_node

    special_list = [deal_special_file_url]
    for special_deal in special_list:
        new_node = special_deal(xmind_info, file_info)
        if new_node is not None:
            return new_node
    return None


def deal_special_file_url(xmind_info, file_info):
    if file_info.name.find('url_') != 0:
        return None
    new_node = xmind_info.xmind_file.createTopic()
    name = file_info.name[4:file_info.name.find('.')]
    name = encode_name(name)
    new_node.setTitle(name)
    new_node.setAttribute('style-id', '4847ntml8ncjh5980d9fptkc14')
    #  增加超链接
    f = open(file_info.now_dir)
    url_name = f.readline()
    new_node.setURLHyperlink(encode_name(url_name))
    xmind_info.father_node.addSubTopic(new_node)
    return new_node

def deal_special_file_ignore(xmind_info, file_info):
    ignore_list = ['Thumbs.db', '~$']
    for ignore in ignore_list:
        if file_info.name.find(ignore) != -1:
            return xmind_info.father_node

def xmind_add_style(style_file_path, xmind_path):
    xmind_file = zipfile.ZipFile(xmind_path, 'a')
    xmind_file.write(style_file_path, 'styles.xml')
    xmind_file.close()


class FileInfo:
    def __init__(self, root_dir=None, now_dir=None, relative_path=None, name=None, is_dir=False):
        self.root_dir = root_dir
        self.now_dir = now_dir
        self.relative_path = relative_path
        self.name = name
        self.is_dir = is_dir


class XMindInfo:
    def __init__(self, xmind_file=None, father_node=None):
        self.xmind_file = xmind_file
        self.father_node = father_node


if __name__ == '__main__':
    root_dir = 'd:\\mytst\\test\\'
    # root_dir = 'D:\\' u'文档\\知识'
    xmind_base_dir = '.\\base\\base.xmind'
    xmind_file_view = '.\\file_view\\test.xmind'
    w = xmind.load(xmind_base_dir)
    sheet = w.getPrimarySheet()
    sheet.setTitle('manage')
    root = sheet.getRootTopic()
    x = XMindInfo(w, root)
    f = FileInfo(root_dir=root_dir, now_dir=root_dir, relative_path='.')

    add_file_to_xmind(x, f)
    # list_all_file(w=w, root_dir=root_dir, father_node=root)
    xmind.save(w, xmind_file_view)
    xmind_add_style('.\\base\\styles.xml', xmind_file_view)
