from pdf2docx import Converter
import PySimpleGUI as sg
import datetime 



font_name = '微软雅黑'

key_transform = '转换'
key_about = '关于'

now_year = datetime.datetime.now().year
title = 'AST For PDF2Word ©2016-' + str(now_year)

def pdf2word(file_path):
    file_name = file_path.split('.')[0]
    docx_file = file_name + '.docx'

    p2w = Converter(file_path)
    p2w.convert(docx_file,start =0,end=None)
    p2w.close()
    return docx_file

def make_window():
    sg.theme('LightBlue')
    layout = [
        [
            sg.Text('[选择]PDF[转换]为Word',font=(font_name,10),background_color='white'),
            sg.Text('',key = 'filename',size=(50,1),font=(font_name,10),text_color='red',background_color='white')
        ],
        [
            sg.Text('',key = 'finis',size=(60,1),font=(font_name,10),text_color='red',background_color='white')
            #sg.Output(size=(80,10),font=(font_name,10))
        ],
        [
            sg.FileBrowse('选择',key='file',target='filename',size=(13,1)),sg.Button(key_transform,size=(13,1)),sg.Button(key_about,size=(13,1))
        ]
    ]

    return sg.Window(title, layout,font=(font_name,16),icon=None,grab_anywhere=True, keep_on_top=True,background_color='white',element_justification='center')


def main():

    window = make_window()

    while 1:
        event, values = window.read()
        if event == None:
            break
        if event == key_about:
            content = 'This tool was designed and developed by AmyliaScarlet.This software belongs to the AST series. If you have any questions in use, please contact me if you can.'
            sg.popup_ok(content,no_titlebar = True,keep_on_top=True)
        if event == key_transform:
            if values['file'] and values['file'].split('.')[1] == 'pdf':
                file_path = pdf2word(values['file'])
                window['finis'].update('成功! 保存路径:'+file_path)
            else:
                window['finis'].update('请选择PDF文件!')
                
    window.close()


main()