import os
import shutil
import time

import pandas as pd
import psutil
import win32api
import win32com.client as win32
import win32con
import xlwings as xw
from PIL import ImageGrab
from win32com.client.gencache import EnsureDispatch


def dir_maker(path, file):
    files = os.listdir(path)  # 获取path路径下的所有文件
    if file in files:
        pass  # 有则pass
    else:
        os.mkdir(path + "\\" + file)  # 无则在此路径下生成file文件夹


def get_picture(district):
    wb = ex.Workbooks.Open(path_excel_output + '%s - trading pending order to 销售 %s.xlsx' % (district, email_day))
    num = ''
    mumber = 0
    for i in wb.Worksheets:
        for n, shape in enumerate(i.Shapes):
            shape.Copy()
            image = ImageGrab.grabclipboard()
            if image.size[0]>=200 and image.size[1]>=200:
            #if (int(image.size[0])-int(image.size[1]))<=300:
            #if (2000, 2000) >= image.size >= (100, 100): #此size指的是图片的分辨率尺寸
                print(image.size[0])
                if image.size[0]>=1000:
                    num='大'
                    try:
                        image.convert('RGB').save(path_picture_output + '{}.jpg'.format(district + num), 'jpeg')
                        print('picture success')
                        num += str(mumber)
                    except Exception as E:
                        print('erro')
                        print('picture erro')
                        image.convert('RGB').save(path_picture_output + '{}.jpg'.format(district + num), 'jpeg')
                else:
                    num = '小'
                    try:
                        image.convert('RGB').save(path_picture_output + '{}.jpg'.format(district + num), 'jpeg')
                        print('picture success')
                        num += str(mumber)
                    except Exception as E:
                        print('erro')
                        print('picture erro')
                        image.convert('RGB').save(path_picture_output + '{}.jpg'.format(district + num), 'jpeg')


            else:
                pass
    # 数据先需要从SAP导出来 之后刷新数据透视表，如果做成VBA小程序会不会更好
    # VBA思路 收集到SAP数据 在该文件中点击按钮启动宏，完成分配。
    ex.Quit()


def kill_excel():
    # 先清理一下可能存在的Excel进程
    pids = psutil.pids()
    for pid in pids:
        try:
            p = psutil.Process(pid)
            # print('pid=%s,pname=%s' % (pid, p.name()))
            # 关闭excel进程
            if p.name() == 'EXCEL.EXE':
                cmd = 'taskkill /F /IM EXCEL.EXE'
                os.system(cmd)
        except Exception as e:
            print('似乎出了个问题')


def send_email(district, to_list, cc_list):
    today = str(time.localtime()[1]) + '/' + str(time.localtime()[2])
    outlook = win32.Dispatch('outlook.application')
    imageFile1 = os.path.abspath(path_picture_output + '%s大.jpg' % district)
    imageFile2 = os.path.abspath(path_picture_output + '%s小.jpg' % district)
    mail = outlook.CreateItem(0)
    ats = mail.Attachments
    att1 = ats.Add(imageFile1, 1, 0)
    att2 = ats.Add(imageFile2, 1, 0)
    mail.To = str(to_list).replace("'", '').replace("[", '').replace("]", '').replace(",", ';')  # 收件人
    mail.CC = str(cc_list).replace("'", '').replace("[", '').replace("]", '').replace(",", ';')  # 抄送邮箱列表
    # mail.BCC = "test@outlook.com"  # 密抄邮箱列表，谨慎使用
    mail.Subject = 'trading 未系统收货清单 %s - 销售part - %s' % (today, district)  # 邮件主题
    mail.HTMLBody = '''<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">
<head>
    <meta http-equiv=Content-Type content="text/html; charset=gb2312">
    <meta name=Generator content="Microsoft Word 15 (filtered medium)">
    <!--[if !mso]>
    <style>v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}

    </style><![endif]-->
    <style><!--
/* Font Definitions */
@font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;}
@font-face
	{font-family:宋体;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:等线;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:"\@等线";
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:"\@宋体";
	panose-1:2 1 6 0 3 1 1 1 1 1;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0cm;
	text-align:justify;
	text-justify:inter-ideograph;
	font-size:10.5pt;
	font-family:"Calibri",sans-serif;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:#0563C1;
	text-decoration:underline;}
span.EmailStyle17
	{mso-style-type:personal-compose;
	font-family:"Calibri",sans-serif;
	color:windowtext;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-family:"Calibri",sans-serif;}
.MsoPapDefault
	{mso-style-type:export-only;
	text-align:justify;
	text-justify:inter-ideograph;}
/* Page Definitions */
@page WordSection1
	{size:612.0pt 792.0pt;
	margin:72.0pt 90.0pt 72.0pt 90.0pt;}
div.WordSection1
	{page:WordSection1;}
-->
    </style>
    <!--[if gte mso 9]>
    <xml>
        <o:shapedefaults v:ext="edit" spidmax="1026"/>
    </xml><![endif]--><!--[if gte mso 9]>
    <xml>
        <o:shapelayout v:ext="edit">
            <o:idmap v:ext="edit" data="1"/>
        </o:shapelayout>
    </xml><![endif]--></head>
<body lang=ZH-CN link="#0563C1" vlink="#954F72" style='word-wrap:break-word;text-justify-trim:punctuation'>
<div class=WordSection1>
    <p class=MsoNormal><b><span lang=EN-US style='font-size:12.0pt'>Dears</span></b><span lang=EN-US><o:p></o:p></span>
    </p>
    <p class=MsoNormal><span lang=EN-US>&nbsp;<o:p></o:p></span></p>
    <p class=MsoNormal><span style='font-family:等线'>各位之所以收到此邮件是因为各位<span style='background:yellow;mso-highlight:yellow'>曾经入过</span></span><span
            lang=EN-US>trading</span><span style='font-family:等线'>相关产品订单。</span><span lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US>&nbsp;<o:p></o:p></span></p>
    <p class=MsoNormal><span style='font-family:等线;background:yellow;mso-highlight:yellow'>附件是截止<span
            lang=EN-US>%s</span></span><span style='font-family:等线'>，</span><span lang=EN-US>SAP</span><span
            style='font-family:等线'>系统还<span style='background:yellow;mso-highlight:yellow'>未收货</span>的<span
            style='background:yellow;mso-highlight:yellow'>工程产品</span>订单清单。</span><span lang=EN-US><o:p></o:p></span>
    </p>
    <p class=MsoNormal style='text-indent:21.0pt'><span lang=EN-US style='font-size:10.0pt'>1. </span><span
            style='font-size:10.0pt;font-family:等线'>并非全部</span><span lang=EN-US
                                                                     style='font-size:10.0pt'>trading</span><span
            style='font-size:10.0pt;font-family:等线'>产品订单<span lang=EN-US> (</span>只含工程产品<span lang=EN-US>)</span></span><span
            lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal style='text-indent:21.0pt'><span lang=EN-US style='font-size:10.0pt'>2. </span><span
            style='font-size:10.0pt;font-family:等线'>已收货订单不在此清单中</span><span lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal style='text-indent:21.0pt'><span lang=EN-US
                                                        style='font-size:10.0pt;font-family:等线'>3. </span><span
            style='font-size:10.0pt;font-family:等线'>已剔除<span lang=EN-US>%s</span>当天系统收货订单</span><span lang=EN-US><o:p></o:p></span>
    </p>
    <p class=MsoNormal style='text-indent:21.0pt'><span lang=EN-US style='font-size:10.0pt;font-family:等线'>&nbsp;</span><span
            lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal><span style='font-family:等线'>作为工作日常，我们会与供应商沟通入单，出货，出货后回单的催缴及后续系统收货已提现销量等操作，但由于物流回单返回战线比较长，速度相对较慢，如清单内各位可以协调客户提供签收回单，并电子档邮件给到我们，我们也可以作为收货依据进行系统收货（<span
            style='background:yellow;mso-highlight:yellow'>需清晰可辨认</span>），以便于以<span
            style='background:yellow;mso-highlight:yellow'>最快速度体现销量</span>，还请各位</span><span
            lang=EN-US>support</span><span style='font-family:等线'>。</span><span lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US>&nbsp;<o:p></o:p></span></p>
    <p class=MsoNormal><span style='font-family:等线'>以上如有任何问题，请保持联系</span><span lang=EN-US>~ <o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US>&nbsp;<o:p></o:p></span></p>
    <p class=MsoNormal><span style='font-family:等线'>邮件回复，可联系各自对应的</span><span lang=EN-US>trading</span><span
            style='font-family:等线'>计划员，</span><span lang=EN-US>thanks~<o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US>&nbsp;<o:p></o:p></span></p>
    <p class=MsoNormal><b><span lang=EN-US style='color:#ED7D31'>Textur</span></b><b><span lang=EN-US
                                                                                           style='font-family:等线;color:#ED7D31'>e</span></b><b><span
            style='font-family:等线;color:#ED7D31'>相关产品计划邮箱：</span></b><span lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US>Meng, H. (Emily) <a
            href="mailto:hui.meng@akzonobel.com">hui.meng@akzonobel.com</a> <o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US>Jia Yin Li <a href="mailto:jiayin.li@akzonobel.com">jiayin.li@akzonobel.com</a><o:p></o:p></span>
    </p>
    <p class=MsoNormal><span lang=EN-US>Zhu, X.(Eric)</span><b><span lang=EN-US
                                                                     style='font-family:等线;color:#ED7D31'> </span></b><span
            class=MsoHyperlink><span lang=EN-US><a
            href="mailto:eirc.zhu@akzonobel.com">eirc.zhu@akzonobel.com</a></span></span><span lang=EN-US><o:p></o:p></span>
    </p>
    <p class=MsoNormal><span lang=EN-US>&nbsp;<o:p></o:p></span></p>
    <p class=MsoNormal><b><span style='font-family:等线;color:#ED7D31'>腻子相关产品计划邮箱：</span></b><span lang=EN-US><o:p></o:p></span>
    </p>
    <p class=MsoNormal><span lang=EN-US>He, B. (Sarah) <a href="mailto:bin.bpo.he@akzonobel.com">bin.bpo.he@akzonobel.com</a><o:p></o:p></span>
    </p>
    <p class=MsoNormal><span lang=EN-US>&nbsp;<o:p></o:p></span></p>
    <p class=MsoNormal><b><span lang=EN-US style='font-family:等线;color:#ED7D31'>HPC</span></b><b><span
            style='font-family:等线;color:#ED7D31'>相关产品计划邮箱：</span></b><span lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US>Zhu, X.(Eric)</span><b><span lang=EN-US
                                                                     style='font-family:等线;color:#ED7D31'> </span></b><span
            class=MsoHyperlink><span lang=EN-US><a
            href="mailto:eirc.zhu@akzonobel.com">eirc.zhu@akzonobel.com</a></span></span><span lang=EN-US><o:p></o:p></span>
    </p>
    <p class=MsoNormal><span lang=EN-US>&nbsp;<o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US style='font-family:等线'>&nbsp;</span><span lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal><b><span style='font-size:14.0pt;font-family:等线;background:yellow;mso-highlight:yellow'>附件是最新的未收货清单，请主要关注这些已出货未有回单的订单，谢谢<span
            lang=EN-US>~</span></span></b><span lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal><b><span lang=EN-US style='font-size:11.0pt;color:white;background:blue;mso-highlight:blue'>Texture&amp; putty </span></b><b><span
            lang=EN-US
            style='font-size:11.0pt;font-family:Wingdings;color:white;background:blue;mso-highlight:blue'>è</span></b><b><span
            lang=EN-US
            style='font-size:11.0pt;color:white;background:blue;mso-highlight:blue'> &nbsp;</span></b><b><span
            style='font-size:11.0pt;font-family:等线;color:white;background:blue;mso-highlight:blue'>实物出货 </span></b><b><span
            lang=EN-US style='font-size:11.0pt;color:white;background:blue;mso-highlight:blue'>vs </span></b><b><span
            style='font-size:11.0pt;font-family:等线;color:white;background:blue;mso-highlight:blue'>订单录入 </span></b><b><span
            lang=EN-US style='font-size:11.0pt;color:white;background:blue;mso-highlight:blue'>&gt; 3</span></b><b><span
            style='font-size:11.0pt;font-family:等线;color:white;background:blue;mso-highlight:blue'>周的订单（蓝色部分）</span></b><b><span
            lang=EN-US style='font-size:11.0pt;color:white;background:blue;mso-highlight:blue'>, </span></b><b><span
            style='font-size:11.0pt;font-family:等线;color:white;background:blue;mso-highlight:blue'>请销售</span></b><b><span
            style='font-size:16.0pt;font-family:等线;color:#ED7D31;background:blue;mso-highlight:blue'>优先帮忙</span></b><b><span
            style='font-size:11.0pt;font-family:等线;color:white;background:blue;mso-highlight:blue'>协调客户反馈回单，便于系统收货，体现销量，</span></b><b><span
            lang=EN-US style='font-size:11.0pt;color:white;background:blue;mso-highlight:blue'>thanks~</span></b><span
            lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal style='mso-margin-top-alt:auto;margin-bottom:14.4pt'><span lang=EN-US><img border=0 width=834
                                                                                                  height=725
                                                                                                  style='width:8.6875in;height:7.5555in'
                                                                                                  id="图片"
                                                                                                  src="%s"><img
            border=0 width=339 height=295 style='width:3.5347in;height:3.0763in' id="图片2"
            src="%s"></span><span lang=EN-US
                                                                  style='font-size:9.0pt;font-family:"Arial",sans-serif'><o:p></o:p></span>
    </p>
    <p class=MsoNormal style='mso-margin-top-alt:auto;margin-bottom:14.4pt'><b><span lang=EN-US
                                                                                     style='font-size:9.0pt;font-family:"Arial",sans-serif;color:#244061'>Thanks!</span></b><span
            lang=EN-US> <o:p></o:p></span></p>
    <p class=MsoNormal style='mso-margin-top-alt:auto;margin-bottom:14.4pt'><b><span lang=EN-US
                                                                                     style='font-size:9.0pt;font-family:"Arial",sans-serif;color:#244061'>B.rgds/Tracy Dong</span></b><span
            lang=EN-US style='font-size:9.0pt;font-family:"Arial",sans-serif;color:#244061'><br clear=all>Supply Chain - Project Channel </span><span
            lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal style='mso-margin-top-alt:auto;margin-bottom:14.4pt'><span lang=DE
                                                                                  style='font-size:8.0pt;font-family:"Arial",sans-serif;color:#244061'>T&nbsp; +86 21 37606831<br
            clear=all>M +139-1625-4904</span><span lang=DE
                                                   style='font-size:8.0pt;font-family:"Arial",sans-serif;color:#383838'> <br
            clear=all></span><span lang=DE
                                   style='font-size:8.0pt;font-family:"Arial",sans-serif;color:#244061'>E&nbsp;</span><span
            lang=DE style='font-size:8.0pt;font-family:"Arial",sans-serif;color:#383838'> </span><span lang=EN-US
                                                                                                       style='font-size:8.0pt;font-family:"Arial",sans-serif;color:#0092BB'><a
            href="mailto:Tracy.dong@akzonobel.com"><span lang=DE style='color:blue'>Tracy.dong@akzonobel.com</span></a></span><span
            lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal style='mso-margin-top-alt:auto;margin-bottom:14.4pt'><span lang=EN-US
                                                                                  style='font-size:8.0pt;font-family:"Arial",sans-serif;color:#948A54'>Focus on solution instead of situation</span><span
            lang=EN-US><o:p></o:p></span></p>
    <p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p>
</div>
</body>
</html>''' % (today, today, att1, att2)
    path = os.path.abspath(path_excel_output + '%s - trading pending order to 销售 %s.xlsx' % (district, email_day))
    mail.Attachments.Add(path)
    mail.Display()


def sale_df(region):  # 获得这次数据的区域

    # print('sale_df')
    df = Excel_data[Excel_data['供货region'] == region]

    data_write = df.to_numpy()

    # print(data_write)
    return data_write


def xl_open(regions, origin_file):
    wb = xw.Book(origin_file)  # 打开的源文件
    sht = wb.sheets[0]  # 选择源文件的第一个sheet
    sheet = wb.sheets[1]  # xlwings 操纵excel

    for region in regions:
        row = sht.used_range.last_cell.row
        sht['2:%d' % row].delete()  # 使用行最大值删掉
        sht.range('A2').value = sale_df(region)
        # 会存在一个问题，如果后续改透视表则会出现无法匹配到该透视表的问题
        sheet.api.PivotTables("数据透视表4").PivotCache().Refresh()
        sheet.api.PivotTables("数据透视表4").PivotFields("供货region").CurrentPage = "(All)"
        sheet.api.PivotTables("数据透视表4").PivotFields("供货region").PivotItems(region).Visible = True
        # 数据透视表6

        sheet.api.PivotTables("数据透视表4").PivotFields("供货region").ClearAllFilters
        sheet.api.PivotTables("数据透视表4").PivotFields("供货region").CurrentPage = region
        sheet.api.PivotTables("数据透视表26").PivotFields("供货region").ClearAllFilters
        sheet.api.PivotTables("数据透视表26").PivotFields("供货region").CurrentPage = region
        wb.save(path_excel_output + '%s - trading pending order to 销售 %s.xlsx' % (region, email_day))
    wb.close()  # 必须要关闭workbook 不然会出错


def init_start():
    '''
    如果是第一次启动则会弹出2个窗口
    :return:
    '''
    if len(files) <= 2:#防止第一次就有人把待分配文件也放进去了
        win32api.MessageBox(0, '欢迎使用邮件分配小程序\n请仔细看使用说明', '运行前重要提示', win32con.MB_ICONEXCLAMATION)


        dir_check()
        text_create()
        win32api.MessageBox(0, '请维护好"源文件"文件夹内的邮箱Excel，\n(请勿改名)如果丢失第二次运行程序会自动生成再重新维护', '欢迎第一次使用',
                            win32con.MB_ICONEXCLAMATION)
        win32api.MessageBox(0, '初始化完毕\n请仔细看使用说明', '运行成功', win32con.MB_ICONEXCLAMATION)
        quit()
    else:
        files_level2 = os.listdir(path_origin + '\\' + '源文件')
        #获取源文件的文件数

        if '邮箱.xlsx' in files_level2:
            #print('在里面')
                #如果邮箱在该文件夹里则pass
            pass
        else:
            win32api.MessageBox(0, '未检测到邮箱.xlsx文件', '文件缺失提醒', win32con.MB_ICONEXCLAMATION)
            create_email_excel()
            win32api.MessageBox(0, '因为文件缺失，请重新运行点击"确定"', '错误', win32con.MB_ICONEXCLAMATION)
            quit()



def text_create():
    file = open('使用说明.txt', 'w')
    text = '''文件运行中可以智能检测：
常见问题：
	1，邮箱.xlsx文件没有正确维护  区域全国为每次都需要CC （大老板）
	2，需要分配的文件没有放在exe文件目录
	3，要点：excel内容被改动，此程序对excel识别精确，截取图片需要第二张表（table）中图片为2张图片，不能组合在一起。如有疑问可在随exe文件一并发来的模板里进行查询。



每次将需要分配的excel放到exe目录，运行之后的分配源文件将会被移动到源文件历史中。
如有疑问 Email：gordon.guo@akzonobel.com
'''
    file.write(text)

    file.close()


def create_email_excel():
    data_email_excel = {'发送规则': ['to', 'cc'], '区域': ['东区', '全国'], '销售名': ['如有', '疑问'],
                        '销售邮件': ['gordon.guo@akzonobel.com', 'guoxin12980@icloud.com']}
    df = pd.DataFrame(data_email_excel)
    df.to_excel(path_email, index=False)
    win32api.MessageBox(0, '已为您生成邮件.xlsx文件请后期维护好勿修改文件名称', '文件创建成功', win32con.MB_ICONEXCLAMATION)


def dir_check():
    sys_time = time.localtime()
    year = sys_time[0]
    month = sys_time[1]
    day = sys_time[2]
    time_dir = str(year) + '-' + str(month) + '-' + str(day)
    check_list_level1 = ['源文件', '输出文件']
    check_list_output = ['Excel', '图片']
    check_list_origin = ['源文件历史']
    path_origin = os.getcwd()
    for i in check_list_level1:  # 在 源文件 和输出文件夹
        path = path_origin  # 初始化path
        dir_maker(path, i)  # 创建 源文件 和输出文件夹
        path = path_origin + '\\' + i  # 赋值新地址
        if i == '源文件':
            for o_i in check_list_origin:
                dir_maker(path, o_i)
                if '邮箱.xlsx' in os.listdir(path):
                    pass
                else:
                    #win32api.MessageBox(0, '未检测到邮箱.xlsx文件', '文件缺失提醒', win32con.MB_ICONEXCLAMATION)
                    create_email_excel()
                path += '\\' + o_i
                dir_maker(path, time_dir)
        elif i == '输出文件':
            for output in check_list_output:
                dir_maker(path, output)
                path_temp = path + '\\' + output
                dir_maker(path_temp, time_dir)


def auto_select_excel(files):
    for i in files:
        if '.xlsx' in i:
            return i
        else:
            pass


def get_email_data(region):
    to_list_name = list(set(Excel_data[Excel_data['供货region'] == region]['客户销售联系人名字'].values.tolist()))
    Email_region = Email_data[Email_data['区域'] == region]
    Email_region1 = Email_data[Email_data['区域'] == '全国']
    to_name = Email_data['销售名'].values.tolist()
    to_email = Email_data['销售邮件'].values.tolist()
    dict_data = dict(zip(to_name, to_email))
    to_list = []
    for i in to_list_name:
        to_list.append(dict_data[i])
    cc_list1 = Email_region[['发送规则', '销售邮件']][Email_region['发送规则'] == 'cc']['销售邮件'].values.tolist()
    cc_list2 = Email_region1[['发送规则', '销售邮件']][Email_region1['发送规则'] == 'cc']['销售邮件'].values.tolist()
    cc_list = cc_list1 + cc_list2
    return to_list, cc_list


path_origin = os.getcwd()  # 获取当前位置
files = os.listdir(path_origin)  # 检测当前文件夹
sys_time = time.localtime()
year = sys_time[0]
month = sys_time[1]
day = sys_time[2]
email_day = str(month) + '-' + str(day)
time_dir = str(year) + '-' + str(month) + '-' + str(day)
path_email = path_origin + '\\' + '源文件' + '\\' + '邮箱.xlsx'
path_excel_output = path_origin + '\\' + '输出文件' + '\\' + 'Excel' + '\\' + time_dir + '\\'
path_picture_output = path_origin + '\\' + '输出文件' + '\\' + '图片' + '\\' + time_dir + '\\'
path_history_excel = path_origin + '\\' + '源文件' + '\\' + '源文件历史' + '\\' + time_dir + '\\'
init_start()  # 判断然后初始化生成文件夹
text_create()  # 每次都创建使用说明
dir_check()  # 对文件夹内进行判断 顺便增加当日的文件夹 如果缺失则增加文件

win32api.MessageBox(0, '欢迎使用邮件分配小程序\n请先关闭所有正在运行的Excel再点击"确定"', '运行前重要提示', win32con.MB_ICONEXCLAMATION)
kill_excel()
origin_file = auto_select_excel(files)

xw.App(visible=False)  # 设置xlwings为不可见
try:
    Excel_data = pd.read_excel(origin_file, sheet_name=0)  # 读源文件
except Exception as E:
    win32api.MessageBox(0, '当前文件夹未检测到需要分配的文件,点击"确定”退出', '错误', win32con.MB_ICONEXCLAMATION)
    quit()
regions = list(set(Excel_data['供货region'].values.tolist()))

Email_data = pd.read_excel(path_email, sheet_name=0)
if len(Email_data)<=5:
    win32api.MessageBox(0, '请核实邮箱.xlsx是否正确维护', '错误', win32con.MB_ICONEXCLAMATION)
    quit()


ex = EnsureDispatch('Excel.Application')
xl_open(regions, origin_file=origin_file)
for i in regions:
    get_picture(i)
    to_cc = get_email_data(i)
    send_email(i, to_cc[0], to_cc[1])
try:
    shutil.move(path_origin + '\\' + origin_file, path_history_excel)
    win32api.MessageBox(0, '此次分配文件已经移动到"源文件->源文件历史文件夹"中如需查看请在该文件夹查看', '服务结束', win32con.MB_ICONEXCLAMATION)
except Exception as E:
    win32api.MessageBox(0, '该文件已经存在，如需保留请重命名', '运行前重要提示', win32con.MB_ICONEXCLAMATION)
