# -*- coding: cp936 -*-
import win32com
from win32com.client import Dispatch, constants
import sys
import os
import shutil
import _winreg
import urllib2
from xml.etree import ElementTree
import traceback
import pycurl
import StringIO
import easygui

global w


#=============================================================
host='http://pg.china-epli.com'
xmlname='/tempxml.xml'
#=============================================================

def repword(ts,ta):
    global w

    if ta == None:
        ta = ''

    ta=ta.replace('\n','\r')
    
    for i in range(len(ta)/200+1):
        if i == len(ta)/200:
            w.Selection.Find.Execute(ts, False, False, False, False, False, True, 1, True, ta[i*200:i*200+200], 2)
        else:
            w.Selection.Find.Execute(ts, False, False, False, False, False, True, 1, True, ta[i*200:i*200+200] + ts, 2)


if os.path.exists('C:\\EiaRpSys')==False:
    os.makedirs('C:\\EiaRpSys')
f1=open('C:\\EiaRpSys\\log.txt','w')



print 'v2.3'
f1.write('v2.3'+'\n')
easygui.msgbox('rprinter打印控件 版本v2.3'.decode('cp936'),'rprinter打印控件'.decode('cp936'))

for i in range(len(sys.argv)):
    print "参数", i, sys.argv[i]
    f1.write("参数"+str(i)+sys.argv[i]+'\n')

if len(sys.argv)>2:
    host=sys.argv[2]

if len(sys.argv)>3:
    xmlname=sys.argv[3]
    
wordurl = host+'/mouldboard.doc'
xmlurl = host + xmlname

if 'plugin' in sys.argv[0].lower():

    try:
        try:
            print 'begin download'
            f1.write('begin download'+'\n')
            usock = urllib2.urlopen(xmlurl)
            xmlstring = usock.read()
            usock.close()
        except:
            print 'change curl'
            f1.write('change curl'+'\n')
            c = pycurl.Curl()
            c.setopt(pycurl.URL, xmlurl)
            c.setopt(pycurl.HTTPHEADER, ["Accept:"])
            b = StringIO.StringIO()
            c.setopt(pycurl.WRITEFUNCTION, b.write)
            c.setopt(pycurl.FOLLOWLOCATION, 1)
            c.setopt(pycurl.MAXREDIRS, 5)
            c.perform()
            xmlstring = b.getvalue()
            c.close()
            
        xmlstring=xmlstring.replace('''encoding="gb2312"''','''encoding="utf-8"''').decode('gbk','xmlcharrefreplace').encode('utf-8')
        root = ElementTree.fromstring(xmlstring)        
        print 'opened xml',xmlurl
        f1.write('opened xml '+xmlurl+'\n')
        

        # 后台运行，不显示，不警告
        w = win32com.client.DispatchEx('Word.Application')
        w.Visible = 0
        w.DisplayAlerts = 0
        print 'started Word'
        f1.write('started Word'+'\n')

        # 下载文件模版到本地
        try:
            print 'begin download'
            f1.write('begin download'+'\n')
            re = urllib2.Request(wordurl)
            rs = urllib2.urlopen(re).read()
        except:
            print 'change curl'
            f1.write('change curl'+'\n')
            c = pycurl.Curl()
            c.setopt(pycurl.URL, wordurl)
            c.setopt(pycurl.HTTPHEADER, ["Accept:"])
            b = StringIO.StringIO()
            c.setopt(pycurl.WRITEFUNCTION, b.write)
            c.setopt(pycurl.FOLLOWLOCATION, 1)
            c.setopt(pycurl.MAXREDIRS, 5)
            c.perform()
            rs = b.getvalue()
            c.close()
        open('C:\\EiaRpSys\\mouldboard.doc', 'wb').write(rs)
        
        print 'downloaded word',wordurl
        f1.write('downloaded word '+ wordurl+'\n')

        
        # 打开新的文件
        doc = w.Documents.Add('C:\\EiaRpSys\\mouldboard.doc')# 创建新的文档
        #worddoc = w.Documents.Open(wordurl) # 打开文档
        print 'opened Documents'
        f1.write('opened Documents'+'\n')

        '''
        # 插入文字
        myRange = doc.Range(0,0)
        myRange.InsertBefore('Hello from Python!')
        # 使用样式
        wordSel = myRange.Select()
        wordSel.Style = constants.wdStyleHeading1
        # 正文文字替换
        w.Selection.Find.Clearformatting()
        w.Selection.Find.Replacement.Clearformatting()
        w.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1, True, NewStr, 2)
        # 页眉文字替换
        w.ActiveDocument.Sections[0].Headers[0].Range.Find.Clearformatting()
        w.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.Clearformatting()
        w.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(OldStr, False, False, False, False, False, True, 1, False, NewStr, 2)
        # 表格操作
        doc.Tables[0].Rows[0].Cells[0].Range.text ='123123'
        worddoc.Tables[0].Rows.Add() # 增加一行
        # 转换为html
        wc = win32com.client.constants
        w.ActiveDocument.WebOptions.RelyOnCSS = 1
        w.ActiveDocument.WebOptions.OptimizeforBrowser = 1
        w.ActiveDocument.WebOptions.BrowserLevel = 0 # constants.wdBrowserLevelV4
        w.ActiveDocument.WebOptions.OrganizeInFolder = 0
        w.ActiveDocument.WebOptions.UseLongFileNames = 1
        w.ActiveDocument.WebOptions.RelyOnVML = 0
        w.ActiveDocument.WebOptions.AllowPNG = 1
        w.ActiveDocument.SaveAs( FileName = filenameout, Fileformat = wc.wdformatHTML )
        '''
        
        repword("a_year2", root.find('a_year2').text)
        
        repword("a_month2", root.find('a_month2').text)
        
        repword("y_rs", root.find('y_rs').text)
        
        if root.find('a_lx').text == 'False':
            repword("■承保前评估        □承保后第一次服务", "□承保前评估        ■承保后第一次服务")
            
        if root.find('e_gczdwxy').text == 'True':
            repword("/该企业无重大危险源", "")
        else:
            repword("该企业已构成重大危险源/", "")

        if root.find('e_gwgy').text == "" or root.find('e_gwgy').text == None :
            repword("/该厂生产过程中的e_gwgy为高危工艺", "")
        else:
            repword("该厂生产过程中不涉及危险性化工工艺/", "")
            repword("e_gwgy", root.find('e_gwgy').text)

        print "s0"
        f1.write('s0'+'\n')
        
        repword("a_year", root.find('a_year').text)
        repword("a_month", root.find('a_month').text)
        repword("a_day", root.find('a_day').text)
        repword("xa_szqy", root.find('xa_szqy').text) #------所在区域
        repword("a_dwdz", root.find('a_dwdz').text)
        repword("a_lxr", root.find('a_lxr').text)
        repword("a_zbdh", root.find('a_zbdh').text)
        repword("a_czhm", root.find('a_czhm').text)
        repword("a_cqmj", root.find('a_cqmj').text)
        repword("a_zgrs", root.find('a_zgrs').text)
        repword("e_zygy", root.find('e_zygy').text)
        #repword("a_yqschjmgbs", root.find('a_yqschjmgbs').text)    #在表格处理后替换
        repword("a_xckcsm", root.find('a_xckcsm').text)
        repword("f_fs", root.find('f_fs').text)
        repword("g_fq", root.find('g_fq').text)
        repword("h_gf", root.find('h_gf').text)
        repword("j_xcglqk", root.find('j_xcglqk').text)
        repword("xi_sm1", root.find('xi_sm1').text)
        repword("xi_sm2", root.find('xi_sm2').text)
        repword("i_sm", root.find('i_sm').text)
        repword("m_fxdj", root.find('m_fxdj').text)
        repword("xa_jybe", root.find('xa_jybe').text) #------建议保额
        repword("n_qtjy", root.find('n_qtjy').text)

        print "s1"
        f1.write('s1'+'\n')

        for i in range(3):
            if root.find("o_zj" + str(i)).text <> "" and root.find("o_zj" + str(i)).text <> None:
                repword("o_zj" + str(i), root.find("o_zj" + str(i)).text)
                repword("o_zc" + str(i), root.find("o_zc" + str(i)).text)
            else:
                w.Selection.Find.Execute("o_zj" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()
        
        for i in range(3):
            if root.find("o_ry" + str(i)).text <> "" and root.find("o_ry" + str(i)).text <> None:
                repword("o_ry" + str(i), root.find("o_ry" + str(i)).text)
                repword("o_zyfx" + str(i), root.find("o_zyfx" + str(i)).text)
            else:
                w.Selection.Find.Execute("o_ry" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()

        print "s2"
        f1.write('s2'+'\n')
        
        for i in range(15):
            if root.find("c_mc" + str(i)).text <> "" and root.find("c_mc" + str(i)).text <> None:
                repword("c_mc" + str(i) + "|", root.find("c_mc" + str(i)).text)
                repword("c_ncl" + str(i) + "|", root.find("c_ncl" + str(i)).text)
                repword("c_cyfs" + str(i) + "|", root.find("c_cyfs" + str(i)).text)
            else:
                w.Selection.Find.Execute("c_mc" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()
            

        print "s3"
        f1.write('s3'+'\n')

        for i in range(15):
            if root.find("b_mc" + str(i)).text <> "" and root.find("b_mc" + str(i)).text <> None:
                repword("b_mc" + str(i) + "|", root.find("b_mc" + str(i)).text)
                repword("b_nyl" + str(i) + "|", root.find("b_nyl" + str(i)).text)
                repword("b_cyfs" + str(i) + "|", root.find("b_cyfs" + str(i)).text)
            else:
                w.Selection.Find.Execute("b_mc" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()
            

        print "s4"
        f1.write('s4'+'\n')

        for i in range(15):
            if root.find("d_mc" + str(i)).text <> "" and root.find("d_mc" + str(i)).text <> None:
                repword("d_mc" + str(i) + "|", root.find("d_mc" + str(i)).text)
                repword("d_nyl" + str(i) + "|", root.find("d_nyl" + str(i)).text)
                repword("d_cyfs" + str(i) + "|", root.find("d_cyfs" + str(i)).text)
            else:
                w.Selection.Find.Execute("d_mc" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()

        print "s5"
        f1.write('s5'+'\n')
        
        fwnum=0
        
        if root.find("a_d1").text <> "" and root.find("a_d1").text <> None:
            repword("a_d1", root.find("a_d1").text)
            repword("a_d3", root.find("a_d3").text)
            fwnum=fwnum+1
        else:
            w.Selection.Find.Execute("a_d1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.Rows.Delete()
        
        if root.find("a_n1").text <> "" and root.find("a_n1").text <> None:
            repword("a_n1", root.find("a_n1").text)
            repword("a_n3", root.find("a_n3").text)
            fwnum=fwnum+1
        else:
            w.Selection.Find.Execute("a_n1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.Rows.Delete()
        
        if root.find("a_x1").text <> "" and root.find("a_x1").text <> None:
            repword("a_x1", root.find("a_x1").text)
            repword("a_x3", root.find("a_x3").text)
            fwnum=fwnum+1
        else:
            w.Selection.Find.Execute("a_x1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.Rows.Delete()
        
        if root.find("a_b1").text <> "" and root.find("a_b1").text <> None:
            repword("a_b1", root.find("a_b1").text)
            repword("a_b3", root.find("a_b3").text)
            fwnum=fwnum+1
        else:
            w.Selection.Find.Execute("a_b1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.Rows.Delete()
        
        if root.find("a_dn1").text <> "" and root.find("a_dn1").text <> None:
            repword("a_dn1", root.find("a_dn1").text)
            repword("a_dn3", root.find("a_dn3").text)
            fwnum=fwnum+1
        else:
            w.Selection.Find.Execute("a_dn1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.Rows.Delete()

        if root.find("a_xn1").text <> "" and root.find("a_xn1").text <> None:
            repword("a_xn1", root.find("a_xn1").text)
            repword("a_xn3", root.find("a_xn3").text)
            fwnum=fwnum+1
        else:
            w.Selection.Find.Execute("a_xn1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.Rows.Delete()
    
        if root.find("a_db1").text <> "" and root.find("a_db1").text <> None:
            repword("a_db1", root.find("a_db1").text)
            repword("a_db3", root.find("a_db3").text)
            fwnum=fwnum+1
        else:
            w.Selection.Find.Execute("a_db1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.Rows.Delete()

        if root.find("a_xb1").text <> "" and root.find("a_xb1").text <> None:
            repword("a_xb1", root.find("a_xb1").text)
            repword("a_xb3", root.find("a_xb3").text)
            fwnum=fwnum+1
        else:
            w.Selection.Find.Execute("a_xb1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.Rows.Delete()
        
        if fwnum>0:
            w.Selection.Find.Execute("a_yqschjmgbs", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.MoveLeft(Unit=1,Count=4)
            w.Selection.MoveDown(Unit=5,Count=1)
            w.Selection.MoveDown(Unit=5,Count=fwnum-1,Extend=1)
            w.Selection.MoveRight(Unit=1,Count=2, Extend=1)
            w.Selection.Cut()
            w.Selection.MoveUp(Unit=5,Count=1)
            w.Selection.Paste()
            w.Selection.MoveRight(Unit=1,Count=1)
            w.Selection.MoveDown(Unit=5,Count=fwnum)
            w.Selection.Rows.Delete()
        
        repword("a_yqschjmgbs", root.find('a_yqschjmgbs').text)
        

        print "s6"
        f1.write('s6'+'\n')
        
        if root.find("a_xckcsm").text == "" or root.find("a_xckcsm").text == None:
            repword("现场勘查说明：^p", "")
        
        for i in range(10):
            if root.find("k_hpypfdyq" + str(i)).text <> "" and root.find("k_hpypfdyq" + str(i)).text <> None:
                repword("k_hpypfdyq" + str(i), root.find("k_hpypfdyq" + str(i)).text)
                repword("k_lsqk" + str(i), root.find("k_lsqk" + str(i)).text)
            else:
                w.Selection.Find.Execute("k_hpypfdyq" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()

        print "s7"
        f1.write('s7'+'\n')
        
        for i in range(10):
            if root.find("e_wzmc" + str(i)).text <> "" and root.find("e_wzmc" + str(i)).text <> None:
                repword("e_wzmc" + str(i), root.find("e_wzmc" + str(i)).text)
                repword("e_cnunh" + str(i), root.find("e_cnunh" + str(i)).text)
                repword("e_zt" + str(i), root.find("e_zt" + str(i)).text)
                repword("e_sd" + str(i), root.find("e_sd" + str(i)).text)
                repword("e_rsx" + str(i), root.find("e_rsx" + str(i)).text)
                repword("e_bzjx" + str(i), root.find("e_bzjx" + str(i)).text)
                repword("e_hzwxtx" + str(i), root.find("e_hzwxtx" + str(i)).text)
                repword("e_ld50" + str(i), root.find("e_ld50" + str(i)).text)
                repword("e_dxfj" + str(i), root.find("e_dxfj" + str(i)).text)
            else:
                w.Selection.Find.Execute("e_wzmc" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()

        for i in range(10):     #fxwztxb2
            if root.find("e_wzmcx" + str(i)).text <> "" and root.find("e_wzmcx" + str(i)).text <> None:
                repword("e_wzmcx" + str(i), root.find("e_wzmcx" + str(i)).text)
                repword("xe_hjsj" + str(i), root.find("xe_hjsj" + str(i)).text)
                repword("xe_yjcs" + str(i), root.find("xe_yjcs" + str(i)).text)
            else:
                w.Selection.Find.Execute("e_wzmcx" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()
                
        for i in range(19):
            if root.find("n_xcfx" + str(i)).text <> "" and root.find("n_xcfx" + str(i)).text <> None:
                repword("n_xcfx" + str(i) + "|", root.find("n_xcfx" + str(i)).text)
                repword("n_knczdhjfx" + str(i) + "|", root.find("n_knczdhjfx" + str(i)).text)
                repword("n_gscsjjy" + str(i) + "|", root.find("n_gscsjjy" + str(i)).text)
            else:
                w.Selection.Find.Execute("n_xcfx" + str(i), False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()

        print "s8"
        f1.write('s8'+'\n')
        
        #现场照片
        picflag = 0
        picis = False
        for i in range(18):
            picflag = picflag + 1
            if root.find("imxc" + str(i)).text <> "" and root.find("imxc" + str(i)).text <> None:
                if root.find("tmxc" + str(i)).text == "" or root.find("tmxc" + str(i)).text == None:
                    w.Selection.Find.Execute("tmxc" + str(i)+ "|", False, False, False, False, False, True, 1, True, "", 0)
                    w.Selection.Rows.Delete()
                    w.Selection.Find.Execute("imxc" + str(i)+ "|", False, False, False, False, False, True, 1, True, "", 0)
                    w.Selection.Rows.Delete()
                else:
                    picis = True
                    repword("tmxc" + str(i) + "|",root.find("tmxc" + str(i)).text)
                    w.Selection.Find.Execute("imxc" + str(i) + "|", False, False, False, False, False, True, 1, True, "", 0)
                    for j in range (i + 1 - picflag,i+1):
                        if root.find("imxc" + str(j)).text <> "" and root.find("imxc" + str(j)).text <> None:
                            w.Selection.InlineShapes.AddPicture(root.find("imxc" + str(j)).text.replace('~',host))
                            w.Selection.MoveLeft(Unit=1, Count=1, Extend=1)
                            w.Selection.InlineShapes(1).Height = 150    #330 / picflag
                            w.Selection.InlineShapes(1).Width = 200     #445 / picflag
                            w.Selection.MoveRight(Unit=1, Count=1)
                            w.Selection.text=' '
                            w.Selection.MoveRight(Unit=1, Count=1)
                    picflag = 0
            else:
                w.Selection.Find.Execute("tmxc" + str(i)+ "|", False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()
                w.Selection.Find.Execute("imxc" + str(i)+ "|", False, False, False, False, False, True, 1, True, "", 0)
                w.Selection.Rows.Delete()
            
        
        #周边环境图
        if root.find("imzbhj").text <> "" and root.find("imzbhj").text <> None:
            w.Selection.Find.Execute("imzbhj", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.InlineShapes.AddPicture(root.find("imzbhj").text.replace('~',host))
            w.Selection.MoveLeft(Unit=1, Count=1, Extend=1)
            w.Selection.InlineShapes(1).Height = 300
            w.Selection.InlineShapes(1).Width = 400
            w.Selection.MoveRight(Unit=1, Count=1)
        else:
            repword("imzbhj","")
        
        #厂区平面图
        if root.find("imcqpm").text <> "" and root.find("imcqpm").text <> None:
            w.Selection.Find.Execute("imcqpm", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.InlineShapes.AddPicture(root.find("imcqpm").text.replace('~',host))
            w.Selection.MoveLeft(Unit=1, Count=1, Extend=1)
            w.Selection.InlineShapes(1).Height = 300
            w.Selection.InlineShapes(1).Width = 400
            w.Selection.MoveRight(Unit=1, Count=1)
        else:
            repword("imcqpm","")

        if (root.find("imzbhj").text == "" or root.find("imzbhj").text == None) and (root.find("imcqpm").text == "" or root.find("imcqpm").text == None):
            w.Selection.Find.Execute("附图1  a_name周边环境图", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.TypeBackspace()
            w.Selection.TypeBackspace()
        if root.find("imzbhj").text == "" or root.find("imzbhj").text == None:
            repword("附图1  a_name周边环境图","")
        if root.find("imcqpm").text == "" or root.find("imcqpm").text == None:
            repword("附图2  a_name厂区平面图","")

        print "s9"
        f1.write('s9'+'\n')
        
        for i in range(10):
            repword("l_zhhcnr" + str(i), root.find("l_zhhcnr" + str(i)).text)
            repword("l_zhxcjl" + str(i), root.find("l_zhxcjl" + str(i)).text)
            repword("l_zhbz" + str(i), root.find("l_zhbz" + str(i)).text)
        

        for i in range(19):
            repword("l_wxhcnr" + str(i) + "|", root.find("l_wxhcnr" + str(i)).text)
            repword("l_wxxcjl" + str(i) + "|", root.find("l_wxxcjl" + str(i)).text)
            repword("l_wxbz" + str(i) + "|", root.find("l_wxbz" + str(i)).text)
        
        for i in range(3):
            repword("l_schcnr" + str(i), root.find("l_schcnr" + str(i)).text)
            repword("l_scxcjl" + str(i), root.find("l_scxcjl" + str(i)).text)
            repword("l_scbz" + str(i), root.find("l_scbz" + str(i)).text)
        
        for i in range(2):
            repword("l_hjhcnr" + str(i), root.find("l_hjhcnr" + str(i)).text)
            repword("l_hjxcjl" + str(i), root.find("l_hjxcjl" + str(i)).text)
            repword("l_hjbz" + str(i), root.find("l_hjbz" + str(i)).text)
        
        for i in range(7):
            repword("l_jyhcnr" + str(i), root.find("l_jyhcnr" + str(i)).text)
            repword("l_jyxcjl" + str(i), root.find("l_jyxcjl" + str(i)).text)
            repword("l_jybz" + str(i), root.find("l_jybz" + str(i)).text)

        print "s10"
        f1.write('s10'+'\n')
        
        #调查表1
        if root.find("imhjdc1").text <> "" and root.find("imhjdc1").text <> None:
            w.Selection.Find.Execute("imhjdc1", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.InlineShapes.AddPicture(root.find("imhjdc1").text.replace('~',host))
            w.Selection.MoveLeft(Unit=1, Count=1, Extend=1)
            w.Selection.InlineShapes(1).Height = 300
            w.Selection.InlineShapes(1).Width = 400
            w.Selection.MoveRight(Unit=1, Count=1)
        else:
            repword("imhjdc1","")
        
        #调查表2
        if root.find("imhjdc2").text <> "" and root.find("imhjdc2").text <> None:
            w.Selection.Find.Execute("imhjdc2", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.InlineShapes.AddPicture(root.find("imhjdc2").text.replace('~',host))
            w.Selection.MoveLeft(Unit=1, Count=1, Extend=1)
            w.Selection.InlineShapes(1).Height = 300
            w.Selection.InlineShapes(1).Width = 400
            w.Selection.MoveRight(Unit=1, Count=1)
        else:
            repword("imhjdc2","")


        repword("a_name", root.find('a_name').text)
        

        print "s11"
        f1.write('s11'+'\n')
        if picis == False :
            w.Selection.Find.Execute("八、 现场照片", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.TypeBackspace()
            w.Selection.TypeBackspace()
            w.Selection.TypeBackspace()

            w.Selection.Find.Execute("八、现场照片", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.TypeBackspace()
            w.Selection.TypeBackspace()
            w.Selection.TypeBackspace()
            
        w.Selection.Find.Execute("目   录", False, False, False, False, False, True, 1, True, "", 0)
        w.Selection.MoveDown()
        w.Selection.MoveDown()
        w.Selection.MoveDown()
        w.Selection.Fields.Update()

        #封面
        if root.find("imfm").text <> "" and root.find("imfm").text <> None:
            w.Selection.Find.Execute("imfm", False, False, False, False, False, True, 1, True, "", 0)
            w.Selection.InlineShapes.AddPicture(root.find("imfm").text.replace('~',host))
            w.Selection.MoveLeft(Unit=1, Count=1, Extend=1)
            w.Selection.InlineShapes(1).Height = 280
            w.Selection.InlineShapes(1).Width = 350
            w.Selection.MoveRight(Unit=1, Count=1)
        else:
            repword("imfm","")
        
        print "s12"
        f1.write('s12'+'\n')

        w.Selection.HomeKey (6)
        while w.Selection.Find.Execute("|U*U|", False, False, True, False, False, True, 0, True, "", 0):
            w.Selection.Font.Superscript = True

        w.Selection.HomeKey (6)
        while w.Selection.Find.Execute("|D*D|", False, False, True, False, False, True, 0, True, "", 0):
            w.Selection.Font.Subscript = True

        repword("|U","")
        repword("U|","")
        repword("|D","")
        repword("D|","")
        

        print "s13"
        f1.write('s13'+'\n')


        w.Selection.HomeKey (6)
        if len(sys.argv)>2:
            if '0' in sys.argv[1]:
                doc.PrintOut()
                print 'printing'
                w.Documents.Close(0)
                w.Quit()
            else:
                w.Visible = 1
                print 'previewing'
        else:
            w.Visible = 1
            print 'previewing'

        print 'finish'
        f1.write('finish'+'\n')
        

    except:
        print 'run error'
        f1.write('run error'+'\n')

        traceback.print_exc()
        f1.write(traceback.format_exc()+'\n')
        easygui.msgbox(('生成word异常\n'+traceback.format_exc()).decode('cp936'),'rprinter打印控件'.decode('cp936'))


        try:
            w.Quit()
        except:
            pass

else:
    try:

        print 'begin copy'
        f1.write('begin copy'+'\n')
        if os.path.exists('C:\\Program Files\\plugin')==False:
            os.makedirs('C:\\Program Files\\plugin')
        shutil.copy(sys.argv[0],'C:\\Program Files\\plugin\\rprinter.exe')

        print 'begin regedit'
        f1.write('begin regedit'+'\n')

        key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE,r"Software\Classes")
        newKey = _winreg.CreateKey(key,"PLUGIN")
        _winreg.SetValueEx(newKey, "URL Protocol", 0, 1, "C:\\PROGRA~1\\plugin\\rprinter.exe %l")

        key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE,r"Software\Classes\PLUGIN")
        newKey = _winreg.CreateKey(key,"Shell")

        key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE,r"Software\Classes\PLUGIN\Shell")
        newKey = _winreg.CreateKey(key,"open")

        _winreg.SetValue(newKey,"command",1,"C:\\PROGRA~1\\PLUGIN\\RPRINTER.exe %l")

        print 'finish setup'
        f1.write('finish setup'+'\n')
        easygui.msgbox('控件安装成功'.decode('cp936'),'rprinter打印控件'.decode('cp936'))

    except:
        print 'setup error'
        f1.write('setup error'+'\n')
        
        print traceback.format_exc()
        f1.write(traceback.format_exc()+'\n')
        easygui.msgbox(('控件安装异常 请联系管理员手动安装\n'+traceback.format_exc()).decode('cp936'),'rprinter打印控件'.decode('cp936'))

f1.write('End\n')
f1.close()

try:
    raw_input('press Enter to close')
except:
    pass

