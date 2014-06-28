# --*-- coding: utf-8 --*--
#
# 遍历指定目录，修改文件名中含的` `及`()`为`_`
# 如 6912345678901 （1）.jpg （注意指定后缀名）改名为 6912345678901_1.jpg
#
# 另拷贝指定文件列表中的文件到指定目录
# -------------------------------------------------------------------
__version__ = "0.1.2"

from Tkinter import *
import ttk
import tkFileDialog
import os
import shutil
import re

from readxls import ReadData, AddContentFromXls

# import datetime
# import logging
#
# logging.basicConfig(format='%(asctime)s:%(levelname)s:%(name)s:%(message)s')
# logging.getLogger().setLevel(logging.DEBUG)
#
# log = logging.getLogger(__name__)

#6909931461116
#filename = tkFileDialog.asksaveasfilename()
#dirname = tkFileDialog.askdirectory()

DICT = {'default':[u'图片文件所在目录..', u'图片需求清单文件..', u"规范文件名", u"拷贝图片文件到..."],
        'mode':[u'基础资料文档文件', u'追加资料文档文件', u"规范文件名", u"查重及追加"]}

def as_unicode(s):
    try:
        s = s.decode('utf8')
        return s
    except:
        try:
            s = s.decode('gbk')
            return s
        except:
            return unicode(s)

basedir = as_unicode(os.path.abspath(os.path.dirname(__file__)))
if not os.path.exists(os.path.join(basedir,'sku.cfg')):
    config={}
else:
    config = dict( [ line.strip().split('=') for line in open(os.path.join(basedir,'sku.cfg')).readlines() ] )

def setConfig():
    f = open(os.path.join(basedir,'sku.cfg') ,'w')
    f.write(as_unicode('dir_select=%s'%config.get('dir_select','')).encode('utf8')+os.linesep)
    f.write(as_unicode('list_file=%s'%config.get('list_file','')).encode('utf8')+os.linesep)
    f.write(as_unicode('base_file=%s'%config.get('base_file','')).encode('utf8')+os.linesep)
    f.write(as_unicode('add_file=%s'%config.get('add_file','')).encode('utf8')+os.linesep)
    f.close()



class main(object):
    def __init__(self):
        self.mywin = Tk()
        self.mywin.title(u'商品图片文件处理 V%s( By yelord@qq.com with Python)'%__version__)
        self.mywin.geometry("640x480")
        try:
            self.mywin.iconbitmap('logo.ico')
        except:
            pass

        self.curdir = StringVar()
        self.listfile = StringVar()
        self.chkvar = StringVar()
        self.chkvar.set('default')

        #随窗口变化，自动扩展
        self.mywin.grid_columnconfigure(0, weight=1)
        self.mywin.grid_rowconfigure(0, weight=1)

        # main frame
        self.mainF = Frame(self.mywin, padx=10, pady=10)
        self.mainF.grid(column=0, row=0, sticky=(N, W, E, S))
        # col 1,row 2 auto size
        self.mainF.grid_columnconfigure(1, weight=1)
        self.mainF.grid_rowconfigure(5, weight=1)

        # status frame
        self.statusF = Frame(self.mywin, borderwidth=2, relief="sunken")
        self.statusF.grid(column=0, row=1, sticky=(W, E, S))

        # main layout
        self.dir_select = Entry(self.mainF,  textvariable=self.curdir)
        self.dir_select.grid(row=1, column=1, columnspan=3, sticky=(W, E))
        self.basecb = Button(self.mainF, text=DICT['default'][0], command=self.Do_basecb)
        self.basecb.grid(row=1,column=0,sticky=W)

        self.list_file = Entry(self.mainF,  textvariable=self.listfile)
        self.list_file.grid(row=2, column=1, columnspan=3, sticky=(W, E))
        self.listcb = Button(self.mainF, text=DICT['default'][1], command=self.Do_listcb)
        self.listcb.grid(row=2,column=0,sticky=W)  #

        Label(self.mainF,text=u'处理信息>',foreground='red').grid(row=4, column=0, sticky=W)
        self.msg = Text(self.mainF,fg="white", bg="black",relief="sunken")
        self.msg.grid(row=5, column=0, padx=5, pady=5, columnspan=4, sticky=(N, W, E, S))

        self.chk = Checkbutton(self.mainF, text=u"文件查重追加模式", variable=self.chkvar,
                               command=self.chgmode, onvalue="mode", offvalue="default")
        self.chk.grid(row=3, column=0, sticky=W)

        self.rencb = Button(self.mainF, text=DICT['default'][2], command=self.DoRename)
        self.rencb.grid(row=3, column=2, padx=10, ipadx=40, sticky=E)

        self.cpcb = Button(self.mainF, text=DICT['default'][3], command=self.Do_cpcb)
        self.cpcb.grid(row=3, column=3, padx=10, ipadx=40, sticky=E)  #

        show_tips(self,u'''
        ==========================tips=====================
        默认：
        1.首先请点`图片文件所在目录`按钮选择图片文件主目录。
        2.如果要检查或者修改文件名，[ 把空格与()去掉，改为_ ]，则点`规范文件名`执行
        3.要拷贝图片，请先点`图片需求清单文件`按钮选择要拷贝的图片清单文件.
        4.清单文件支持.xls或.txt文件，要求xls第一列(A)放条码，txt则每行放一个条码。
        选中追加模式：
        1.分别选择`基础资料`及`追加资料` xls文件，点追加按钮。
              Version: %s  yelord@qq.com
        ===================================================
        '''%__version__)

        # status layout
        self.statusF.grid_columnconfigure(1, weight=1)
        self.proc = ttk.Progressbar(self.statusF, mode='determinate')  #indeterminate , determinate
        self.proc.grid(row=1,column=1,sticky=(W,E,S,N))

        self.curdir.set(config.get('dir_select',''))
        self.listfile.set(config.get('list_file',''))

        self.dir_select.focus()

        self.mywin.mainloop()


    def chgmode(self):
        chk = self.chkvar.get()

        if chk =='mode':
            self.curdir.set(config.get('base_file',''))
            self.listfile.set(config.get('add_file',''))
            self.rencb['state'] = DISABLED
            show_tips(self,u'==== 已切换到 追加 模式 =====')
        else:
            self.curdir.set(config.get('dir_select',''))
            self.listfile.set(config.get('list_file',''))
            self.rencb['state'] = NORMAL
            show_tips(self,u'==== 已切换到 默认 模式 =====')
        # mode & default
        self.basecb['text'] = DICT[chk][0]
        self.cpcb['text'] = DICT[chk][3]
        self.rencb['text'] = DICT[chk][2]
        self.listcb['text'] = DICT[chk][1]


    def Do_basecb(self):
        chk = self.chkvar.get()
        if chk =='mode':
            self.GetBaseFile()
        else:
            self.GetCurrentDir()

    def Do_listcb(self):
        chk = self.chkvar.get()
        if chk =='mode':
            self.GetAddFile()
        else:
            self.GetListFile()

    def Do_cpcb(self):
        chk = self.chkvar.get()
        if chk =='mode':
            self.DO_Add()
        else:
            self.CopyFileTo()


    def RenFilename(self,parent,Ofn,rules=None):
        '''
        parent:文件路径
        ofn:原文件名
        rules：改名规则  ( '([0-9]*)[\s(]*(\d)[\s)]', r'\1_\2' )
        os.rename
        '''
        if not rules:
            rules = ( r'([0-9]*)[\s(]*(\d)[\s)]', r'\1_\2' )

        Ofn = as_unicode(Ofn)
        parent = as_unicode(parent)

        fn = re.sub(rules[0],rules[1],Ofn)

        if fn == Ofn :
            show_tips(self,u'文件 %s pass' % (Ofn))
            return

        Ofn = os.path.join(parent,Ofn)
        #TODO
        try:
            os.rename(Ofn, os.path.join(parent,fn))
            tips = u'文件 %s 改名为 %s ' % (Ofn,fn)
        except:
            tips = u'==== ！！！！文件 %s 改名失败！！！！====' % (Ofn)

        show_tips(self,tips)
        # log.info( tips )

    def DoRename(self):

        selectdir =  self.curdir.get()
        if not selectdir:
            show_tips(self,u'!!!请先选择路径!!!')
            return

        #save config
        setConfig()

        show_tips(self,u'''
                ---------   开始  --------
                ''')

        try:
            show_tips(self,selectdir + u'的文件处理中，请稍候...')
            # 总文件数
            _sum = sum([len(files) for root,dirs,files in os.walk(selectdir)])
            self.proc.configure(maximum=_sum)
            self.proc.start()
            _cnt = 0
            for  parent, dirnames, filenames in os.walk(selectdir):
                for fn in filenames:    #enumerate
                    self.RenFilename(parent,fn)
                    _cnt+=1

                self.proc.step(_cnt)
            self.proc.stop()
            tips = selectdir + u'目录下文件改名完成。\n----共修改了' + str(_sum) + u'个文件-----'
            show_tips(self,tips)
            # log.info(tips)
        except IOError:
            tips = u'文件处理发生错误'
            show_tips(self,tips)
            # log.error(tips)


     #选择图片目录
    def GetCurrentDir(self):
        cur_dir = tkFileDialog.askdirectory(title=u'选择图片文件所在的主目录')
        if cur_dir:
            cur_dir = as_unicode(cur_dir)
            self.curdir.set(cur_dir)
            config['dir_select'] = cur_dir


    # 选择清单文件
    def GetListFile(self):
        filename = tkFileDialog.askopenfilename(filetypes=[(u"xls文件","*.xls"),(u"txt文件","*.txt")],title=u'请选择清单文件')
        if filename:
            filename = as_unicode(filename)
            self.listfile.set(filename)
            config['list_file'] = filename

    # 选择基础文件
    def GetBaseFile(self):
        filename = tkFileDialog.askopenfilename(filetypes=[(u"xls文件","*.xls")],title=u'请选择基础文档文件')
        if filename:
            filename = as_unicode(filename)
            self.curdir.set(filename)
            config['base_file'] = filename

    # 选择追加文件
    def GetAddFile(self):
        filename = tkFileDialog.askopenfilename(filetypes=[(u"xls文件","*.xls")],title=u'请选择追加文档文件')
        if filename:
            filename = as_unicode(filename)
            self.listfile.set(filename)
            config['add_file'] = filename

    # 执行追加
    def DO_Add(self):
        basefile = as_unicode(self.curdir.get())
        if not basefile:
            show_tips(self,u'!!!请选择文件!!!')
            return

        addfile = as_unicode(self.listfile.get())
        if not addfile:
            show_tips(self,u'!!!请选择追加文件!!!')
            return

        #save config
        setConfig()
        self.proc.start()
        AddContentFromXls(basefile,addfile,self.show_tips)
        self.proc.stop()

    # 拷贝文件
    def CopyFileTo(self):

        filename = self.listfile.get()
        if not filename:
            show_tips(self,u'!!!请先选择清单文件!!!')
            return
        # 图片目录
        cur_dir = self.curdir.get()
        if not cur_dir:
            show_tips(self,u'!!!请先选择路径!!!')
            return

        #save config
        setConfig()

        curfiles=[]
        for parent,dirs,files in os.walk(cur_dir):
            parent = as_unicode(parent)
            for f in files:
                f = as_unicode(f)
                curfiles.append(os.path.join(parent,f))

        if len(curfiles) == 0:
            show_tips(self,u'!!!所选图片目录无文件!!!')
            return

        # copy to
        to_dir = tkFileDialog.askdirectory(title=u'文件将拷贝到的目录')
        if not to_dir:
            return   #取消

        bar_cnt = [0,0]
        fn_cnt = 0
        to_dir = as_unicode(to_dir)
        self.proc.start()
        f_yes = open(os.path.join(basedir,'copied.txt'),'w')
        f_yes.write(u'-----已拷贝文件------'.encode('gbk') + os.linesep)
        f_no = open(os.path.join(basedir,'not_copied.txt'),'w')
        f_no.write(u'------未匹配条码------'.encode('gbk') + os.linesep)

        requsts = ReadData(filename)  #unicode
        for barcode in requsts:
            ok = False
            if not barcode: continue
            bar_cnt[0]+=1
            for fn in curfiles:
                fn = as_unicode(fn)
                f = os.path.basename(fn)
                try:
                    if re.match(barcode+r'.*\.\w{2,4}$',f):
                        #TODO
                        shutil.copyfile(fn,os.path.join(to_dir,f))
                        show_tips(self,f + u' 拷贝成功')
                        ok = True
                        fn_cnt+=1   #拷贝文件数
                        f_yes.write(fn.encode('gbk') + os.linesep)
                except:
                    show_tips(self,u'======= ' + f + u' 拷贝失败 =========')

            if not ok:
                show_tips(self,barcode + u' 没匹配上')
                bar_cnt[1]+=1  # 没匹配条码数
                f_no.write(barcode.encode('gbk') + os.linesep )

        f_no.close()
        f_yes.close()
        self.proc.stop()
        tips = u'''
                     !!!! 清单里的文件拷贝完成 !!!!
                ----      共拷贝了 %s 个文件  见 copied.txt  -----
                ----    清单共有 %s 个单品(条码)     -----
                -----其中有 %s 个单品（条码）未匹配上 见 not_copied.txt ------
                ''' % (fn_cnt,bar_cnt[0],bar_cnt[1])

        show_tips(self,tips)
        # log.info(tips)

    def show_tips(self,msg, refresh=True):
        show_tips(self,msg,refresh)


def show_tips(root,msg,refresh=True):
    root.msg.insert('@0,0' ,msg + os.linesep)
    if refresh:
        root.mywin.update()
    
if __name__ == '__main__':

    root = main()

