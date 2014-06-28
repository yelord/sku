sku
===

Python,tk 私人练习之作。商品资料文档改文件名，拷贝文件，读写excel文件，查重等小工具

Tips
==================================================
        默认：
        1.首先请点`图片文件所在目录`按钮选择图片文件主目录。
        2.如果要检查或者修改文件名，[ 把空格与()去掉，改为_ ]，则点`规范文件名`执行
        3.要拷贝图片，请先点`图片需求清单文件`按钮选择要拷贝的图片清单文件.
        4.清单文件支持.xls或.txt文件，要求xls第一列(A)放条码，txt则每行放一个条码。
        选中追加模式：
        1.分别选择`基础资料`及`追加资料` xls文件，点追加按钮。
              Version: 0.1.2  yelord@qq.com
              
Run           
===================================================
        python sku.py
        
        or 
        
windows: build exe
------------------
pyinstaller 2.1
        
        python ..\pyinstaller.py  -F -w -i logo.ico sku.py
        
        -F: 生成单个文件
        -w: 有窗口界面，不显示控制台

参考资料
--------
* [tkinter] http://effbot.org/tkinterbook/
* [python-excel]  https://github.com/python-excel/
