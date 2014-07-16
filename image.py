# -*- coding: utf-8 -*- 

__author__ = 'yelord'

from PIL import Image

# 640 * 640
MINPIXEL = 640

def img_resize(fn_source,fn_taget):

    im = Image.open(fn_source)

    # (3456, 2304) => (960,640)
    ratio = im.size[0]*1.000 / im.size[1]

    if im.size[0] < MINPIXEL or im.size[1] < MINPIXEL :
        return

    if ratio > 1:  # w > h
        h = MINPIXEL
        w = int(h * ratio)   #960
        x = (w - MINPIXEL)/2
        y = 0
    else:
        w = MINPIXEL
        h = int(w / ratio)
        y = (h - MINPIXEL)/2
        x = 0

    im = im.resize ( (w, h) )

    # 四元组(左，上，右，下)
    box = (x,y,MINPIXEL + x,MINPIXEL + y)
    im = im.crop(box)

    #im.show()
    im.save(fn_taget,im.format)


if __name__ == '__main__':
    import os
    img_resize(u'/Users/yelord/PycharmProjects/sku_info/demo图片/6901009600328_1.JPG','test.jpg')
    print os.path.splitext(u'/Users/yelord/PycharmProjects/sku_info/demo图片/6901009600328_1.JPG')
