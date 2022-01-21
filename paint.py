# -*- coding: utf-8 -*-
"""
Created on Thu Jan 20 18:14:27 2022

@author: LENOVO
"""

from PIL import Image
from xlsxwriter import Workbook

class ExcelPic(object):
    FORMAT_CONSTRTANT = 65536   # ?
   
    # 初始化
    def __init__(self, pic_file, ratio= 0.5):
        self._pic_file= pic_file
        self._ratio= ratio
        
        self._zoomed_out= False #???
        
        self._formats= dict()
        
    # 缩小
    def zoom_out(self, _imp):
        # _size= _imp.size
        _imp.thumbnail((int(_imp.size[0]*self._ratio), int(_imp.size[1]*self._ratio)))
        self._zoomed_out= True
        # ???
        
    # 颜色圆整，规整
    def round_rgb(self,rgb, model):
        return tuple([int(round(x/model)*model)for x in rgb])
    
    # 颜色样式，去重
    def get_format(self, color):
        _format= self._formats.get(color,None)
        
        if _format is None:
            _format=self.__wb.add_format({'bg_color':color})
            self._formats[color]= _format
        
        return _format
    
    # 操作流程
    def procese(self,output_file='_pic.xlsx',color_rounding=False, color_rounding_model=5.0):
        # 定义 初始属性
        self.__wb= Workbook(output_file)
        self.__sht = self.__wb.add_worksheet()
        self.__sht.set_default_row(height=9)
        self.__sht.set_column(0,5000,width=1) # 0-5000列
        # 打开图片
        _img= Image.open(self._pic_file)
        print('picture filename', self._pic_file)
        # 是否缩小  
        #???
        if self._ratio <1:
            self.zoom_out(_img)
            
        # 遍历像素点，填充对应单元格颜色
        _size= _img.size
        print('picture size:',_size)
        for (x,y) in [(x,y) for x in range(_size[0]) for y in range(_size[1])]:
            _clr= _img.getpixel((x,y))  # 二维数组
            # 减少颜色种类
            if color_rounding:
                # ??? 默认是false,主程序传入是true
                _clr= self.round_rgb(_clr,color_rounding_model)
                
             # '''这是在干嘛'''   
            _color= '#%02X%02X%02X' % _clr 
            self.__sht.write(y,x,'',self.get_format(_color))
        
        self.__wb.close()
    
        


if __name__ == '__main__':
    r= ExcelPic('0.jpg',ratio=0.5)
    r.procese('0.xlsx',color_rounding=True, color_rounding_model=5.0)