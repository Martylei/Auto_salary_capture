# -*- coding: utf-8 -*-
"""
An excel capture program created by Marty
09.15.2020 created file

"""


from win32com.client import Dispatch, DispatchEx
import pythoncom
from PIL import ImageGrab, Image
import datetime
import xlwings as xw 
from shutil import copyfile





def start():
    screen_area = 'A1:X3'#area that you need to clip
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')+' : program start')
    
    
    for i in range(0,12):#the loop time you want to do with how many rows are there
        filename = r'C:/Users/Marty/Desktop/for_demo/files/demo'+str(i)+'.xls'
        target = r'C:/Users/Marty/Desktop/for_demo/files/demo'+str(i+1)+'.xls'
        # print(filename)
        copyfile(filename, target)
        cap(filename,screen_area,target)
    
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')+' : end of program') 



def delete(target):
    wb_delete = xw.Book(target)
    #wb_delete.app.calculation = 'automatic'
    sht = wb_delete.sheets[0]
    sht.range('A3:X3').api.Delete()
    wb_delete.save(target)
    wb_delete.close()

def cap(filename, screen_area, target):
    pythoncom.CoInitialize()  
    excel = DispatchEx("Excel.Application")  
    
    wb = excel.Workbooks.Open(filename)
    ws = wb.Worksheets('Sheet1')  
    ws.Range(screen_area).CopyPicture()
    ws.Paste(ws.Range(screen_area))
    
    name = ws.Range("C3").Value
    ID = ws.Range("B3").Value
    print(name)
    print(ID)
    
    excel.Selection.ShapeRange.Name = name
    ws.Shapes(name).Copy()
    img = ImageGrab.grabclipboard()  # Grap data from clipboard
    img_name = ID + "_" + name + ".PNG" #Image name 
    img.save(img_name)  #Svae image

    delete(target)#Delete the part you just clipped
  
    wb.Save()
    wb.Close()
if __name__ == "__main__":
    start() 