import cv2,os
import numpy as np
from PIL import Image 
import pickle
import xlsxwriter
import xlrd
from xlutils.copy import copy
from datetime import datetime, date, time
import time
import os.path

save_path ='/home/pi/Desktop/Stud_proj/abc/excel/'
ndate =((str(datetime.now())).split(' '))[0]
filename = "Attendance" + ndate+ ".xls"
file_name = os.path.join(save_path, filename)
file_name = str(file_name)
print file_name
x=os.path.isfile(file_name)
if x==True:
    print("file arready available")
    
    
    
else:
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold':1})
    center = workbook.add_format({'align':'center'})
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    time_format = workbook.add_format({'num_format': 'hh:mm:ss AM/PM'})
    worksheet.set_column(1,15)
    worksheet.set_column('B1:C1',30)
    worksheet.set_column('D1:E1',15)
    worksheet.set_column('E1:F1',15)
    
        

    worksheet.write('A1','Sr.No.')
    worksheet.write('A2','1')
    worksheet.write('A3','2')
    worksheet.write('A4','3')
    worksheet.write('A5','4')
    worksheet.write('A6','5')
    worksheet.write('A7','6')
    worksheet.write('A8','7')
    worksheet.write('A9','8')
    worksheet.write('A10','9')
    worksheet.write('A11','10')
    worksheet.write('B1','student name')
    worksheet.write('B2','mani')
    worksheet.write('B3','Abhishek')
    worksheet.write('B4','neha')
    worksheet.write('B5','mansi')
    worksheet.write('B6','datta')
    worksheet.write('B7','chandu')
    worksheet.write('B8','akshay')
    worksheet.write('B9','rohit')
    worksheet.write('B10','dhanush')
    worksheet.write('B11','yogesh')
    worksheet.write('C1','Attendance status')
    worksheet.write('C2','ABSENT')
    worksheet.write('C3','ABSENT')
    worksheet.write('C4','ABSENT')
    worksheet.write('C5','ABSENT')
    worksheet.write('C6','ABSENT')
    worksheet.write('C7','ABSENT')
    worksheet.write('C8','ABSENT')
    worksheet.write('C9','ABSENT')
    worksheet.write('C10','ABSENT')
    worksheet.write('C11','ABSENT')

    
                
    workbook.close()



recognizer = cv2.face.LBPHFaceRecognizer_create()
recognizer.read('/home/pi/Desktop/Stud_proj/abc/trainer/trainer.yml')
cascadePath = "/home/pi/Desktop/Stud_proj/abc/Classifiers/face.xml"
faceCascade = cv2.CascadeClassifier(cascadePath);
path = '/home/pi/Desktop/Stud_proj/abc/dataSet/'

cam = cv2.VideoCapture(0)
font = cv2.FONT_HERSHEY_SIMPLEX #Creates a font
while True:
    ret, im =cam.read()
    gray=cv2.cvtColor(im,cv2.COLOR_BGR2GRAY)
    faces=faceCascade.detectMultiScale(gray, scaleFactor=1.2, minNeighbors=5, minSize=(100, 100), flags=cv2.CASCADE_SCALE_IMAGE)
    for(x,y,w,h) in faces:
        nbr_predicted, conf = recognizer.predict(gray[y:y+h,x:x+w])
        cv2.rectangle(im,(x-50,y-50),(x+w+50,y+h+50),(225,0,0),2)
        print(nbr_predicted)
        if(nbr_predicted==1):
             nbr_predicted='unknown'
             
        elif(nbr_predicted==2):
             nbr_predicted='unknown'
             
        elif(nbr_predicted==3):
             nbr_predicted='mani'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(1,2,"present")
             Twb.save(file_name)
        elif(nbr_predicted==4):
             nbr_predicted='abhishek'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(2,2,"present")
             Twb.save(file_name)
        elif(nbr_predicted==5):
             nbr_predicted='neha'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(3,2,"present")
             Twb.save(file_name)
        elif(nbr_predicted==6):
             nbr_predicted='mansi'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(4,2,"present")
             Twb.save(file_name)
        elif(nbr_predicted==7):
             nbr_predicted='datta'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(5,2,"present")
             Twb.save(file_name)
        elif(nbr_predicted==8):
             nbr_predicted='chandu'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(6,2,"present")
             Twb.save(file_name)
        elif(nbr_predicted==9):
             nbr_predicted='akshay'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(7,2,"present")
             Twb.save(file_name)
        elif(nbr_predicted==10):
             nbr_predicted='rohit'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(8,2,"present")
             Twb.save(file_name)    

        elif(nbr_predicted==11):
             nbr_predicted='dhanush'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(9,2,"present")
             Twb.save(file_name)
        elif(nbr_predicted==12):
             nbr_predicted='yogesh'
             Trb = xlrd.open_workbook(file_name)
             Twb = copy(Trb)
             Tw_sheet = Twb.get_sheet(0)
             Tw_sheet.write(10,2,"present")
             Twb.save(file_name)    
        else:
             nbr_predicted='unknown' 
        cv2.putText(im,str(nbr_predicted)+"--"+str(conf), (x,y+h),font,1, (0,255,0),1) #Draw the text
        cv2.imshow('im',im)
        cv2.waitKey(10)









