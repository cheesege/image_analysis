import os
import cv2
import numpy as np

file_amount = 0
folder_path = (r"C:\Users\eric\Downloads\1208") #資料夾路徑  #r是把字串變成不受斜線影響的生字串
os.chdir(folder_path)#執行位置改為資料夾
file_list = [] #儲存檔名


first_point_sample = (0,98)#起始座標(左上ㄘ)
width_sample = 256#寬度
height_sample = 10#高度
second_point_sample = (first_point_sample[0]+width_sample,first_point_sample[1]+height_sample)

first_point_blank = (0,148)#起始座標(左上ㄘ)
width_blank = 256#寬度
height_blank = 10#高度
second_point_blank = (first_point_blank[0]+width_blank,first_point_blank[1]+height_blank)

folder_list = os.listdir(folder_path) 
for name in folder_list:#用name遍歷folder_path中的資料
    #print(name)
    file_amount = file_amount + 1 #計算資料夾檔案數量
    file_list.append(name)#新增檔名


"""for times in range(1):
    img_name = file_list[times]
    #print(times)
    img_origin = cv2.imread(img_name,-1)
    img_frag_sample = img_origin[first_point_sample[1]:second_point_sample[1], first_point_sample[0]:second_point_sample[0]]#y0:y1, x0:x1
    img_frag_blank = img_origin[first_point_blank[1]:second_point_blank[1], first_point_blank[0]:second_point_blank[0]]#y0:y1, x0:x1
    siz_sample = img_frag_sample.shape
    siz_blank = img_frag_blank.shape
    summ =0
    
    print(img_name)
    print("Sample:")
    for yt in range(siz_sample[1]):
        summ =0
        for xt in range(siz_sample[0]):
            summ +=img_frag_sample[xt][yt]
        summ = summ/siz_sample[0]
        print(summ)
    print('\n')
    print("Blank:")
    for yt in range(siz_blank[1]):
        summ =0
        for xt in range(siz_blank[0]):
            summ +=img_frag_blank[xt][yt]
        summ = summ/siz_blank[0]
        print(summ)
    print('\n')     """
    
    
    
    
    
    

