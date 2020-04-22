# -*- coding: utf-8 -*-
"""
Created on Tue Apr 21 19:26:28 2020

@author: hoang
"""

from selenium import webdriver
import time
import pandas as pd
""""""

""""""""""""""""""""""""""""""""""GLOBAL RESOURCES"""""""""""""""""""""""""""""""""""""""
####https://chromedriver.chromium.org/downloads just in case 



"""Đây là những tham số của cái code - như link của mof, file names etc. Anything that should be changed should be done here"""
mof_site_url = 'https://olt.mof.gov.vn/Portal.IU.CMS/Default.aspx#CourseAcceptCSDT'

excel_filename = "Danh sách dự thi OB online Miền HCM ngay 22-04-2020.xlsx "  #COPY CA CAI FILE (INCL. .xlsx part) vao trong cai ""

MADUTHI_COL_INDEX = 9 - 1  #in computing, things count from 0, so we need to offset 1 
                            #trong file xlsx, mã dự thi ở cột I = cột 9 nếu cột 1 tính là 1
                            
MADUTHI_OFFSET = 1  # tại cái row 1 empty, nên code nó sẽ bắt đầu từ dòng 2
                    #nhưng dòng 2 là header, nên offset phát nữa 
"""Technical parameters"""
DELAY_BETWEEN_TASK = 3 #seconds
PAGE_LOAD_DELAY = 3 #seconds
BATCH_SIZE = 7

PERSISTENT_MAX_ATTEMPTS = 1

"""""""""""""""""""""""""""""""""""""""CODE SECTION"""""""""""""""""""""""""""""""""""""""



def persistent_find(Driver,LINK_TEXT):
    """Try to find elem by LINK TEXT for 1 mins. give up after that"""
    load = 1
    count = 1
    while(load):
            try:
                loading = Driver.find_element_by_link_text(LINK_TEXT)
                load = 0
            except:
                print(f'{LINK_TEXT} not found...try again in 5s')
                time.sleep(5)
                count += 1
                if (count == PERSISTENT_MAX_ATTEMPTS):
                    load = 0
                    print('Gave this up.')
                pass
    return loading 

def collate_maduthi(excel_filepath):
    mainDF = pd.read_excel(excel_filepath)
    maduthiList = mainDF.iloc[:,MADUTHI_COL_INDEX][MADUTHI_OFFSET:]   # do not take the first element
    return maduthiList

def collect_max_20_files(mof_site_url,maduthiList):
    """Trong trường hợp overflow Chrome,t sẽ chỉ cho nó chạy 20 tab trong một lần login thôi"""

    driver = webdriver.Chrome()
    driver.get(mof_site_url)
    user_input = input('Proceed?: ')    #đợi login và vào cái trang kiểm duyệt j j đó
    
    assert(len(maduthiList) <= BATCH_SIZE ), "Nhiều mã dự thi qué"
    
    if user_input in ['y','Y','yes','YES',1]:
        
        main_tab = driver.window_handles[0]
        
        tab_index = 1
       
        for ma_du_thi in maduthiList:
            try:
                user_link = driver.find_element_by_link_text(ma_du_thi)
                user_link.click()
                driver.switch_to.window(driver.window_handles[tab_index])
                time.sleep(PAGE_LOAD_DELAY)
                dwnload_btn = persistent_find(driver,'Tải danh sách')
                dwnload_btn.click()
                close_btn = persistent_find(driver,'Đóng')
                close_btn.click()
                time.sleep(DELAY_BETWEEN_TASK)
                driver.switch_to_window(main_tab)
                tab_index += 1
                print(f'Task succedded for mã dự thi {ma_du_thi}')
            except:
                print(f'Task failed for mã dự thi {ma_du_thi}.')
    driver.close()
    driver.quit()


def main():
    
   
    
    #Getting all the variables
    maduthiList = collate_maduthi(excel_filename)
    num = len(maduthiList)
    assert num > 0, "Cái list náy should contain mã dự thi"
    
    #Setup the driver
    

    #do tasks in batch of BATCH_SIZE
    numOfBatch = num//BATCH_SIZE
    for batch_i in range(numOfBatch):
        collect_max_20_files(mof_site_url,maduthiList[batch_i*BATCH_SIZE:(batch_i+1)*BATCH_SIZE])
    
    #dealing with the remnants
    if num%BATCH_SIZE!=0 and num > BATCH_SIZE:
        collect_max_20_files(mof_site_url,maduthiList[numOfBatch*BATCH_SIZE:])
    print("Program finished.")
if __name__ == '__main__':
    main()
        