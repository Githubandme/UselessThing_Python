#coding=utf-8
from selenium import webdriver
import time
#import getopt
import sys
import shutil
import os
import xlrd
from datetime import datetime
import re
import urllib.request
import requests


def CompareString(s1,s2,samePercent=0.8):
    if(s1==s2):
        return True
    elif(s1 in s2):
        return True
    elif(s2 in s1):
        return True
    else:
        sameCount1=0
        for i in s1:
            if(i in s2):
                sameCount1+=1
        if(sameCount1/len(s1)>samePercent):
            return True

        sameCount2 = 0
        for i in s2:
            if (i in s1):
                sameCount2 += 1
        if (sameCount2 / len(s2) > samePercent):
            return True
    return False


def GetNamesNeedToDownLoad():
    nameArr=[]
    excel=xlrd.open_workbook(r"C:\Users\Administrator\Desktop\Auto\投标项目时间安排.xlsx")
    table=excel.sheets()[0]

    rowCount=table.nrows
    colCount=table.ncols
    print(str(rowCount)+'行')
    allArr=[]
    for i in range(rowCount):
        rowArr=[]
        timeStr=table.cell_value(i,2)
        #print(timeStr)
        if(timeStr=='开标时间'):
            continue

        month=int(re.findall(r'年(.*?)月',timeStr)[0])
        day=int( re.findall(r'月(.*?)日',timeStr)[0])
        hour= int(re.findall(r' (.*?):',timeStr)[0])
        minute= int(re.findall(r':(.*?)$',timeStr)[0])
        dt= datetime(2019,month,day,hour,minute)

        if(datetime.now()>dt):#时间判断-------------------------------------------------------
            #print(dt)
            continue

        for j in range(colCount):
            v= table.cell_value(i, j)
            #print(v)
            rowArr.append(str(v))
        allArr.append(rowArr)

    # for a in allArr:
    #     if('日' in str(a)):
    #         print(a)
    for row in allArr:
        if (row[3] != '桐庐小额电子标'):
            #print(row[3])
            continue

        excelName=row[1]
        sameName=False
        for fileName in fileList:#字符串判断 + 表格索引
            if(fileName.strip()=='' or str(excelName).strip()=='' or excelName==[]):
                print(2)
                continue

            #print('-----------------------')
            if (CompareString(fileName,excelName)):
                sameName=True
                #print(excelName)
                break



        if(not sameName):
            #nameArr.append(excelName)
            nameArr.append(row)

    return nameArr



def readFilename(path, allfile):
    filelist = os.listdir(path)

    for filename in filelist:
        filepath = os.path.join(path, filename)
        if os.path.isdir(filepath):
            readFilename(filepath, allfile)
        else:
            allfile.append(filepath)
    return allfile



def AddDownloadLink(url,nameArr):
    print('添加此文件链接：'+nameArr[1])
    b=True
    browser2 = webdriver.Firefox()

    browser2.implicitly_wait(10)
    browser2.get(url)


    js2 = "var q2=document.documentElement.scrollTop=1000"
    browser2.execute_script(js2)
    rows= browser2.find_elements_by_class_name("Row")

    trs = browser2.find_elements_by_class_name("MsoNormal")
    project_type=trs[len(trs)-4].text#人员要求
    project_type=str(project_type).replace('?','')

    for r in rows:
        if(".招标书" in r.text):
            url=r.find_element_by_link_text("文件下载").get_attribute("href")


            downLinks.append(url)
            fileNameList.append(nameArr[2]+'--------'+project_type+'----------'+nameArr[1])
            print('添加链接'+str(url))
            browser2.quit()
            return

    time.sleep(2)




def DownLoadByPage(first):

    js = "var q=document.documentElement.scrollTop=200"
    browser.execute_script(js)
    print('下一页'+str(first))
    if(first==False):
        browser.find_element_by_link_text('【后页】').click()
    browser.implicitly_wait(10)
    nameList = browser.find_elements_by_class_name('BulletinTitleEnded')
    for name in nameList:

        #time.sleep(10)#
        browser.implicitly_wait(10)

        url = name.get_attribute("href")
        for n in needDownNames:
            if(CompareString(n[1],name.text)):
                #name.click()  #
                AddDownloadLink(url,n)






if __name__ == "__main__":

    # oh='https://file.zhaobide.com/dservice/UpLoadFiles/ProAfficheWord/132006333939091989.%e6%8b%9b%e6%a0%87%e4%b9%a6'
    # re=requests.get(oh)
    # f=open(r'C:\Users\Administrator\Desktop\1.招标书','wb')
    # f.write(re.content)




    #time.sleep(100)
    # downloadPath = r'C:\Users\Administrator\Desktop\TODO'
    # temp=r'https://file.zhaobide.com/dservice/UpLoadFiles/ProAfficheWord/131987347909678566.%e6%8b%9b%e6%a0%87%e4%b9%a6'
    # name_t='城西区块拆迁指挥部业务用房提升改造工程'
    # os.makedirs(downloadPath+"\\"+name_t)
    #os.makedirs(r'C:\Users\Administrator\Desktop\TODO\2019年4月23日 9:00横村镇白水路与横政路交叉处两侧临时围墙工程')
    # urllib.request.urlretrieve(temp, downloadPath+"\\"+name_t+"\\"+name_t+r'.招标文件')  # 下载图片。

    path =r"C:\Users\Administrator\Desktop\TODO"
    mainUrl="http://www.tlztb.com.cn/yj.aspx?xm=tlzbd&xj=tlzbd_jsgczbgg"
    fileList = os.listdir(path)

    needDownNames=GetNamesNeedToDownLoad()

    #print(needDownNames)
    # for f in fileList:
    #     print(f)
    downLinks = []
    fileNameList=[]

    #time.sleep(20)
    # for name in fileList:
    #     print(name)

    browser = webdriver.Firefox()

    browser.get(mainUrl)

    browser.implicitly_wait(20)

    js = "var q=document.documentElement.scrollTop=200"
    browser.execute_script(js)
    try:

        browser.switch_to.frame('showright')
        time.sleep(1)
        browser.switch_to.frame('showList')
        time.sleep(1)
        #nameList = browser.find_elements_by_class_name('BulletinTitleEnded')
        DownLoadByPage(True)

        #DownLoadByPage(False)

        #DownLoadByPage(False)


        downloadPath=r'C:\Users\Administrator\Desktop\TODO'
        print('要下载的链接：')
        print(downLinks)
        print('要下载的项目名称：')
        print(fileNameList)
        for link,name in zip(downLinks,fileNameList):
            name = str(name).replace("\n", "")
            name = str(name).replace(":", "_")

            filepath=downloadPath+"\\"+name

            #filepath= filepath.replace("\n","")



            if (not os.path.exists(filepath)):
                print('start------------download')

                os.makedirs(filepath)

                re = requests.get(link)
                f=open(filepath+"\\"+name+".招标书",'wb')
                f.write(re.content)


                #urllib.request.urlretrieve(url, filepath+"\\"+name+".招标书")  # 下载图片。
                print(filepath)
                print(name)

                print('download--------success!')


    except AssertionError:
        print('failllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllll')
    finally:
        browser.close()
