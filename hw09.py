# -*- coding: utf-8 -*-
"""
Created on Tue Dec 24 09:52:02 2019

@author: q122858
"""
import requests
from bs4 import BeautifulSoup  
import xlwt #將資料寫入 Excel 文件
 
file=xlwt.Workbook(encoding='utf-8',style_compression=0)#新建一?sheet  
sheet=file.add_sheet('Sheet')  

index=1 #共1頁
n=0 #計算總共有多少家店
list1 = [] 
list2 = [] 
list3 = [] 
i=1
name1 = []
cat = []
cat1 = []
price = []
phone = []
location = []
pic = []
www = []
location2 = []
web = []
point= []

for i in range(1,10,1):
    url = 'https://guide.michelin.com/tw/zh_TW/taipei-region/taipei/restaurants/page/'+str(i)
    #url = 'https://guide.michelin.com/tw/zh_TW/taipei-region/taipei/restaurants?lat=23.947673599999998&lon=120.93030399999999'
    html = requests.get(url)
    soup = BeautifulSoup(html.text,'html.parser') 
    
    items = soup.select('div.col-md-6.col-lg-6.col-xl-3')
    
    for item in items: 
        try:       
            n+=1
            print("\nn=",n)
        
            item__hero=item.select('div.card__menu-image a')[0]  #圖片名稱 div
            imgurl=item__hero['data-bg'].split("url(")[1] #圖片名稱
            print("圖片名稱:",imgurl.strip(')'))  #圖片名稱
            
            name=item.select('div.card__menu-image a')[0]  #店名
            name=name['aria-label'].split("Open ")[1]  #刪除跳列字元、濾除前後空白、刪除最後一個字元如(「‹、m、=」
            print("店名:",name)
            
            category=item.find("div",{"class":"card__menu-footer"}).text #分類、區、價格 div
            #category=category.strip()# 去除首尾空格 
            category=category.split()  #以「 ·」分割字串
            kind=category[0].strip()  #分類 (濾除前後空白)
            area=category[1].strip()  #地區
            #price=category[2].strip() #價格
            print("地區:",kind)
            print("分類:",area)
    #        listdata=[imgurl,name,kind,area]
    #        list1.append(listdata)        
            
        except: # 
            print("\n資料不存在!")
    
    for j in range(0,20,1):
        if(i==9):
            list2.append(soup.select('div.col-md-6.col-lg-6.col-xl-3')[j].h5.a['href'])
            break
        else:
            list2.append(soup.select('div.col-md-6.col-lg-6.col-xl-3')[j].h5.a['href'])
            
for j in range(0,161,1):
    list3.append('https://guide.michelin.com/'+ list2[j])
    
print("第"+str(i)+"頁,共有"+ str(len(items)) + "間")
print("\n全部總共有",n,"間")
i+=1

for link in list3:
    try:
        url = link
        html1 = requests.get(url)
        soup = BeautifulSoup(html1.text,'html.parser') 
        
        name1.append(soup.find('h2').text)
        
        cat.append(soup.find('li',class_="restaurant-details__heading-price").text.split())
        cat1.append(cat[len(cat)-1][-1])
    
        price.append(cat[len(cat)-1][0])
        
        phone.append(soup.find('span', class_='flex-fill').text)
        
        location.append(soup.select('body > main > div > div > div > div> section > div > div.col-md-6 > div > iframe ')[0]['src'])
        
        pic.append(soup.select(' body > main > div > div > div > div')[0].img['src'])
        
        www.append(soup.find('link')['href'])
        
        location2.append(soup.select('body > main > div > div > div > div > section > div > ul > li')[0].text)
        
        web.append(soup.select('body > main > div > div > div > div > section > div > div > div > div > a')[0]['href'])
        
        point.append(soup.find('p',class_='js-show-description-text').text.split()[0])
        
        listdata=[name1,cat1,price,phone,location,pic,www,location2,web,point]
        list1.append(listdata) 
    except: # 
        print("\n資料不存在!")
# excel 資料
listtitle=['餐廳名稱',"餐廳分類","價格","餐廳電話",'餐廳的經緯度','餐廳在米其林招牌相片','餐廳在米其林的網址','餐廳地址','餐廳自己的網址','米其林指南的觀點']
row=0
for item in listtitle: # 標題
    sheet.write(0,row,item)
    row+=1

row=1 
for item1 in list1[0]: #資料
    col=0 
    for item2 in item1:
        sheet.write(col+1,row-1,item2)
        col=col+1
    row=row+1
    
file.save('MiChiLinXls.xls')#存檔 
