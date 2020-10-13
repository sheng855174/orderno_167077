from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import datetime
import json
import os
import xlwt
import pyautogui
from shutil import copyfile


def get_htmltext(username, password, fb_page,下滑次數,下載等待時間):
	driver = webdriver.Firefox()
	driver.get("http://www.facebook.com")
	time.sleep(3)
	driver.find_element_by_id("email").send_keys(username)
	driver.find_element_by_id("pass").send_keys(password)
	driver.find_element_by_id("u_0_b").click()
	time.sleep(3)
	driver.get(fb_page)
	time.sleep(3)
	for i in range(下滑次數):
		y = 4000 * (i + 1)
		driver.execute_script(f"window.scrollTo(0, {y})")
		time.sleep(2)
	pyautogui.hotkey('ctrl', 's')
	time.sleep(5)
	pyautogui.typewrite("webpage.html")
	time.sleep(5)
	pyautogui.hotkey('enter')
	time.sleep(下載等待時間)
	#driver.close()


def parse_htmltext(下載目錄路徑,群組id):
	filename = "\\webpage.html"
	f = open(下載目錄路徑+filename, "r",encoding="utf-8")
	htmltext = f.read()
	post_persons = []
	comment_persons = []
	good_urllist = []
	soup = BeautifulSoup(htmltext, 'html.parser')
	body = soup.find('body')
	posts = body.select('div[data-pagelet="GroupFeed"]')[0]
	feed_articles = posts.select('div[role="feed"]')[0].select('div[role="article"]')
	articles = feed_articles
	
	#print("總共爬到文章數量 : " , len(articles))
	
	result = ""
	
	excel_file = xlwt.Workbook()
	sheet = excel_file.add_sheet("爬蟲結果")
	row=0
	col=0
	
	for article in articles:
		try:
			col=0
			check = True
			link = False
			貼文時間 = re.findall('<div.[^>]*aria-label="(.[^"]*月.[^"]*)".[^>]*role="button" tabindex="0".[^>]*>', str(article))
			貼文連結 = re.findall('<a.[^>]*href="(.[^"]*permalink.[^"]*)".[^>]*role="link" tabindex="0".[^>]*>', str(article))
			貼文內容 = re.findall('#代購.[^#]*#[^#]*<a.[^#]*role="link" tabindex="0">#(.[^<]*)<\/a>', str(article))
			簡化連結 = ""
			if len(貼文時間) == 0 :
				貼文時間 = re.findall('<div.[^>]*aria-label="(.[^"]*天.[^"]*)".[^>]*role="button" tabindex="0".[^>]*>', str(article))
			if len(貼文時間) == 0 :
				貼文時間 = re.findall('<div.[^>]*aria-label="(.[^"]*小時)".[^>]*role="button" tabindex="0".[^>]*>', str(article))
			if len(貼文時間) == 0 :
				貼文時間 = re.findall('<div.[^>]*aria-label="(.[^"]*分鐘)".[^>]*role="button" tabindex="0".[^>]*>', str(article))
			if len(貼文連結) == 0 :
				貼文連結 = re.findall('<a aria-label="圖像.[^"]*".[^>]*href=".[^"]*pcb\.(.[0-9a-zA-Z]*).[^"]*".[^>]*>', str(article))
				貼文連結 = "https://www.facebook.com/groups/"+群組id+"/permalink/"+str(貼文連結[0])+"/"
				link = True
			if len(貼文內容) == 0 :
				貼文內容 = re.findall('#代購.[^#]*#[^#]*<a.[^＃]*role="link" tabindex="0">＃(.[^<]*)<\/a>', str(article))
			
			if len(貼文時間) == 0 :
				print("沒找到貼文時間")
				check = False
			else :
				貼文時間 = str(貼文時間[0])
				print("貼文時間 : " + 貼文時間)
			if len(貼文連結) == 0 :
				print("沒找到貼文連結")
				check = False
			else :
				if link == False:
					貼文連結 = str(貼文連結[0])
					簡化連結 = 貼文連結[0:貼文連結.find("?")]
				print("貼文連結 : " + 貼文連結)
			if len(貼文內容) == 0 :
				print("沒找到貼文內容")
				check = False
			else :
				貼文內容 = str(貼文內容[0])
				print("貼文內容 : " + 貼文內容)
			print("=============================")
			
			if check == True:
				sheet.write(row, col, 貼文時間)
				col = col + 1
				sheet.write(row, col, xlwt.Formula('HYPERLINK("' + 簡化連結 + '";"' + 貼文內容 + '")'))
				col = col + 1
				sheet.write(row, col, 貼文連結)
				col = col + 1
				row = row + 1
		except Exception as e:
			#print(e)
			continue
	
	f.close()
	copyfile(下載目錄路徑+filename, "." + filename)
	os.remove(下載目錄路徑+filename)
	excel_file.save('fb_result.xls')
	
	
	


if __name__ == '__main__':
	file = open("config.txt", "r",encoding="utf-8")
	config = json.loads(file.read())
	file.close()
	print("config.txt 讀取完成.")

	username = config["帳號"]
	password = config["密碼"]
	fb_page = config["群組專頁"]
	群組id = config["群組id"]
	下滑次數 = config["下滑次數"]
	下載等待時間 = config["下載等待時間"]
	下載目錄路徑 = config["下載目錄路徑"]

	get_htmltext(username, password, fb_page,下滑次數,下載等待時間)
	parse_htmltext(下載目錄路徑,群組id)
	print("執行完畢")
	os.system("pause")
