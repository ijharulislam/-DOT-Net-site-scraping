import urllib
import urllib2
from datetime import datetime
import xlrd
from lxml import html
import requests
from bs4 import BeautifulSoup
import xlsxwriter
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import csv 
book = xlrd.open_workbook("tax_info_2015.xls")

print book.sheet_names()

secend_sheet = book.sheet_by_index(0)
print secend_sheet

data = []


try:
	from selenium import webdriver
	from fake_useragent import UserAgent
	ua = UserAgent()
	profile = webdriver.FirefoxProfile()
	profile.set_preference("general.useragent.override", ua.random)
	driver = webdriver.Firefox(profile)

	for i in range(50,2000):
		try:
			num_dict = {}
			o = secend_sheet.row_values(i)
			num = o[0]
			property_address = o[9]
			full_add = "%s, %s"%(property_address,"Pima County,AZ")
			print full_add
			nam_add1 = o[6]
			nam_add2 = o[7]

			name_add = "%s %s"%(nam_add1,nam_add2)
			driver.get("http://www.realtor.com/")
			elem = driver.find_element_by_id("searchBox")
			elem.clear()
			elem.send_keys("%s"%full_add)
			time.sleep(8)
			elem.send_keys(Keys.RETURN)
			import time
			time.sleep(15)
			output = {}
			output["ParcelNumber"] = num
			url = driver.current_url
			if url =="http://www.realtor.com/":
				elem = driver.find_element_by_id("searchBox")
				elem.clear()
				elem.send_keys("%s"%name_add)
				time.sleep(8)
				elem.send_keys(Keys.RETURN)
				time.sleep(15)
				print name_add
				url = driver.current_url
				if url =="http://www.realtor.com/":
					continue
				else:
					su = BeautifulSoup(driver.page_source,"lxml")
					try:
						price = su.find("span",attrs={"itemprop":"price"}).text
						output["Price"] = price
						city = su.find("span",attrs={"itemprop":"addressLocality"}).text
						output["City"] = city
						state = su.find("span",attrs={"itemprop":"addressRegion"}).text
						output["State"] = state
						postal_code = su.find("span",attrs={"itemprop":"postalCode"}).text
						output["Postal Code"] = postal_code
						output["Search Address"] = name_add
						output["Realtor Url"] = url
					except:
						continue
			else:
				try:
					url = driver.current_url
					suu = BeautifulSoup(driver.page_source,"lxml")
					price = suu.find("span",attrs={"itemprop":"price"}).text
					output["Price"] = price
					city = suu.find("span",attrs={"itemprop":"addressLocality"}).text
					output["City"] = city
					state = suu.find("span",attrs={"itemprop":"addressRegion"}).text
					output["State"] = state
					postal_code = suu.find("span",attrs={"itemprop":"postalCode"}).text
					output["Postal Code"] = postal_code
					output["Realtor Url"] = url
					output["Search Address"] = full_add
					# url = driver.current_url
					# print url
					# output["Realtor Url"] = url
				except:
					title_error = suu.find("h3",class_="title-error")
					print "PRint Title Error"
					print title_error
					if title_error is not None:
						driver.back()
						driver.get("http://www.realtor.com/")
						elem = driver.find_element_by_id("searchBox")
						elem.clear()
						elem.send_keys("%s"%name_add)
						time.sleep(8)
						elem.send_keys(Keys.RETURN)
						time.sleep(15)
						url = driver.current_url
						try:
							soo = BeautifulSoup(driver.page_source,"lxml")
							output["Search Address"] = name_add
							price = soo.find("span",attrs={"itemprop":"price"}).text
							output["Price"] = price
							city = soo.find("span",attrs={"itemprop":"addressLocality"}).text
							output["City"] = city
							state = soo.find("span",attrs={"itemprop":"addressRegion"}).text
							output["State"] = state
							postal_code = soo.find("span",attrs={"itemprop":"postalCode"}).text
							output["Postal Code"] = postal_code
							output["Realtor Url"] = url
						except Exception, e:
							print e
							continue
					else:
						driver.back()
						driver.get("http://www.realtor.com/")
						elem = driver.find_element_by_id("searchBox")
						elem.clear()
						elem.send_keys("%s"%name_add)
						time.sleep(8)
						elem.send_keys(Keys.RETURN)
						time.sleep(15)
						url = driver.current_url
						print name_add
						if url == "http://www.realtor.com/":
							continue
						else:
							try:
								sop = BeautifulSoup(driver.page_source,"lxml")
								output["Search Address"] = name_add
								price = sop.find("span",attrs={"itemprop":"price"}).text
								output["Price"] = price
								city = sop.find("span",attrs={"itemprop":"addressLocality"}).text
								output["City"] = city
								state = sop.find("span",attrs={"itemprop":"addressRegion"}).text
								output["State"] = state
								postal_code = sop.find("span",attrs={"itemprop":"postalCode"}).text
								output["Postal Code"] = postal_code
								url = driver.current_url
								output["Realtor Url"] = url
								
							except Exception, e:
								print e
								continue
			data.append(output)
			print output
			for i in range(1,1000,50):
				if len(data) == i:
					print "it is !!!!!! %s"%i
					def WriteDictToCSV(csv_columns,dict_data):
						f_name = "realtor1-%s.csv"%num
						with open(f_name, 'w') as csvfile:
							writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
							writer.writeheader()
							for row in dict_data:
								writer.writerow(row)
					csv_columns =['ParcelNumber','Search Address','Price','City', 'State', 'Postal Code',"Realtor Url"]
					WriteDictToCSV(csv_columns,data)
					driver.close()
					time.sleep(10)


		except Exception, e:
			print e
			continue

except Exception, e:
	print e
	pass


finally:
	def WriteDictToCSV(csv_columns,dict_data):
		with open("realtor1.csv", 'w') as csvfile:
			writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
			writer.writeheader()
			for row in dict_data:
				writer.writerow(row)
	csv_columns =['ParcelNumber','Search Address','Price','City', 'State', 'Postal Code',"Realtor Url"]
	WriteDictToCSV(csv_columns,data)
	driver.close()



