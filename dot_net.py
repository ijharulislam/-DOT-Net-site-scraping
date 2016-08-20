
import urllib
import urllib2


from datetime import datetime

import xlrd
from lxml import html
import requests
from bs4 import BeautifulSoup
import xlsxwriter
import re

import csv 
book = xlrd.open_workbook("AZ_Pima_to_scrub.xls")

print book.sheet_names()

secend_sheet = book.sheet_by_index(1)
print secend_sheet

number_list =[]
for i in range(2,10986):
        o = secend_sheet.row_values(i)
        percel_number = o[2]
        number_list.append(o[2])

print number_list




data = []
try:
  for k in number_list:
    output = {}
    uri = 'http://www.to.pima.gov/pcto/tweb/property_inquiry'


    headers = {
        'HTTP_USER_AGENT': 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.13) Gecko/2009073022 Firefox/3.0.13',
        'HTTP_ACCEPT': 'text/html,application/xhtml+xml,application/xml; q=0.9,*/*; q=0.8',
        'Content-Type': 'application/x-www-form-urlencoded'
    }
     
    encodedFields = urllib.urlencode({'statecode':k, 'taxyear':2015, 'date':'08/19/2016','submit': "SEARCH"})

    print encodedFields


    enc = "statecode=101013930&taxyear=2015&date=08%2F18%2F2016&submit=SEARCH"
    try:
      req = urllib2.Request(uri, encodedFields, headers)
      f= urllib2.urlopen(req)
      content = f.read()
      su = BeautifulSoup(content, "lxml")
    except:
      import time
      time.sleep(10)
      continue
    # print su 
    percel_num = k 
    output["ParcelNumber"] = percel_num
    output["Tax Year"] = 2015
    total_due = su.find("th", text="TOTAL DUE:")
    if total_due is not None:
      total_due = su.find("th", text="TOTAL DUE:").findNext('td').text
    output["FY 2015 DUE"] = total_due
    try:
      property_type = su.find("th", text="PROPERTY TYPE:").findNext('td').text
      output["PROPERTY TYPE"] = str(property_type).strip()
    except:
      output["PROPERTY TYPE"] = ""
    try:
      tax_area = su.find("th", text="TAX AREA:").findNext('td').text
      output["TAX AREA"] = tax_area.strip()
    except:
      output["TAX AREA"] = ""
    try:
      tax_payer_add = su.find("th", text=re.compile(r'TAXPAYER')).parent.findNext("tr").find("td")
      t = str(tax_payer_add)
      ts = t.replace('<td class="color-light" style="width: 100%;">',"").replace("</td>","")
      t_list = ts.split("<br/>")
      if len(t_list) == 3:
        output["NAMEADDRESS_LINE1"] = t_list[0].strip()
        output["NAMEADDRESS_LINE2"] = t_list[1]
        output["NAMEADDRESS_LINE3"] = t_list[2]
      elif len(t_list) == 4:
        output["NAMEADDRESS_LINE1"] = t_list[0].strip()
        output["NAMEADDRESS_LINE2"] = t_list[1]
        output["NAMEADDRESS_LINE3"] = t_list[2]
        output["NAMEADDRESS_LINE4"] = t_list[3]
      elif len(t_list) == 2:
        output["NAMEADDRESS_LINE1"] = t_list[0].strip()
        output["NAMEADDRESS_LINE2"] = t_list[1]
    except:
      pass

    try:
      property_add = su.find("th", text="PROPERTY ADDRESS").parent.findNext("tr").find("td").text
      output["PROPERTY ADDRESS"] = str(property_add).strip()
    except:
      output["PROPERTY ADDRESS"] = ""
    try:
      legal_description = su.find("th", text="LEGAL DESCRIPTION").parent.findNext("tr").find("td").text
      output["LEGAL DESCRIPTION"] = str(legal_description).strip()
    except:
      output["LEGAL DESCRIPTION"] = ""
    try:
      paid_by = su.find("th", text="PAID BY").parent.findNext("tr").find("td").text
      output["PAID BY"] = paid_by.strip()
    except:
      output["PAID BY"] = ""
    try:
      on_behalf = su.find("th", text="ON BEHALF OF").parent.findNext("tr").find("td").text
      output["ON BEHALF OF"] = on_behalf.strip()
    except:
      output["ON BEHALF OF"] = ""

    data.append(output)

    print output
except:
  pass

finally:
  def WriteDictToCSV(csv_columns,dict_data):
    with open("tax_info.csv", 'w') as csvfile:
      writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
      writer.writeheader()
      for row in dict_data:
        print row
        writer.writerow(row)


  csv_columns =['ParcelNumber', 'Tax Year','FY 2015 DUE','PROPERTY TYPE', 'TAX AREA','NAMEADDRESS_LINE1','NAMEADDRESS_LINE2','NAMEADDRESS_LINE3','NAMEADDRESS_LINE4', 'PROPERTY ADDRESS','LEGAL DESCRIPTION', 'PAID BY','ON BEHALF OF']

  WriteDictToCSV(csv_columns,data)






