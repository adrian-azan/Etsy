import xlsxwriter
import _csv
import os
import glob
import time
import shutil
import selenium.webdriver.support.ui as ui
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException 
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities


#Allows 'Remember Me' option to stay selected
options = webdriver.ChromeOptions();
options.add_argument("user-data-dir=C:\\Users\\Adrian\\AppData\\Local\\Google\\Chrome\\User Data\\Defualt");


months = ["January","February", "March", "April",\
          "May", "June", "July", "August","September",\
          "October", "November", "December" ]


class Sale:
    def __init__(self, date, description, order_number, type, fees):
        self.date = date
        self.description = description
        self.type = type 
        self.fees = fees
        self.order_number = order_number
        

def checkXpath(xpath):
    try:
        driver.find_element_by_xpath(str(xpath))
    except NoSuchElementException:
        return False
    return True

def checkLinkText(link):
    try:
        driver.find_element_by_partial_link_text(str(link))
    except NoSuchElementException:
        return False
    return True

def writeFee(sales, current_row, type):
    for sale in sales:
        if sale.type == type:
            worksheet.write_string(current_row, 0, sale.date)
            worksheet.write_string(current_row, 1, sale.type)
            worksheet.write_number(current_row, 2, sale.fees)
            worksheet.write_string(current_row, 3, sale.description)
            worksheet.write_string(current_row, 4, sale.order_number)
            
            current_row = current_row + 1
    return current_row + 1


def dateFormat(date):
    for i in range(0, 12):
        if date[:3] == months[i][:3]:
            day = date.split(' ')[1]
            return str(i+1) + "/" + str(day) + '/' + str(date[-2:])

def dateCut(date):
    temp = date.split(',')
    return temp[0] + " " + temp[1]


driver = webdriver.Chrome(chrome_options=options)
driver.get('https://www.etsy.com/your/bill?ref=seller-platform-mcnav')
driver.implicitly_wait(1)

username = "userName"
password = "password"

userBox = "//input[@name='username']"
passwordBox = "//input[@name = 'password']"

fees_element = "//span [@class='currency-value']"
dates_elements = "//td [@style='font-size:12px; font-weight:bold; padding:10px 10px;']"
description_elements = "//table [@style='vertical-align:middle;']/tbody/tr/td[2]"
types_elements = "//td [@style='font-size:12px; padding:10px 0px 10px 0px;']"

workbook = xlsxwriter.Workbook("Etsy.xlsx")
worksheet = workbook.add_worksheet()

products = []
current_row = 1
all_sales = []
for month in months:
    if checkLinkText(month):
        link = driver.find_element_by_partial_link_text(month)
        link.click()
        
        fees = driver.find_elements_by_xpath(fees_element)
        dates = driver.find_elements_by_xpath(dates_elements)
        descriptions = driver.find_elements_by_xpath(description_elements)
        types = driver.find_elements_by_xpath(types_elements)
        year = driver.find_element_by_xpath("//*[@id='your-shop-content']/div[2]/div/h2")
                
        for i in range(0, len(dates)):
            full_date = dates[i].text + " " + year.text[-2:]
            dates[i] = dateFormat(full_date)
        
        for j in range(0, len(fees)):
            sale = Sale(dates[j], descriptions[j].text, "", types[j].text, float(fees[j].text))
            all_sales.append(sale)

        driver.back()


driver.get('https://www.etsy.com/your/account/payments?ref=seller-platform-mcnav')


cells = driver.find_elements_by_xpath("//td [@class ='first first-two']")
sales = []
for i in range(0, len(cells)):
    check_sale = "//tbody/tr[" + str(i+1) + "]/td[2]/p[2]/span[@class='description']"
    if checkXpath(check_sale):
        vat_cost = 0
        vat_element = "//tbody/tr[" + str(i+1) + "]/td[2]/p[3]/span/span[@class='currency-value']"
        if checkXpath(vat_element):
            vat_cost = driver.find_element_by_xpath(vat_element)
            vat_cost = float(vat_cost.text)

        cost = driver.find_element_by_xpath("//tbody/tr[" + str(i+1) + "]/td[3]/span[2]")
        date = driver.find_element_by_xpath("//tbody/tr[" + str(i+1) + "]/td[1]")
        order_number = driver.find_element_by_xpath("//tbody/tr[" + str(i+1) + "]/td[2]/p[1]/span[1]/a")
        fee = driver.find_element_by_xpath("//tbody/tr[" + str(i+1) + "]/td[4]/span[2]")

        date = dateCut(date.text)
        date = dateFormat(date)
        cost = float(cost.text)
        cost = cost - float(vat_cost)

        orderNumber = order_number.text
        fee = float(fee.text)

        order_number.click()
        title = driver.find_element_by_xpath("//div [@class='item-details receipt-column']/h4/a")
        title = title.text
        driver.back()
        
        sale = Sale(date, title, orderNumber, "selling fee", fee)
        all_sales.append(sale)

        sale = Sale(date, title, orderNumber, 'sales', cost)
        sales.append(sale)
       
driver.close()

        
current_row = writeFee(all_sales, current_row, "listing")
print("Writing Listing Fees...")
current_row = writeFee(all_sales, current_row, "transaction")
print("Writing Transaction Fees...")
current_row = writeFee(all_sales, current_row, "auto-renew sold")
print("Writing Auto-Renew Fees...")
current_row = writeFee(all_sales, current_row, "selling fee")
print("Writing Selling Fees...")
current_row = writeFee(sales, current_row, "sales")
print("Writing Sales...")




worksheet.write_string('A1', "Date")
worksheet.write_string('B1', "Fee Type")
worksheet.write_string('C1', "Amount")
worksheet.write_string('D1', "Description")

worksheet.write_string('G2', "Listing Fee")
worksheet.write_formula('H2', '=SUMIFS(C:C,B:B,"=listing")')

worksheet.write_string('G3', "Transaction Fee")
worksheet.write_formula('H3', '=SUMIFS(C:C,B:B,"=transaction")')

worksheet.write_string('G4', "Auto-Renew Fees")
worksheet.write_formula('H4', '=SUMIFS(C:C,B:B,"=auto-renew sold")')

worksheet.write_string('G5', "Selling Fee")
worksheet.write_formula('H5', '=SUMIFS(C:C,B:B,"=selling fee")')

worksheet.write_string('G6', "Sales")
worksheet.write_formula('H6', '=SUMIFS(C:C,B:B,"=sales")')

worksheet.write_string('G7', "Gross Profit")
worksheet.write_formula('H7', '=H6-SUM(H2:H5)')

worksheet.write_string('G9', "Title")
worksheet.write_string('H9', "Sales")
worksheet.write_string('I9', "Fees")
worksheet.write_string('J9', "Gross")

fin = open("Products.txt", "r")
for line in fin:
    line = line[:-1]
    products.append(line)
    
for i in range(0, len(products)):
    temp = products[i].split(' ')
    title = temp[0] + ' ' + temp[1] + ' ' + temp[2]
    saleFormula = '=SUMIFS(C:C, D:D,"=' + products[i] + '", B:B, "=sales")'
    feeFormula = '=SUMIFS(C:C, D:D,"=' + products[i] + '", B:B, "<>sales")'
    grossFormula = '=H' + str(i+10) + '- I' + str(i+10)
    worksheet.write_string('G' + str(i+10), title)
    worksheet.write_formula('H' + str(i+10), saleFormula )
    worksheet.write_formula('I' + str(i+10), feeFormula )
    worksheet.write_formula('J' + str(i+10), grossFormula )

red = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'red'} )
green = workbook.add_format({'bold': True, 'font_color': 'green', 'bg_color': 'lime'})
worksheet.conditional_format('J10:J' + str(len(products)+ 9), {'type': 'cell', 'criteria': '>',\
                                    'value': 0, 'format': green})

worksheet.conditional_format('J10:J' + str(len(products)+ 9), {'type': 'cell', 'criteria': '<=',\
                                    'value': 0, 'format': red})
worksheet.set_column(0,0,11)
worksheet.set_column(1,1,15)
worksheet.set_column(2,2,8)
worksheet.set_column(3,3,74)
worksheet.set_column(6,6,20)
worksheet.set_column(4,4,18)


fin.close()
workbook.close()
