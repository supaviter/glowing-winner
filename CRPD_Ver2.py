import pyodbc # library used for Pandas
import datetime
import re, pandas, os.path, time, sys, os
import shutil # to Move and Rename a file from one path to another
#import pyperclip # Library to Copy and paste via Clipboard
import threading # For Multi Threading
import pyautogui

from time import sleep
from selenium import webdriver  # For initiating webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys   # For feeding user creditials
from selenium.common.exceptions import TimeoutException, NoSuchElementException # Used for defining exceptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import ElementNotVisibleException
from datetime import datetime, timedelta # To Extract System Date
from selenium.webdriver.support import expected_conditions as EC
from os import path
from subprocess import check_output # To implement cmd code
from datetime import date

today = date.today() # Current System date
today_date = today.strftime("%Y-%m-%d")
print(today_date)

prev_date = datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d') # Code for 1 Previous date.
print(prev_date)

dateTimeObj = datetime.now()
timestampStr = dateTimeObj.strftime('%Y_%m_%d_%H_%M_%S') # Code for current Date and Time
print('Current Timestamp :', timestampStr)

req_url = r"R:\Development_NII\CRPD_Auto\Ver2\Processor\Config_Ver2.xlsx"
df = pandas.read_excel(req_url, error_bad_lines=False) # Reading data from Configuration file in Excel format:
chromedriver_path = df.loc[0][1] # Chrome driver
#print(chromedriver_path)
download_path = df.loc[1][1] # Downloading path
#print(download_path)

chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : download_path} # Path to Download file
chrome_options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(chromedriver_path, chrome_options=chrome_options)
#driver.set_window_position(-10000,0)

url = df.loc[2][1] # Main Home page of Target website
driver.get(url)
time.sleep(4)
username = driver.find_element_by_id("j_username") # User attribute name in source code
username.send_keys('joe')
password = driver.find_element_by_id("j_password") # Password attribute name in source code
password.send_keys('password') # User's Password for Login
login_attempt = driver.find_element_by_xpath('//*[@id="container_header"]/div/table/tbody/tr[7]/td/input').click()
time.sleep(10)
print("Login Successfull")

# Function to perform Click action:
def simulate(timestampStr, prev_date, today_date, download_path):
	# Picking Required URL from Excel file:
	req_url = r"R:\Development_NII\CRPD_Auto\Ver2\Processor\Config_Ver2.xlsx"
	df = pandas.read_excel(req_url, error_bad_lines=False) # Reading data from Configuration file in Excel format:
	rpt_p1_OBD = df.loc[3][1] # [r][c]
	rpt_p4_OBD = df.loc[3][2] # Last part of URL
	mv_fm_OBD = df.loc[3][6] # Moving downloaded file from
	mv_to_OBD = df.loc[3][7] # Downloaded file final location
	#ptnr_OBD = df.loc[3][3]
	#print(rpt_p1_OBD, rpt_p4_OBD, mv_fm_OBD, mv_to_OBD)

	rpt_p1_ERA = df.loc[4][1]
	rpt_p4_ERA = df.loc[4][2]
	mv_fm_ERA = df.loc[4][6]
	mv_to_ERA = df.loc[4][7]
	#ptnr_ERA = df.loc[4][3]
	#print(rpt_p1_ERA, rpt_p4_ERA, mv_fm_ERA, mv_to_ERA)

	time.sleep(10)
	# Code for click opertaion using selenium webdriver
	folder_click = driver.find_element_by_xpath('//*[@id="unified_remit_operation_report"]/table/tbody/tr/td[1]/img').click()
	time.sleep(10)

	file_click = driver.find_element_by_xpath('//*[@id="file-revremit_outbound_files.prpt"]/tbody/tr/td[1]/img')
	# Code for double click
	actionchains = ActionChains(driver)
	actionchains.double_click(file_click).perform()
	time.sleep(15)

	print("Loading...")
	rpt_p2 = "&choice=Deposit+Date&From+Date="
	rpt_p3 = "T00%3A00%3A00.000-0400&To+Date="

	print("Enter 'From Date' as YYYY-MM-DD")
	prev_date = input()
	print("Enter 'To Date' as YYYY-MM-DD")
	today_date = input()
	print("Enter Partner name")
	ptnr = input()
	#print(prev_date)
	#print(today_date)

	# ---- Code for adding another Tab in web browser within Same Window (To avoid re-login into Nav Pentaho website) # -----
	# ** Outbound File Download **
	body = driver.find_element_by_tag_name("body")
	body.send_keys(Keys.CONTROL + 't') # Performing CTRL + T Key operation
	#partner = "BBVA"
	down_url1 = rpt_p1_OBD+ptnr+rpt_p2+prev_date+rpt_p3+today_date+rpt_p4_OBD
	#print('\n', down_url1)
	driver.get(down_url1)
	print("Waiting for report download response")
	time.sleep(10)
	shutil.move(mv_fm_OBD, mv_to_OBD+ptnr+'_'+timestampStr+'.csv')
	print("Outbound File Downloaded Successfully")


	# ** ERA outbound File Download **
	body = driver.find_element_by_tag_name("body")
	body.send_keys(Keys.CONTROL + 't') # Performing CTRL + T Key operation
	#partner = "BBVA"
	down_url2 = rpt_p1_ERA+ptnr+rpt_p2+prev_date+rpt_p3+today_date+rpt_p4_ERA
	#print(down_url2)
	driver.get(down_url2)
	print("Waiting for report download response")
	time.sleep(10)
	os.rename(download_path+r'\revremit_era_outbound_files .csv',mv_fm_ERA) # Renaming due to Space issue in Filename
	shutil.move(mv_fm_ERA, mv_to_ERA+ptnr+'_'+timestampStr+'.csv')
	print("ERA Outbound File Downloaded Successfully")

	driver.close()

# Code Starts from Here...
t1 = threading.Thread(target=simulate, args=(timestampStr,prev_date,today_date,download_path))
t1.start()