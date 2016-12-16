# automated downloading of Signals from FDI markets 
from sys import argv
from sys import exit
from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException 
from selenium.common.exceptions import UnexpectedAlertPresentException 
from selenium.common.exceptions import NoAlertPresentException 
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import ElementNotVisibleException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
from datetime import timedelta
from time import sleep
from os import chdir, remove, rename, listdir
from math import ceil
from shutil import copy


def browserInitialize(report_page):
	profile = FirefoxProfile()
	profile.set_preference("browser.helperApps.neverAsk.saveToDisk",\
		       "application/vnd.ms-excel")
	profile.set_preference("browser.helperApps.neverAsk.saveToDisk",\
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	driver = webdriver.Firefox(firefox_profile = profile)
	driver.get(report_page)
	print "Browser loaded"

	return driver


def fdiLogin(driver, user_id, user_password):
	sleep(3)
	if driver.title == "fDi Markets - Global Investment Database - ***":
		username = driver.find_element(By.ID, "email")
		password = driver.find_element(By.ID, "password")
		username.send_keys(user_id)
		password.send_keys(user_password)
		sleep(2)
		password.send_keys(Keys.ENTER)
		print "Login info entered"
		
		WebDriverWait(driver,20).until(EC.title_contains("fDi Markets - Global Investment Database"))
		print "Login successful: " + driver.title

	else:
		print "Wrong page title: %s" % (driver.title)

	return driver

def getQueryRanges(months_per_query):
	query_ranges = {}

	# Signals database starts Nov-2009
	if (months_per_query == "all"):
		query_ranges[1] = [date(2009,11,2), date(2016,9,1)]
	else:	
		months_per_query = int(months_per_query)
		two_weeks = timedelta(weeks = 2)
		one_month = timedelta(weeks = 4)
		months_delta = timedelta(weeks = 4 * months_per_query)
		start_date = date(2009, 11, 2) 

		query_no = 0
		last_batch_done = False
	
		while last_batch_done == False:
			query_no += 1

			# Create the end_date and check if it exceeds the database range
			end_date = start_date + months_delta - one_month
			if end_date >= date(2016, 9, 1): # Database ends Sep-2016
				end_date = date(2016,9,1)
				last_batch_done = True

			# Insert start_date and end_date into the query list
			query_ranges[query_no] = [start_date, end_date]

			# Move the start date for next query
			old_start_date = start_date
			start_date = end_date + one_month
			# When 28 days is not enough to change the month
			if(start_date.month == old_start_date.month and start_date.year == old_start_date.year):
				start_date = start_date + two_weeks

	print "%d queries created" % (len(query_ranges))

	return query_ranges


def getNextDateRange(query_ranges, query_no):
	date_range = query_ranges[query_no]
	start_date = date_range[0]
	end_date = date_range[1]

	print "Start: %s-%s End: %s-%s" % (start_date.year, start_date.month, end_date.year, end_date.month)

	return start_date, end_date

def openSearchMenu(driver):


	# Start new search
	search_menu = WebDriverWait(driver, 15).until(\
                      EC.presence_of_element_located((By.ID, "search_nav_button"))
		      )
	search_menu.click()

	search_select = WebDriverWait(driver, 15).until(\
	                EC.presence_of_element_located((By.ID, "search_select"))
			)
	

	return

def clearDates(driver):

	""" For clearing country
	delete_country = WebDriverWait(driver, 15).until(\
			 EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[1]/div[1]/div/table/tbody/tr/td[3]/input[2]"))
			 )
	delete_country.click()
	sleep(5)
	"""

	delete_date_range = WebDriverWait(driver, 15).until(\
			    EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[1]/div[2]/div/table/tbody/tr/td[3]/input[2]"))
			    )
	delete_date_range.click()
	sleep(5)
 

	""" The drop-down method
	current_search_dropdown = WebDriverWait(driver, 15).until(\
				  EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[1]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/form/select"))
                                  )
	current_search_dropdown.click()

	all_data_option = WebDriverWait(driver, 15).until(\
			  EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[1]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/form/select/optgroup[1]/option[1]"))
                          )
	all_data_option.click()

	continue_without_saving = WebDriverWait(driver, 15).until(\
				  EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[3]/div/button[2]"))
				  )
	continue_without_saving.click()
	print "Click 1"
	continue_without_saving.click()
	print "Click 2"
	continue_without_saving.click()
	print "Click 3"
	"""

	return

def chooseSourceCountry(driver, country):


	# Choose filter criteria: Source Country
	search_select =  WebDriverWait(driver, 15).until(\
                         EC.presence_of_element_located((By.ID, "search_select"))
			 )
	search_select.click()
	search_select_source = WebDriverWait(driver, 15).until(\
			       EC.presence_of_element_located((By.XPATH,\
                               "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[3]/table/tbody/tr/td[2]/select/option[2]"))
			       )
	search_select_source.click()
	sleep(3)

	# Enter the Source Country name

	""" Code to search only countries
	search_level = WebDriverWait(driver, 15).until(\
		       EC.presence_of_element_located((By.ID, "search_level"))
		       )
	search_level.click()
	search_level_country = WebDriverWait(driver, 15).until(\
			       EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[1]/table/tbody/tr/td[1]/select/option[3]"))
                               )
	search_level_country.click()
	"""

	search_term = driver.find_element(By.ID, "search_term")
	search_term.send_keys(country)
	search_button = driver.find_element(By.ID, "search_button")
	search_button.click()
	search_button.click()
	sleep(3)

	add_country =  WebDriverWait(driver, 20).until(\
		       EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[1]/div[1]/table/tbody/tr[2]/td[5]/a[1]/strong"))
		       )
	add_country.click()
	sleep(3)
	confirm_selection = WebDriverWait(driver, 15).until(\
                            EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[3]/div[1]/a/span/span"))
			    )
	confirm_selection.click()
	sleep(5)

	return True

def chooseDateRange(year_start, month_start, year_end, month_end):

	sleep(5)
	
	# Filter criteria Date range
	search_select = WebDriverWait(driver, 20).until(\
			EC.presence_of_element_located((By.ID, "search_select"))
			)
	search_select.click()

	search_select_source = WebDriverWait(driver, 15).until(\
			       EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[3]/table/tbody/tr/td[2]/select/option[7]"))
                               )
	search_select_source.click()
	sleep(3)
	
	# Date range toggle: click if date range selector is a slider
	# Switch to drop-down menu
	try:
		menu_start_month = WebDriverWait(driver, 15).until(\
				   EC.presence_of_element_located((By.ID, "start_month"))
				   )
		print "Date dropdown menu found"
		menu_start_month.click()
	except ElementNotVisibleException:
		selector_toggle = WebDriverWait(driver, 15).until(\
				  EC.presence_of_element_located((By.XPATH,\
				  "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[1]/div/div/div/div/center[2]/a"))
				  )
		print "Selector_toggle found"
		sleep(3)
		selector_toggle.click()
		selector_toggle.click()
		selector_toggle.click()
		print "Use date-range selector toggle"
		sleep(5)

	# Choose date ranges
	menu_start_month = WebDriverWait(driver, 15).until(\
		           EC.presence_of_element_located((By.ID, "start_month"))
			   )
	menu_start_month.click()
	select_start_month = driver.find_element(By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[1]/div/div/center/div/form/select[1]/option["\
                             + str(month_start) +"]")
	select_start_month.click()

	menu_start_year = driver.find_element(By.ID, "start_year")
	menu_start_year.click()
	select_start_year = driver.find_element(By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[1]/div/div/center/div/form/select[2]/option["\
                            + str(year_start-2002) +"]")
	select_start_year.click()


	menu_end_month = driver.find_element(By.ID, "end_month")
	menu_end_month.click()
	select_end_month = driver.find_element(By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[1]/div/div/center/div/form/select[3]/option["\
                             + str(month_end) +"]")
	select_end_month.click()

	menu_end_year = driver.find_element(By.ID, "end_year")
	menu_end_year.click()
	select_end_year = driver.find_element(By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[1]/div/div/center/div/form/select[4]/option["\
                            + str(year_end-2002) +"]")
	select_end_year.click()

	# Confirm selection
	sleep(2)
	confirm_selection = WebDriverWait(driver, 15).until(\
			    EC.element_to_be_clickable((By.XPATH, "/html/body/center/div/div[3]/div[1]/table/tbody/tr/td/div[2]/div/div[3]/div[1]/a/span/span"))
                            )
	confirm_selection.click()
	sleep(5)


	return True

def getSignalCount(driver):

	signal_count_elem = WebDriverWait(driver, 30).until(\
               	            EC.element_to_be_clickable((By.XPATH, "/html/body/center/div/div[1]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[2]/div/div/table/tbody/tr[2]/td/a"))
               	            )
	signal_count = int(signal_count_elem.text)
	print "%d signals found" % (signal_count)

	return signal_count
	

def renameFile(downloaded_files, query_no, start_date, end_date):
	rename_success = False

	# Generate new filename = "<query#>_<start_date>_<end_date>.xls"
	new_name =  str(query_no) + "_" +\
                    str(start_date.year) + "-" + str(start_date.month) + "_" +\
                    str(end_date.year) + "-" + str(end_date.month) +\
                    ".xlsx"

	for file_name in downloaded_files:
		if (file_name[0:5] == "9940_" and file_name[-5:] == ".xlsx"):
			try:
				rename(file_name, new_name)
				rename_success = True
				print "%s renamed to %s" % (file_name, new_name)
				break
			except WindowsError:
				print "%s used by another file" % (new_name)
				break


	return rename_success

def fdiLogout(driver, main_window):
	try:
		driver.switch_to.window(main_window)
		logout_link = driver.find_element(By.LINK_TEXT, "Logout")
        	logout_link.click()
		sleep(2)
		print "Logging out and exiting"
		driver.close()

	except (TimeoutException, UnexpectedAlertPresentException, NoSuchElementException):
		print "!Exception encountered during logout"
	
	finally: 
		driver.close()

	return

### Main() ###


# Initialize arguments
source_country = argv[1]
months_per_query = argv[2]
query_type = argv[3]

if(len(argv) < 3):
	print("Syntax: report.py <Source Country> <# months per query> <query_type> [query no]")
	exit("Insufficient arguments. At least 2")

print "Downloading Source Country: %s Months per query: %s" % (source_country, months_per_query)

# Initialize date query ranges and query list
query_ranges = getQueryRanges(months_per_query)
query_total = int(len(query_ranges))
download_list = []

print argv[3]
if(query_type == "all"):
	download_list = range(1, query_total+1)
elif(query_type == "list"):
	download_list.extend(argv[4:])
	download_list = [int(x) for x in download_list]
elif(query_type == "range"):
	first_query = int(argv[4])
	last_query = int(argv[5])
	download_list = range(first_query, last_query + 1)

download_list = sorted(download_list)
print download_list



# Initialize the browser and load FDI markets
report_page = "https://app.fdimarkets.com/markets/index.cfm"
driver = browserInitialize(report_page)
main_window = driver.current_window_handle
print "FDI markets website loaded"

# Login
driver = fdiLogin(driver, "sooyeon.kim@nus.edu.sg", "ksooyeonGPN@NUS999")
sleep(3)

# Begin main loop
main_window = driver.current_window_handle
openSearchMenu(driver)
print "Main page reached"


for query_no in download_list:
	print "***** Query #%d *****" % (query_no)
	start_date, end_date = getNextDateRange(query_ranges, query_no)

	try:
		# Create filters: Source country + date range
		if(query_no == 1):
			chooseSourceCountry(driver, source_country)
			print "%s selected as Source Country" % (source_country)
	
		sleep(3)
		chooseDateRange(start_date.year, start_date.month, end_date.year, end_date.month)
	
		# Confirm the search criteria and choose database
		sleep(6)

		signal_count = getSignalCount(driver)

		# Case where there are zero signals found
		# Create a dummy file, rename it according to the query no
		# Proceed to next query
		if signal_count == 0:
			copy("C:/Selenium/fdi_markets/zero_signal_file.xlsx", "C:/Users/faslxkn/Downloads/9940_zerosignals.xlsx")
			print "Zero signal file created in downloads"

			# Rename downloaded file appropriately
			chdir("C:/Users/faslxkn/Downloads")
			downloaded_files = listdir("C:/Users/faslxkn/Downloads")
			rename_success = renameFile(downloaded_files, query_no, start_date, end_date)
			if rename_success == False:
				print "Renaming dummy file was unsuccessful"

			continue	

			

		# Where signal count is not zero, proceed to search database
		database = WebDriverWait(driver, 15).until(\
		   	   EC.element_to_be_clickable((By.XPATH, "/html/body/center/div/div[1]/table/tbody/tr[2]/td/div[1]/a[2]"))
                   	   )
		database.click()
		print "Searching database"
		sleep(3)


		# 4: Export the data in excel format and download
		export_link = WebDriverWait(driver, 20).until(\
                      	      EC.element_to_be_clickable((By.XPATH, "/html/body/center/div/div[3]/div/div/div/div[3]/div[2]/a/span/span"))
	              	      )
		export_link.click()
		print "Downloading %s Dates: %d-%d to %d-%d" % (source_country, start_date.year, start_date.month, end_date.year, end_date.month)
	
		download_link = export_link = WebDriverWait(driver, 120).until(\
                                EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Download Excel (.xlsx 2007+)"))
            			)
		download_link.click()
		print "Link found! Downloading excel file"
		sleep(5)

		close_button = driver.find_element(By.XPATH, "/html/body/div[3]/div[11]/div/button")   
		close_button.click()


		# Rename downloaded file appropriately
		chdir("C:/Users/faslxkn/Downloads")
		downloaded_files = listdir("C:/Users/faslxkn/Downloads")
		rename_success = renameFile(downloaded_files, query_no, start_date, end_date)
		if rename_success == False:
			print "Renaming was unsuccessful"		

	except (TimeoutException, UnexpectedAlertPresentException, NoSuchElementException):
		print "!Exception! Next batch."     
		driver.refresh()
	
	finally:
		driver.switch_to_window(main_window)
		openSearchMenu(driver)
		clearDates(driver)
		print "Clearing date ranges"
	
		sleep(5)

		print "=========="


# Logout and close the browser
fdiLogout(driver, main_window)



print "Script End"
