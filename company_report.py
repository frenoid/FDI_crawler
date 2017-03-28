# automated downloading of Signals from FDI markets from sys import argv
from sys import exit, argv
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
from os import chdir, remove, rename, listdir, getcwd
from shutil import copy, move
from math import ceil

def readDownloadDir(config_file):
	cwd = getcwd()
	with open(config_file, "r") as config:
		download_dir = config.readline()

	try:
		chdir(download_dir)
		print "%s is the download directory" % (download_dir)
	except OSError:
		print "Cannot access %s" % (download_dir)
		download_dir = "Invalid"

	finally:
		chdir(cwd)

	return download_dir

def browserInitialize(report_page):
	profile = FirefoxProfile()

	# .doc and .docx
	profile.set_preference("browser.download.folderList",2)
	profile.set_preference("browser.download.manager.showWhenStarting","false")
	profile.set_preference("browser.download.dir","c:/Users/faslxkn/Downloads/fdi_co_reports")
	profile.set_preference("browser.helperApps.neverAsk.saveToDisk",\
			       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
	print "Set to auto-download .docx"


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


def chooseDestRegion(driver, region):


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


def getCompanyCount(driver):

	company_count_elem = WebDriverWait(driver, 30).until(\
               	            EC.presence_of_element_located((By.XPATH, "/html/body/center/div/div[1]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[2]/div/div/table/tbody/tr[3]/td/a"))
               	            )
	company_count_text = filter(lambda ch: ch not in ",", company_count_elem.text)
	print "%s companies found" % (company_count_text)


	return int(company_count_text)
	
def set20RowsPerPage(driver):
	success = False
	try:
		set_to_20_link = WebDriverWait(driver, 10).until(\
		       	         EC.presence_of_element_located((By.XPATH,\
			         "/html/body/center/div/div[3]/div/div[2]/div/div/table[1]/tfoot/tr/td/div[2]/a[4]")) 
		                 )
		set_to_20_link.click()
		set_to_20_link.click()
		success = True
		
	except TimeoutException:
		print "Unable to set number of rows to 20"


	return success


def getPageTotal(driver, company_count, rows_per_page):
	return int(ceil(float(company_count)/float(rows_per_page)))

def getCompanyList(driver):
	print "Getting companies in page"
	company_list = {}
	row_no = 1
	while True:
		try:
			company_name = WebDriverWait(driver,10).until(\
			               EC.presence_of_element_located((By.XPATH,\
			               "/html/body/center/div/div[3]/div/div[2]/div/div/table[1]/tbody/tr["\
			               +str(row_no)+\
			               "]/td[1]/a"))
			               )
			company_link = WebDriverWait(driver,10).until(\
				       EC.presence_of_element_located((By.XPATH,\
                                       "/html/body/center/div/div[3]/div/div[2]/div/div/table[1]/tbody/tr["
                                       +str(row_no)+\
				       "]/td[7]/a"))
                                       )
			# print "Company %d is %s" % (row_no, company_link.text)
			company_list[company_name.text] = company_link
			row_no += 1
		except (TimeoutException, UnexpectedAlertPresentException, NoSuchElementException):
			print "End of list table"
			print "%d Companies found in page" % (len(company_list))
			break

	return company_list

def getDownloadName(url):
	filename = "Invalid"

	try:
		last_slash_idx = url.rindex("/")
		filename = url[last_slash_idx+1:]
	except ValueError:
		print "Filename not found"

	return filename



def downloadReport(driver, download_dir, company_name, company_link):
	# Click on report
	download_name = "Invalid"
	company_link.click()
	print "Open report dialog"
	sleep(3)

	try:
		# Name report
		name_report = WebDriverWait(driver,30).until(\
		      EC.element_to_be_clickable((By.XPATH,\
                      "/html/body/div[4]/div[2]/div/form/table[1]/tbody/tr/td[2]/a"))
		      )

		name_report.click()
		print "Use default name for report"
		sleep(3)

		# Create report
		create_report = WebDriverWait(driver,10).until(\
		        EC.presence_of_element_located((By.XPATH,\
                        "/html/body/div[4]/div[11]/div/button[1]"))
		        )
		create_report.click()
		print "Generating report"

		# Wait for download link then click on it
		download_elem = WebDriverWait(driver,85).until(\
		        EC.presence_of_element_located((By.XPATH,\
                        "/html/body/div[5]/div[2]/div/div/div/a"))
		        )

		download_elem.click()
		download_link = download_elem.get_attribute("href")
		download_name = getDownloadName(download_link)
		print "Get report from %s" % (download_link)
		sleep(1)

		# Close the download dialog
		close_button =  WebDriverWait(driver,10).until(\
		        EC.presence_of_element_located((By.XPATH,\
                        "/html/body/div[5]/div[11]/div/button"))
			)
		close_button.click()
		sleep(2)
		print "%s report is downloading" % (company_name)
	except(TimeoutException, NoSuchElementException, UnexpectedAlertPresentException):
		print "!Exception encountered while downloading report"
	finally:
		# Close any dialog windows
		try:
			close_button = WebDriverWait(driver,3).until(\
				       EC.presence_of_element_located((By.CLASS_NAME,\
				       "ui-dialog-titlebar-close ui-corner-all ui-state-hover"))
				       )
		except(TimeoutException, NoSuchElementException, UnexpectedAlertPresentException):
			pass

	return download_name

def closeExtraWindows(driver, main_window):
    for win_handle in driver.window_handles:
        if win_handle != main_window:
            driver.switch_to_window(win_handle)
            driver.close()

    driver.switch_to_window(main_window)

    return
		        

def goToPage(driver, page_no):
	page_index = WebDriverWait(driver,15).until(\
		     EC.presence_of_element_located((By.ID,\
		     "page_index"))\
		     )
	page_submit = WebDriverWait(driver,15).until(\
		      EC.presence_of_element_located((By.XPATH,\
                      "/html/body/center/div/div[3]/div/div[2]/div/div/table[1]/tfoot/tr/td/div[2]/img"))\
		      )
	page_index.click()
	page_index.clear()
	page_index.send_keys(str(page_no))

	page_submit.click()

	return

def checkDownloadDirClear(download_dir):
	dir_clear_status = True

	file_names = listdir(download_dir)
	for file_name in file_names:
		if file_name[-5:] == ".part":
			dir_clear_status = False
			break

	return dir_clear_status

def sweepDownloadDir(download_dir, destination_dir):
	file_names = listdir(download_dir)
	files_moved_counter = 0

	# test if dest_dir exists, if not, mkdir
	try:
		listdir(destination_dir)
	except WindowsError:
		mkdir(destination_dir)

	# Move downloaded files into storage_dir
	for file_name in file_names:
		if file_name[-5:] == ".docx":
			move(download_dir+"/"+file_name, destination_dir+"/"+file_name)
			files_moved_counter += 1
			# print str(file_name), "was moved to", str(destination_dir)

	return files_moved_counter

def renameFiles(filedir, filename_coname):
	files_renamed_count = 0

	for filename in filename_coname:
		try:
			newname = filename_coname[filename] + "_report.docx"
			rename(filedir+"/"+filename, filedir+"/"+newname)
			files_renamed_count += 1
			# print "%s renamed to %s" % (filename, newname)

		except OSError:
			print newname, "already exists!" 

	return files_renamed_count

def writeFailLog(download_dir, page_no, failed_companies):
	chdir(download_dir)
	with open(fdi_company_failed_reports.log, "a") as fail_log:
		fail_log.write("# On page", str(page_no), ": \n")
		for failed_company in failed_companies:
			fail_log.write(failed_company)

	return


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
if __name__ == "__main__":
	if(len(argv) < 2):
		print("Syntax: report.py <Dest. Region> <query_type> [query no]")
		exit("Insufficient arguments. At least 2")

	print"*************************************"
	print"fDi markets company report downloader"
	print"*************************************"

	# Initialize arguments
	dest_region = argv[1]
	query_type = argv[2]
	download_dir = readDownloadDir("download_dir.txt")

	if download_dir == "Invalid":
		exit("Invalid download dir")	

	print "Downloading dest_region: %s query_type: %s" % (dest_region, query_type)

	# Initialize the browser and load FDI markets
	report_page = "https://app.fdimarkets.com/markets/index.cfm"
	driver = browserInitialize(report_page)
	main_window = driver.current_window_handle

	# Login
	driver = fdiLogin(driver, "sooyeon.kim@nus.edu.sg", "ksooyeonGPN@NUS999")
	sleep(3)
	main_window = driver.current_window_handle
	# openSearchMenu(driver)
	print "Main page reached"

	# Check that default view is Company Database
	try:
		table_header = WebDriverWait(driver, 15).until(\
		       	       EC.presence_of_element_located((By.CLASS_NAME, "section_header")) 
		               )
		print "Default View: Company Database"

	except TimeoutException:
		print "Company Database header not found"
		print "Check that Company Database is default view"
		exit()

	# Get the number of companies
	try:
		company_count = getCompanyCount(driver)
	except TimeoutException:
		exit("Failed to get company count")

	# Set number of rows per page to predetermined value
	attempt_no = 0
	success = False
	while success != True and attempt_no < 5:
		success = set20RowsPerPage(driver)
	if success == False:
		exit("Failed to set correct rows per page")
		sleep(5)
	else:
		sleep(10)
		print "Rows set to 20 per page"

	# Get number of pages available
	page_total = getPageTotal(driver, company_count, 20)

	# Check if all values are intialized properly
	if(company_count>0 and page_total>0):
		print "There are %d companies in %d pages" % (company_count, page_total)
	else:
		exit("Company_count or page_total is 0")

	# Initiate the download list
	download_pages = []
	if(query_type == "all"):
		download_pages = range(1, page_total+1)
	elif(query_type == "list"):
		download_pages.extend(argv[3:])
		download_pages = [int(x) for x in download_pages]
	elif(query_type == "range"):
		download_pages = range(int(argv[3]), int(argv[4])+1)

	download_pages = sorted(download_pages)
	total_download_pages = len(download_pages)
	if query_type == "all" or query_type == "range":
		print "Downloading pages %d through %d" % (download_pages[0], download_pages[total_download_pages-1])
	else:
		print "Downloading pages: %s" % (str(download_pages))

	# Iterate over pages
	for page_no in download_pages:
		filename_coname = {}
		failed_companies = []

		if page_no != 1:
			goToPage(driver, page_no)
			sleep(5)

		print "***** Page #%d *****" % (page_no)

		try:
			company_list = getCompanyList(driver)
			# Iterate over companies in page
			for company_no, company_name in enumerate(company_list):
				print "* %d: %s" % (company_no+1, company_name)

				try:
					download_name = downloadReport(driver,download_dir,company_name, company_list[company_name])
					# print "File is %s" % (download_name)
					if download_name != "Invalid":
						filename_coname[download_name] = company_name
				except(TimeoutException, NoSuchElementException):
					print "!Error. Next firm"
					failed_companies.append(company_name)

				finally:
					closeExtraWindows(driver, main_window)

		except(TimeoutException, NoSuchElementException):
			print "!Exception while changing pages"

		finally:
			clear_status = False
			while clear_status != True:
				clear_status = checkDownloadDirClear(download_dir)
				if clear_status == False:
					print "Download in progress. Wait 10 secs"
					sleep(10)
				else:
					break

			destination_dir = download_dir + "/fdi_co_reports" 
			move_count = sweepDownloadDir(download_dir, destination_dir)
			print "%d files were moved to %s"  % (move_count, destination_dir)

			rename_count = renameFiles(destination_dir, filename_coname)
			print "%d files were renamed" % (rename_count)

			if len(failed_companies) > 0:
				writeFailLog(download_dir, page_no, failed_comapanies)
				print "%d failed reports written to log" % (len(failed_comapnies))

			driver.switch_to_window(main_window)
			print "Next page"
			print "===================="


	# Logout and close the browser
	fdiLogout(driver, main_window)

	print "Script End"
