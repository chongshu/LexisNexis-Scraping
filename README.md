# LexisNexis-Scraping

## Disclaimer

Disclaimer: The author posed the following code for academic purposes and an illustration of Selenium only. Scraping LexisNexis may be a violation of LexisNexis user policy. Use at your own legal risk.


## Requirement 
LexisNexis Uni subscription

python v3.6

pandas v0.20.1

selenium v3.141

python-docx  v0.8.7

chromedriver.exe


## Getting Started 

The search query

	searchTerms = r'apple shareholder class action'

The Nexis Uni subscription Login Page (Below is the example of USC subscription)
	
	url = r'http://libguides.usc.edu/go.php?c=9232127'
	
Your Login Informaiton

	username = 'MyUSCPassUsername'
	password = 'MyUSCPassWord'
	
	
The folder you store this file (Below is example of my Desktop)
	
	root = r'C:\Users\chongshu\Desktop\LexisNexis'

Second you wish to restart the program if frozen

	dead_time = 300
	
Do not change the following code

	path_to_chromedriver = root + r'\chromedriver'
	download_folder = root + r'\download'

	
## Downloading

	download_file(url = url, searchTerms = searchTerms, username = username, \
                  dead_time = dead_time, path_to_chromedriver=path_to_chromedriver, \
                  download_folder = download_folder)
				  
				  

## Unzipping file

	unzip(download_folder=download_folder)
	
## Create Index

	create_index(download_folder=download_folder, searchTerms = searchTerms)
				  
## Note

You may also need customerize the following lines depending on your institution's login page:

	browser.find_element_by_id('username').send_keys(username)
	browser.find_element_by_id('password').send_keys(password)
	
If your subscription does not require a login, delete the whole login part.





















 


