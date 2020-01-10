# -*- coding: utf-8 -*-
"""
Created on Mon Dec  2 14:43:46 2019

@author: Dennis Huynh

Description: Webscraping program. Read the contents of an excel file and store it into a 2D array. 
             Log into LinkedIn. Use Google search to find pages on LinkedIn, then copy the 
             information from the About section on LinkedIn page for empty cells and store it into
             the 2D array. Export the 2D array into a .csv file.
             *Note: When running script, after a certain number of automated LinkedIn log ins and 
             Google searches, a reCAPTCHA test will occur.
"""

# Import statements
import time
import csv
import xlrd
import sys
import getpass
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from parsel import Selector

# Dictionary for provinces, territories, and states
prov_terr_states = {
    'AB': 'Alberta',
    'BC': 'British Columbia',
    'MB': 'Manitoba',
    'NB': 'New Brunswick',
    'NL': 'Newfoundland and Labrador',
    'NT': 'Northwest Territories',
    'NS': 'Nova Scotia',
    'NU': 'Nunavut',
    'ON': 'Ontario',
    'PE': 'Prince Edward Island',
    'QC': 'Quebec',
    'SK': 'Saskatchewan',
    'YT': 'Yukon',

    'AK': 'Alaska',
    'AL': 'Alabama',
    'AR': 'Arkansas',
    'AS': 'American Samoa',
    'AZ': 'Arizona',
    'CA': 'California',
    'CO': 'Colorado',
    'CT': 'Connecticut',
    'DC': 'District of Columbia',
    'DE': 'Delaware',
    'FL': 'Florida',
    'GA': 'Georgia',
    'GU': 'Guam',
    'HI': 'Hawaii',
    'IA': 'Iowa',
    'ID': 'Idaho',
    'IL': 'Illinois',
    'IN': 'Indiana',
    'KS': 'Kansas',
    'KY': 'Kentucky',
    'LA': 'Louisiana',
    'MA': 'Massachusetts',
    'MD': 'Maryland',
    'ME': 'Maine',
    'MI': 'Michigan',
    'MN': 'Minnesota',
    'MO': 'Missouri',
    'MP': 'Northern Mariana Islands',
    'MS': 'Mississippi',
    'MT': 'Montana',
    'NA': 'National',
    'NC': 'North Carolina',
    'ND': 'North Dakota',
    'NE': 'Nebraska',
    'NH': 'New Hampshire',
    'NJ': 'New Jersey',
    'NM': 'New Mexico',
    'NV': 'Nevada',
    'NY': 'New York',
    'OH': 'Ohio',
    'OK': 'Oklahoma',
    'OR': 'Oregon',
    'PA': 'Pennsylvania',
    'PR': 'Puerto Rico',
    'RI': 'Rhode Island',
    'SC': 'South Carolina',
    'SD': 'South Dakota',
    'TN': 'Tennessee',
    'TX': 'Texas',
    'UT': 'Utah',
    'VA': 'Virginia',
    'VI': 'Virgin Islands',
    'VT': 'Vermont',
    'WA': 'Washington',
    'WI': 'Wisconsin',
    'WV': 'West Virginia',
    'WY': 'Wyoming'
}

''' Read an xlsx file '''
def readFile(fileName): 
    # read the xlsx file
    book = xlrd.open_workbook(fileName)

    # allow user input for the sheet name
    sheet_name = input("Please enter the name of excel sheet (I.e. Pivot Table - For SF): ")

    # read a specified sheet
    sheet = book.sheet_by_name(sheet_name)

    # create a 2D array to store the contents of the sheet
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

    return data

''' Log into LinkedIn '''
def login():
    # driver.get method() will navigate to a page given by the URL address
    driver.get('https://www.linkedin.com/login')

    # get username from user
    un = input("Please enter your LinkedIn username: ")
    
    # get password from user
    pw = getpass.getpass(prompt='Please enter your LinkedIn password: ')

    # locate email form by_class_name
    username = driver.find_element_by_id('username')

    # send_keys() to simulate key strokes
    # input an email
    username.send_keys(str(un))

    # sleep for 0.5 seconds
    time.sleep(0.5)

    # locate password form by_class_name
    password = driver.find_element_by_id('password')

    # send_keys() to simulate key strokes
    # input a password
    password.send_keys(str(pw))

    # sleep for 0.5 seconds
    time.sleep(0.5)

    # Locate submit button by_class_name
    log_in_button = driver.find_element_by_xpath('//*[@type="submit"]')

    # .click() to mimic button click
    log_in_button.click()

    # Sleep for 1 second
    time.sleep(1)

''' Export the data into a .csv file '''
def makeFile(row_list):
    # allow user input to name the new .csv file
    fileName = input("Please save the file as a .csv. Enter the name you wish to save it as (Example: Hemlock-Lead Database.csv): ")

    # Write all the information to a new csv file (replace 'Hemlock-Lead Database.csv' with fileName)
    with open(fileName, 'w', newline = '', encoding = 'utf-8') as file:
        # declare a csv writer
        writer = csv.writer(file)

        # writerow() method to the write to the file object
        writer.writerows(row_list)

''' Scrape information from LinkedIn '''
# parameters are the column headers in the excel file
def scrape(ca, co, d, fo, s, c, sp, zp, w, fn, ln, t, eml, p, e, ls, o):
    # a variable that stores the page source as text
    sel = Selector(text=driver.page_source)

    # if company does not exist, scrape for data
    if co.isspace():
        # scrape the company name
        co = sel.xpath('/html/body/div[5]/div[4]/div[3]/div/div[2]/section/div/div/div[2]/div[1]/div[1]/div/div[2]/div/h1/span/text()').get()
        # if company exists, remove white spaces
        if co:
            co = co.strip()

    # if description does not exist, scrape for data
    if d.isspace():
        # (Need-to-do) Break down description into 3 parts: Owler revenue estimate, Hoovers revenue estimate, Description of company - use concatenation
        d = sel.xpath('/html/body/div[5]/div[4]/div[3]/div/div[2]/div/div[2]/div[1]/section/p/text()').get()

        # if description exists, remove white spaces
        if d:
            d = d.strip()

    # if founded does not exist, scrape for data
    if fo.isspace():
        # scrape the year the company was founded casted to an int
        fo = sel.xpath('/html/body/div[5]/div[4]/div[3]/div/div[2]/div/div[2]/div[1]/section/dl/dd[7]/text()').get()

        # if founded exists, remove white spaces
        if fo:
            fo = fo.strip()

    # if street, city, state/province, or zip/postal is empty, scrape for their data
    if (s.isspace() or c.isspace() or sp.isspace() or zp.isspace()):
        # declare a variable to store the scraped information for location
        location = sel.xpath('/html/body/div[5]/div[4]/div[3]/div/div[2]/div/div[2]/div[1]/div/div[1]/h3/div/p/text()').get()

        # if location exists, remove white spaces
        if location:
            location = location.strip()
            # split the location info by commas and store the info in a list
            if ',' in location:
                locationInfo = [x.strip() for x in location.split(',')]

                # street, city, state/province, and zip/postal code can be extracted from location variable
                # checks if there is a Suite number
                if len(locationInfo) == 5:
                    s = locationInfo[0] + " - " + locationInfo[1]
                    c = locationInfo[2]
                    # declare a variable to split state/province and zip/postal code
                    spzp = locationInfo[3].split()
                    # get the state/province using the prov_terr_states dictionary
                    sp = prov_terr_states.get(spzp[0]) 
                    # checks if it is a postal code
                    if len(spzp) == 3:
                        #postal code
                        zp = spzp[1] + spzp[2] 
                    # otherwise, it is a zip code
                    else:
                        #zip code
                        zp = spzp[1]
                # if the format of location is greater than 5 or less than 4
                elif len(locationInfo) > 5 or len(locationInfo) < 4:
                    print("Please manually check the location information of this company: " + co)
                    s = location
                    c = " "
                    sp = " "
                    zp = " "
                # otherwise, standard location format
                else:
                    s = locationInfo[0]
                    c = locationInfo[1]
                    # declare a variable to split state/province and zip/postal code
                    spzp = locationInfo[2].split()
                    # get the state/province using the prow_terr_states dictionary
                    sp = prov_terr_states.get(spzp[0]) 
                    # checks if it is a postal code
                    if len(spzp) == 3:
                        #postal code
                        zp = spzp[1] + spzp[2] 
                    # otherwise, it is a zip code
                    else:
                        #zip code
                        zp = spzp[1]
            else:
                print("Please manually check the location information of this company: " + co)
                s = location
                c = " "
                sp = " "
                zp = " "

    # if website does not exist, scrape for data
    if w.isspace():
        w = sel.xpath('/html/body/div[5]/div[4]/div[3]/div/div[2]/div/div[2]/div[1]/section/dl/dd[1]/a/span/text()').get()
        # if website exists, remove white spaces
        if w:
            w = w.strip()

    # if No. of Employees does not exist, scrape for data
    if e.isspace():
        e = sel.xpath('/html/body/div[5]/div[4]/div[3]/div/div[2]/div/div[2]/div[1]/section/dl/dd[4]/text()').get()
        # if number of employees exists, remove white spaces
        if e:
            e = e.strip()
            # remove " on LinkedIn"
            eString = e.split()
            # gets the exact number on LinkedIn
            e = eString[0]

    # store all the information into a 1D list
    data_list = [ca, co, d, fo, s, c, sp, zp, w, fn, ln, t, eml, p, e, ls, o]
    # return the list as an updated row
    return data_list

''' Case 1 - Assume Campaign and Company is given '''
def case1():
    # allow user input for what file to choose
    fileName = input("Please enter the name of an xlsx file (I.e. Hemlock-Lead Database-TEMPLATE-190429.xlsx): ")
    # read the .xlsx database file and stores it into a 2D list
    row_list = readFile(fileName)

    # calls login function
    #login()

    # start at row 1
    r = 1
    # do-while loop
    # do: scrape for data and fill in missing information
    # while: row is not equal to the last row in the file
    while True:

    # variable declarations (initalize them for each row of data entries)
    # note campaign, company, fn, ln, title, email, phone, ls, and outreach will likely never be modified
        campaign = row_list[r][0]
        company = row_list[r][1] # will use this in the Google search
        description = row_list[r][2]
        founded = str(row_list[r][3])
        street = row_list[r][4]
        city = row_list[r][5]
        sp = row_list[r][6] # state/province
        zp = row_list[r][7] # zip/postal code
        website = row_list[r][8]
        fn = row_list[r][9] # first name
        ln = row_list[r][10] # last name
        title = row_list[r][11]
        email = row_list[r][12]
        phone = row_list[r][13]
        employees = str(row_list[r][14])
        ls = row_list[r][15] # lead source
        outreach = row_list[r][16]

        '''Google search'''
        # driver.get method() will navigate to a page given by the URL address
        driver.get('https://www.google.ca/')
        time.sleep (2)

        # locate search form by_name
        search_query = driver.find_element_by_name('q')
        # enter search query (Company name)
        search_query.send_keys(company + " AND site:linkedin.com/company") #Fill this part with company name (cn)
        # sleep for 0.5 seconds
        time.sleep(0.5)

        # .send_keys() to simulate the enter key 
        search_query.send_keys(Keys.RETURN)
        time.sleep(1)

        # clicks the matched result to get to linkedin page
        # finds url results
        results = driver.find_elements_by_xpath('//div[@class="r"]/a/h3')
        # clicks the first one
        results[0].click()
        # sleep for 1 seconds
        time.sleep(1)

        '''Goes to About page'''
        # take the current url and split it up by '/' char
        about_url = str(driver.current_url).split('/')
        # build the about url by adding https:// + www.linkedin.com + company + "company name" + about
        about = "https://" + str(about_url[2]) + "/" + str(about_url[3]) + "/" + str(about_url[4]) + "/about/"
        driver.get(about)
        # sleep for 2 seconds
        time.sleep(2)

        # updates the current row
        row_list[r] = scrape(campaign, company, description, founded, street, city, sp, zp, website, fn, ln, title, email, phone, employees, ls, outreach)

        # sleep for 5 seconds
        time.sleep(5)

        if r == (len(row_list)-1):
            break

    # terminates the application
    #driver.quit()

    # make the csv file
    makeFile(row_list)

    return

''' Case 2 - allow user input '''
def case2():
    # calls login function
    #login()

    # variable declarations (initalize them for each row of data entries)
    # note fn, ln, title, email, phone, ls, and outreach will likely never be modified
    campaign = " " # will use this in the Google search
    company = " "
    description = " "
    founded = " "
    street = " "
    city = " "
    sp = " " # state/province
    zp = " " # zip/postal code
    website = " "
    fn = " " # first name
    ln = " " # last name
    title = " "
    email = " "
    phone = " "
    employees = " "
    ls = " " # lead source
    outreach = " "

    print("Please fill in your search query. Press Enter to skip that field.")

    campaign = input("Please enter a Campaign: ")
    #if there is no campaign entered, go back to main menu
    if not campaign:
        menu()
    
    founded = input("Please enter a Founded date: ")
    if not founded:
        founded = " "

    city = input("Please enter a City: ")
    if not city:
        city = " "

    sp = input("Please enter a State/Province: ")
    if not sp:
        sp = " "

    employees_min = input("Please enter the lower range of No. of Employees: ")
    if not employees_min:
        employees_min = 1

    employees_max = input("Please enter the upper range of No. of Employees ")
    # if no employees max range was specified
    if not employees_max:
        # if employees min range exists
        if employees_min:
            # employees max range is equal to twice the min range or 100, whichever is greater
            employees_max = max(int(employees_min) * 2, 100)
        # otherwise, the default of the employees max range is 100
        employees_max = 100

    '''Google search'''
    # change the number of Google results to 50
    driver.get('https://www.google.ca/preferences?hl=en')
    time.sleep(2)
    # slide the value to 50
    search = driver.find_element_by_class_name('goog-slider-thumb')
    move = ActionChains(driver)
    move.click_and_hold(search).move_by_offset(152, 0).release().perform()
    time.sleep(1)

    # find and click the save button
    save_button = driver.find_element_by_xpath('//*[@id="form-buttons"]/div[1]')
    save_button.click()
    # wait 150 seconds in case a reCAPTCHA test appears (time can be edited - consider using a bot to bypass reCAPTCHAs)
    time.sleep(150)

    # locate search form by_name
    search_query = driver.find_element_by_name('q')

    # send_keys() to simulate the search text key strokes (input the parameters from 2D array)
    # format search query with ["site:linkedin.com/company" AND "CAMPAIGN_NAME" AND "LOCATION/REGION" AND ... ; for employee number add "# .. #"]
    if campaign and founded and city and sp:
        search_query.send_keys('site:linkedin.com/company AND ' + campaign + " AND founded on " + founded + " AND " + city + ", " + sp + " " + employees_min + ".." + employees_max)
    elif campaign and city and sp:
        search_query.send_keys('site:linkedin.com/company AND ' + campaign + " AND " + city + ", " + sp + " " + employees_min + ".." + employees_max)
    elif campaign and city:
        search_query.send_keys('site:linkedin.com/company AND ' + campaign + " AND " + city + " " + employees_min + ".." + employees_max)
    elif campaign and sp:
        search_query.send_keys('site:linkedin.com/company AND ' + campaign + " AND " + sp + " " + employees_min + ".." + employees_max)
    else:
        search_query.send_keys('site:linkedin.com/company AND ' + campaign + " " + employees_min + ".." + employees_max)

    # .send_keys() to simulate the enter key 
    search_query.send_keys(Keys.RETURN)
    time.sleep(1)

    # locate URL by_class_name
    linkedin_urls = driver.find_elements_by_partial_link_text("LinkedIn")

    # declare a list to store urls
    url_list = []

    # change the type of each element in the href tags to string and store it into the list of urls
    for elem in linkedin_urls:
        url_list.append(str(elem.get_attribute("href")))

    # print all elements in the list (if error occurs, at least have a list of 50 urls that can be manually checked)
    for x in url_list:
        print(x)
    time.sleep(0.5)

    # format first row of data to hold all the headers
    row_list = [["Campaign", "Company", "Description", "Founded", "Street", "City", "State/Province", 
                 "Zip/Postal Code", "Website", "First Name", "Last Name", "Title", "E-mail", "Phone", 
                 "No. of Employees", "Lead Source", "Outreach"]]
    
    # For loop to iterate over each URL in the list
    for linkedin_url in url_list:

        # get the profile URL 
        driver.get(linkedin_url)

        # add a 5 second pause to load each URL
        time.sleep(5)

        '''Goes to About page'''
        # take the current url and split it up by '/' char
        about_url = str(driver.current_url).split('/')
        # build the about url by adding https:// + www.linkedin.com + company + "company name" + about
        about = "https://" + str(about_url[2]) + "/" + str(about_url[3]) + "/" + str(about_url[4]) + "/about/"
        driver.get(about)
        # sleep for 2 seconds
        time.sleep(2)
        
        # add a row of data
        row_list.append(scrape(campaign, company, description, founded, street, city, sp, zp, website, fn, ln, title, email, phone, employees, ls, outreach))

        # sleep for 5 seconds
        time.sleep(5)

    # terminates the application
    #driver.quit()

    # make the csv file
    makeFile(row_list)
    
    return

''' Main Menu '''
def menu():
    print("************MAIN MENU**************")
    time.sleep(1)
    print()

    choice = input("""
    1: Fill-in an incomplete excel file
    2: Search for companies
    Q: Quit

    Please enter your choice: """)

    if choice == "1":
        case1()
        print("A new .csv file has been created")
        menu()
    elif choice == "2":
        case2()
        print("A new .csv file has been created")
        menu()
    elif choice == "Q" or choice == "q":
        driver.quit()
        sys.exit
    else:
        print("You must only select either 1 or 2")
        print("Please try again")
        menu()

############################# MAIN PROGRAM #############################
''' Initialize Chrome WebDriver '''
# removes "DevTool is listening on..." warning
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])

# specifies the path to the chromedriver.exe
driver = webdriver.Chrome(executable_path='C:/Users/denni/Documents/HEMLOCK/chromedriver', options=options)

# login to LinkedIn
login()

menu()

#Do not forget error prevention (make sure websites have the field that program is scraping for) - feature to be added