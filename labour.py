from selenium.webdriver import Chrome
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from datetime import datetime
import time
import csv
import os
import sys


def element_exists(by, element):
    global driver

    """
    A function that checks if the text to scrape on the page exists
    or not.
    """

    try:
        if by == "class":
            driver.find_element_by_class_name(element)
        elif by == "xpath":
            driver.find_element_by_xpath(element)
    except NoSuchElementException:
        return False
    return True


print("-> Enter the page and row number you want to start scraping from below.")
print("-> Hit enter without typing anything if you wish to start with page number 1\n   and row number 2.")
print("-> Note: Row starts at 2, not 1 or 0.\n")
pageNumber = input("Enter Page No: ")
pageNumber = 1 if len(pageNumber) == 0 else int(pageNumber)

viewDetailXNumber = input("Enter Row No: ")
viewDetailXNumber = 2 if len(viewDetailXNumber) == 0 else int(viewDetailXNumber)

webDriver = "chromedriver.exe"  # Chrome driver
driver = Chrome(webDriver)  # Opens driver

date = datetime.now()  # Current date time
csvFile = "Employment_Agency_Data_{}{}{}.csv".format(date.day, date.month, date.year)  # File name

# If the CSV file doesn't exist in the path then create it and add the rows to it before starting
# However, if it exists it will append the new scraped values to the bottom of the existing CSV.
if not os.path.exists(csvFile):
    print("-> {} doesn't exist, creating ...".format(csvFile))
    with open(csvFile, "w", encoding="utf-8-sig", newline="") as writeFile:
        write_csv = csv.writer(writeFile)
        write_csv.writerow(['English Company Name', 'Chinese Company Name', 'Valid Licence Since',
                            'District', 'Address', 'Telephone No', 'Fax', 'Email', 'Placement Type'])

run = True  # Stops the script if equals to False
agreePage = True  # Reads the disclaimer and statement pages in order to accept them
processNext = True  # Process if errors

try:
    while run:
        if agreePage:
            """
            I've put a wait time of 5 seconds, so a total of 15 seconds before starting the scraping
            in order to avoid being sent back to the disclaimer and statement pages.
            """

            print("-> Reading URL: https://www.eaa.labour.gov.hk/en/disclaimer-search.html")
            print("-> Waiting 5 seconds before accepting to avoid being sent back to this page ...")

            driver.get("https://www.eaa.labour.gov.hk/en/disclaimer-search.html")

            time.sleep(2)

            print("-> Reading URL: {}".format(driver.current_url))

            disclaimerBtn = driver.find_element_by_xpath('//*[@id="button-i-accept"]')
            disclaimerBtn.click()

            print("-> Reading URL: {}".format(driver.current_url))
            print("-> Waiting 5 seconds again ...")

            time.sleep(2)

            statementBtn = driver.find_element_by_xpath('//*[@id="button-i-accept"]')
            statementBtn.click()

            print("-> The scraping will start in 5 seconds without interruption, hang on ...")

            time.sleep(2)

            print("-> Using URL: https://www.eaa.labour.gov.hk/en/result.html?page-no={}&row-per-page=30"
                  "&list_all_agencies=all&sort-by=EN_NAME_ASC&types=".format(pageNumber))

            driver.get("https://www.eaa.labour.gov.hk/en/result.html?page-no={}&row-per-page=30"
                       "&list_all_agencies=all&sort-by=EN_NAME_ASC&types=".format(pageNumber))

            agreePage = False
            processNext = True

        if driver.current_url == "https://www.eaa.labour.gov.hk/en/disclaimer-search.html":
            print("-> Re-accepting the terms ..")
            agreePage = True

        if viewDetailXNumber == 32:
            # If we finished reading the page, change page and start with the first row (2).
            pageNumber += 1
            viewDetailXNumber = 2

            print("-> Using URL: https://www.eaa.labour.gov.hk/en/result.html?page-no={}&row-per-page=30"
                  "&list_all_agencies=all&sort-by=EN_NAME_ASC&types=".format(pageNumber))

            driver.get("https://www.eaa.labour.gov.hk/en/result.html?page-no={}&row-per-page=30"
                       "&list_all_agencies=all&sort-by=EN_NAME_ASC&types=".format(pageNumber))

        if driver.current_url != "https://www.eaa.labour.gov.hk/en/disclaimer-search.html":
            print("-> Reading Page no. {} | row no. {}".format(pageNumber, viewDetailXNumber))

            try:
                """
                If the View Details button isn't visible, we have to wait until it is.
                However if it still can't find it in a range of 30 seconds, it will run
                out of ideas.
                """
                viewDetailsBtn = driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[{}]/p[5]/a'
                                                              .format(viewDetailXNumber))
                print(viewDetailsBtn)
                driver.execute_script("arguments[0].click();", viewDetailsBtn)
            except Exception as e:
                """
                It just couldn't click the View Details button.
                """
                print("-> Couldn't click View Details properly ...")
                print(e)

                # agreePage = True
                processNext = False

                time.sleep(5)

                driver.get("https://www.eaa.labour.gov.hk/en/result.html?page-no={}&row-per-page=30"
                           "&list_all_agencies=all&sort-by=EN_NAME_ASC&types=".format(pageNumber))

                time.sleep(5)

            if processNext:
                #  Self explanatory
                if not element_exists("class", "en-name"):
                    print("-> 'English name' doesn't exist, assigning empty value.")
                    engName = ""
                else:
                    engName = driver.find_element_by_class_name("en-name").text

                if not element_exists("class", "chi-name"):
                    print("-> 'Chinese name' doesn't exist, assigning empty value.")
                    chiName = ""
                else:
                    chiName = driver.find_element_by_class_name("chi-name").text

                if not element_exists("xpath", '//*[@id="main"]/div/p[2]'):
                    print("-> 'Licence since' doesn't exist, assigning empty value.")
                    licenceSince = ""
                else:
                    licenceSince = driver.find_element_by_xpath('//*[@id="main"]/div/p[2]').text
                    if len(licenceSince) == 1:
                        print("-> 'Licence since' has no value, empty.")
                        licenceSince = ""

                if not element_exists("xpath", '//*[@id="main"]/div/p[4]'):
                    print("-> 'District' doesn't exist, assigning empty value.")
                    district = ""
                else:
                    district = driver.find_element_by_xpath('//*[@id="main"]/div/p[4]').text
                    if len(district) == 1:
                        print("-> District has no value, empty.")
                        district = ""

                if not element_exists("xpath", '//*[@id="main"]/div/p[6]'):
                    print("-> 'Address' doesn't exist, assigning empty value.")
                    address = ""
                else:
                    address = driver.find_element_by_xpath('//*[@id="main"]/div/p[6]').text
                    if len(address) == 1:
                        print("-> 'Address' has no value, empty.")
                        address = ""

                if not element_exists("xpath", '//*[@id="main"]/div/p[8]'):
                    print("-> 'Phone number' doesn't exist, assigning empty value.")
                    phoneNumber = ""
                else:
                    phoneNumber = driver.find_element_by_xpath('//*[@id="main"]/div/p[8]').text
                    if len(phoneNumber) == 1:
                        print("-> 'Phone No.' has no value, empty.")
                        phoneNumber = ""

                if not element_exists("xpath", '//*[@id="main"]/div/p[10]'):
                    print("-> 'Fax number' doesn't exist, assigning empty value.")
                    faxNumber = ""
                else:
                    faxNumber = driver.find_element_by_xpath('//*[@id="main"]/div/p[10]').text
                    if len(faxNumber) == 1:
                        print("-> 'Fax No.' has no value, empty.")
                        faxNumber = ""

                if not element_exists("xpath", '//*[@id="main"]/div/p[12]'):
                    print("-> 'Email' doesn't exist, assigning empty value.")
                    email = ""
                else:
                    email = driver.find_element_by_xpath('//*[@id="main"]/div/p[12]').text
                    if len(email) == 1:
                        print("-> 'Email' has no value, empty.")
                        email = ""

                if not element_exists("xpath", '//*[@id="main"]/div/p[14]'):
                    print("-> 'Placement type' doesn't exist, assigning empty value.")
                    placementType = ""
                else:
                    placementType = driver.find_element_by_xpath('//*[@id="main"]/div/p[14]').text
                    if len(placementType) == 1:
                        print("-> 'Placement type' has no value, empty.")
                        placementType = ""

                with open(csvFile, "a", encoding="utf-8-sig", newline="") as writeFile:
                    write_csv = csv.writer(writeFile)
                    write_csv.writerow([engName, chiName, licenceSince, district, address, phoneNumber,
                                        faxNumber, email, placementType])

                driver.execute_script("window.history.go(-1)")

                viewDetailXNumber += 1

except Exception as e:
    # Program quits after 2 minutes
    print(e)
    time.sleep(120)
    driver.quit()
    sys.exit()

driver.close()
