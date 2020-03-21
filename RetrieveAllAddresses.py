import selenium
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

options = Options()
options.headless = True
profile = webdriver.FirefoxProfile()
driver = webdriver.Firefox(firefox_profile=profile, options=options)
driver.get('http://www.bcad.org/clientdb/propertysearch.aspx?cid=1')

#search_query = input("Would you like to search for an address?  ").lower()
search_term = input('Who do you want to search for?  ')
while search_term != "no":
    driver.find_element_by_id('propertySearchOptions_searchText').send_keys(search_term)
    driver.find_element_by_id('propertySearchOptions_search').click()
    addresses = driver.find_elements_by_class_name('searchResultsShow')
    search_term = input("Would you like to search again? ")
    if search_term.lower == "no":
        break

    print("Possible addresses are as follows: ")
    for address in addresses:
        print(address.text)

driver.quit()