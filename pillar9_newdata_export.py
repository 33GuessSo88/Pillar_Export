"""
pillar9_scrape.pyn  
Open pillar 9 site with credentials and download recent sales data.
"""

#! python3
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import WebDriverException
import time
import shutil
import datetime
import keyring

# Pillar 9 credentials using keyring
username = "cbirchch"
password = keyring.get_password("pillar9", "cbirchch")

# Initialize the chrome driver
driver = webdriver.Chrome(
    r"C:\Users\Chris\Documents\Python Scripts\Web Scraping\Real Estate\Pillar9TWO\virenv1\chromedriver.exe"
)
time.sleep(1)
driver.maximize_window()

# Go to Pillar 9 login page
driver.get("https://pillarnine.clareityiam.net/idp/login")
time.sleep(1)
# Find username/email field and send the username from above to the input
driver.find_element_by_id("clareity").send_keys(username)
time.sleep(1)

# Find password input field and insert password
# Field needs input twice sometimes, not sure why?
driver.find_element_by_id("security").send_keys(password)
time.sleep(1)
driver.find_element_by_id("security").send_keys(password)
time.sleep(1)

# Click login button
driver.find_element_by_id("loginbtn").click()
time.sleep(1)

# Click on Matrix MLS System button
pillar9_window = driver.current_window_handle  # Store this window
# print("Pillar9 window is " + pillar9_window)
target = driver.find_element_by_xpath(
    '//*[@id="dashboard"]/div[2]/div/div[1]/div/div[1]/a'
)
target.click()
time.sleep(1)

# Get the child window handle and switch to it
child_window_handle = driver.window_handles
for w in child_window_handle:
    # switch focus to child window
    if w != pillar9_window:
        driver.switch_to.window(w)
        break
time.sleep(1)
# print("Child window title: " + driver.title)

# Find read later button and click until error is thrown
read_later = driver.find_element_by_xpath('//*[@id="NewsDetailPostpone"]')
while read_later.is_displayed():
    try:
        read_later.click()
        time.sleep(1)
    except WebDriverException:
        # print("not clickable")
        time.sleep(1)
else:
    target_8days = driver.find_element_by_link_text("Calgary SOLD last 8 days")
    target_8days.click()
    time.sleep(1)

    # Click on All button to select all sales
    all_button = driver.find_element_by_id("m_lnkCheckAllLink")
    all_button.click()
    time.sleep(1)

    # click Export button
    export_button = driver.find_element_by_id("m_lbExport")
    export_button.click()
    time.sleep(1)

    # Select file format from dropdown
    selection = Select(driver.find_element_by_id("m_ddExport"))
    selection.select_by_visible_text("EXPORT1")
    time.sleep(1)

    # Click Export button
    export_button = driver.find_element_by_id("m_btnExport")
    export_button.click()
    time.sleep(5)

    # Quit and close the intire driver, all windows
    driver.quit()

# Rename export file in downloads folder, add datetime
current_date = datetime.datetime.today().strftime("%a-%b-%d-%YT%H:%M:%S")
shutil.move(
    r"C:\Users\Chris\Downloads\EXPORT1.csv",
    r"C:\Users\Chris\Documents\Real Estate\PowerBI Project\Matrix Sold Data Dumps\Matrix_Dump "
    + str(current_date)
    + ".csv",
)
