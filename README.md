# -Web-Scraping-Property-Data-Automation-Specialist-
################# Name :- Mangesh Bhagwanrao Gaigodhane ##################

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from pytesseract import pytesseract 
from bs4 import BeautifulSoup
from PIL import Image 
import time
import csv
import requests
import time
import traceback


# Initialize the WebDriver (assuming you have Chrome installed)
driver = webdriver.Chrome()

# Open the webpage
url = "https://wbregistration.gov.in/(S(havya3y02mwc2xyslq1nuqjd))/index/Search_By_Property_new.aspx?SearchingFrom=WS"
driver.get(url)

time.sleep(10) 

#IM Taking Year 2021
driver.execute_script('document.getElementById("ctl00_CPH_txt_YearFrom").value = "2021";')

# District fields using XPath
district_select = Select(driver.find_element(By.XPATH, '//*[@id="ctl00_CPH_DDL_prop_district"]'))
district_select.select_by_visible_text("Kolkata")
district_select.select_by_value('19')

time.sleep(5) 

# Thana fields using XPath
thana_select = Select(driver.find_element(By.XPATH, '//*[@id="ctl00_CPH_DDL_thana"]'))
thana_select.select_by_visible_text("Amherst Street")
thana_select.select_by_value('03')

time.sleep(5) 

# time.sleep(5) 

driver.execute_script('document.getElementById("ctl00_CPH_txt_YearFrom").value = "2021";')

time.sleep(5) 


wait = WebDriverWait(driver, 20)
# Click on the checkbox to make 'Road' visible
road_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_CPH_chk_road")))
road_checkbox.click()

time.sleep(15)  # Waiting for 5 seconds for the results to load
road_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_CPH_DDL_road")))
desired_road = "A. P. C. Road"
road_dropdown.click()
road_dropdown.send_keys(desired_road)
road_dropdown.send_keys(Keys.RETURN)

time.sleep(5) 


############################################### CAPTCHA Start ##########################################################

captcha_img = driver.find_element(By.ID,'ctl00_CPH_Img_Capcha')
captcha_img.screenshot('E:\Mangesh\Teleglobal\Captcha_SS\captcha12.png')
time.sleep(5) 

# Defining paths to tesseract.exe 
# and the image we would be using 
path_to_tesseract = r'E:\Mangesh\Teleglobal\Tesseract\tesseract.exe'
image_path = r'E:\Mangesh\Teleglobal\Captcha_SS\captcha12.png'
##for Captcha i have use tesseract which is convert img into text i have take spanshot of captcha
#and convert into text
#we have to install it and give it path accordingly

############################################### Captcha Test NO1 ######################################################################
# Opening the image & storing it in an image object 
img = Image.open(image_path) 
  
# location to pytesseract library 
pytesseract.tesseract_cmd = path_to_tesseract 
# This function will extract the text from the image 
text = pytesseract.image_to_string(img) 
text = text[:-1] 
# Displaying the extracted text 
print(text)

#################################################  Captcha Test NO2 ##################################################################
# img = Image.open(image_path)
# img = img.convert('L')  # Convert to grayscale
# # Add more preprocessing steps here if necessary

# # Providing the tesseract executable location
# pytesseract.pytesseract_cmd = path_to_tesseract  # Corrected here

# # Passing the preprocessed image to image_to_string()
# text = pytesseract.image_to_string(img, lang='eng', config='--psm 6')  # PSM 6 for single uniform block of text

# # Displaying the extracted text
# text = text.strip()
# print(text)
time.sleep(5) 
# driver.execute_script('document.getElementById("ctl00_CPH_txtCapcha").value = {text};')
security_code_box = driver.find_element(By.ID, 'ctl00_CPH_txtCapcha')
security_code_box.clear()  # Clear any existing text in the box
security_code_box.send_keys(text)
time.sleep(5) 

############################################### Captcha Ends ################################################################

#Click on Display Button
display_button = driver.find_element(By.ID, 'ctl00_CPH_btn_SubmitQuery')
display_button.click()

#################################################  Table Data  ###############################################################
time.sleep(15) 

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table[@id="ctl00_CPH_GRV_SearchByProperty"]')))

# Find the table element by its ID (replace 'TableID' with the actual ID of the table)
table = driver.find_element(By.XPATH, '//table[@id="ctl00_CPH_GRV_SearchByProperty"]')

# Parse the table data and save it to a CSV file
with open('Table_Data_.csv', 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)

    # Write the header row
    header_row = table.find_element(By.XPATH, './/tr[1]')  # Adjusted to fetch the first row
    header_cells = header_row.find_elements(By.XPATH, './/th')
    header_data = [cell.text.strip() for cell in header_cells]
    writer.writerow(header_data)

    # Write the data rows excluding the last one
    data_rows = table.find_elements(By.XPATH, './/tr[position()<last()]')  # Excludes the last row
    for row in data_rows:
        cells = row.find_elements(By.XPATH, './/td')
        row_data = [cell.text.strip() for cell in cells]
        writer.writerow(row_data)
        print(row_data)

# Click the "View" button in the last column to view the party

# Function to extract and parse table data from the popup window HTML
def extract_and_parse_table_data(driver, csv_writer):
    try:
        popup_html = driver.page_source

        soup = BeautifulSoup(popup_html, 'html.parser')

        # Find the table element using BeautifulSoup methods
        table = soup.find('table', class_='table table-striped table-bordered table-hover')
        if table:
            # Extract headers and write them to CSV file
            headers = [header.text.strip() for header in table.find_all('th')]
            csv_writer.writerow(headers)

            # Extract and write table data to CSV file
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all('td')
                row_data = [cell.text.strip() for cell in cells]
                csv_writer.writerow(row_data)
        else:
            print("Table not found in the popup window.")

        # Close the popup window
        driver.close()

        # Switch back to the original window
        driver.switch_to.window(driver.window_handles[0])

    except Exception as e:
        print("Error occurred while extracting and parsing table data from popup:", e)
        traceback.print_exc()  # Print the stack trace if an exception occurs

# Wait for the table to load
time.sleep(10)

# Open CSV file for writing
with open('POPUP_Table_Data.csv', 'w', newline='', encoding='utf-8') as csvfile:
    csv_writer = csv.writer(csvfile)

    view_buttons = driver.find_elements(By.XPATH, '//table[@id="ctl00_CPH_GRV_SearchByProperty"]//input[@type="submit"]')
    num_buttons = len(view_buttons)
    print(f"Found {num_buttons} View buttons.")

    for i in range(num_buttons):  # Loop through the range of buttons
        try:
            print(f"Clicking View button {i+1}")
            # Re-locate the buttons before clicking to avoid stale element reference
            view_buttons = driver.find_elements(By.XPATH, '//table[@id="ctl00_CPH_GRV_SearchByProperty"]//input[@type="submit"]')
            view_buttons[i].click()
            time.sleep(10)

            # Switch to the new window
            popup_window = driver.window_handles[-1]
            driver.switch_to.window(popup_window)

            # Scrape and write table data to CSV file
            extract_and_parse_table_data(driver, csv_writer)

            # Wait for the original table page to load again
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//table[@id="ctl00_CPH_GRV_SearchByProperty"]')))

        except Exception as e:
            print(f"Error occurred while processing View button {i+1}: {e}")
            traceback.print_exc()  # Print the stack trace if an exception occurs

print("Crawling completed.")

time.sleep(10) 
driver.quit()

#Thanks
