from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

ACCESS_URL = ""
USER_LOGIN = ""
USER_PASSWORD = ""
FIELD_TYPE = "Select List (multiple choices)"
FIELD_NAME = "Test Field"
FIELD_DESCRIPTION = "Test Description"
FILE_PATH = r"excel_sheet/field_options.xlsx"

if __name__ == '__main__':
    driver = webdriver.Edge()
    driver.maximize_window()
    driver.get(ACCESS_URL)

    # Submit email address
    driver.find_element(by=By.NAME, value="username").send_keys(USER_LOGIN)
    driver.find_element(by=By.ID, value="login-submit").click()

    # Submit password
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.NAME, "password"))).send_keys(USER_PASSWORD)
    driver.find_element(by=By.ID, value="login-submit").click()

    # Open new Custom Field window
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "css-goggrm"))).click()

    # Select field type
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, f"/html/body/div[21]/div[2]/div[1]/div[2]/div/ol/li[9]/h3[contains(text(),'{FIELD_TYPE}')]"))).click()
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[21]/div[2]/div[2]/button"))).click()

    # Field name
    driver.find_element(by=By.NAME, value="custom-field-name").send_keys(FIELD_NAME)

    # Field name
    driver.find_element(by=By.NAME, value="custom-field-description").send_keys(FIELD_DESCRIPTION)

    df = pd.read_excel(FILE_PATH)
    for option in df.iloc[:, df.shape[1] - 1]:
        # Input option
        driver.find_element(by=By.CLASS_NAME, value="custom-field-options-input").send_keys(option)
        driver.find_element(by=By.ID, value="custom-field-options-add").click()
