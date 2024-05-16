from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

from config import USERNAME,PASSWORD              #import config.py having outlook username and password
import pandas as pd

dataframe = pd.read_excel('Client_Details.xlsx')  #importing from an excel sheet with details including client's Names and email-address 
service = Service(executable_path="chromedriver.exe")
driver = webdriver.Chrome(service=service)

driver.maximize_window()
driver.get("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=151&ct=1714709077&rver=7.0.6738.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26cobrandid%3dab0455a0-8d03-46b9-b18b-df2f57b9e44c%26culture%3den-in%26country%3din%26RpsCsrfState%3d9b6b9eed-4f79-207e-0317-e735fb03a48c&id=292841&aadredir=1&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=ab0455a0-8d03-46b9-b18b-df2f57b9e44c")

email_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))
)
email_field.send_keys(USERNAME)

next_btn = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"button[type='submit']"))
)
next_btn.click()

password_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"input[name='passwd']"))
)
password_field.send_keys(PASSWORD)

signin_btn = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"input[type='submit']"))
)
signin_btn.click()

stay_signedin_btn = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"input[type='submit']"))
)
stay_signedin_btn.click()

time.sleep(5)
for i in dataframe.index:
    new_email_btn = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"button[aria-label='New email']"))
    )
    new_email_btn.click()

    to_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"div[aria-label='To']"))
    )
    to_field.send_keys(dataframe.loc[i]['Email'])

    cc_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"div[aria-label='Cc']"))
    )
    cc_field.send_keys("myboss@test.com")
    
    subject_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"input[aria-label='Add a subject']"))
    )
    subject_field.send_keys("Test Notification")

    body_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"div[role='textbox']"))
    )
    body_field.send_keys(f"Hello {dataframe.loc[i]['Name']},\n\nThis is a Test notification\nPurpose of this mail is to test automating bulk emails\n\n")

    send_btn = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,"button[title='Send (Ctrl+Enter)']"))
    )
    send_btn.click()
