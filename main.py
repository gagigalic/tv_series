import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import pandas as pd


options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
service = Service("C:\Development\chromedriver.exe")
driver = webdriver.Chrome(options = options, service=service)
driver.maximize_window()
driver.get("https://www.imdb.com")

select= driver.find_element( by =By.CLASS_NAME, value="ipc-icon--arrow-drop-down")
select.click()

search = driver.find_element(by=By.LINK_TEXT, value="Advanced Search")
search.click()

adv_title_search = driver.find_element(by=By.LINK_TEXT, value="Advanced Title Search")
adv_title_search.click()

tv_series = driver.find_element(by=By.ID, value="title_type-3")
tv_series.click()

min_data  = driver.find_element(by= By.NAME, value = "release_date-min")
min_data.click()
min_data.send_keys("2010")

max_data  = driver.find_element(by= By.NAME, value = "release_date-max")
max_data.click()
max_data.send_keys("2020")

rating_min= driver.find_element(by=By.NAME, value= "user_rating-min")
rating_min.click()
min= Select(rating_min)
min.select_by_visible_text("7.0")

rating_max= driver.find_element(by=By.NAME, value= "user_rating-max")
rating_max.click()
max = Select(rating_max)
max.select_by_visible_text("10")

genres1 = driver.find_element(by=By.ID, value="genres-10")
genres1.click()

genres2 = driver.find_element(by=By.ID, value="genres-17")
genres2.click()

genres3 = driver.find_element(by=By.ID, value="genres-21")
genres3.click()

color= driver.find_element(by=By.ID, value="colors-1")
color.click()

language = driver.find_element(by=By.NAME, value="languages")
lng = Select(language)
lng.select_by_visible_text("English")

results = driver.find_element(by=By.ID, value="search-count")
results.click()
res=Select(results)
res.select_by_visible_text("250 per page")

submit = driver.find_element(by=By.XPATH, value="(//button[@type='submit'])[2]")
submit.click()

current_url = driver.current_url

response = requests.get(current_url)
soup = BeautifulSoup(response.content, "html.parser")

list = soup.find_all("div", {"class": "lister-item"})

series_title = [series.find("h3").find("a").get_text() for series in list]
series_year = [series.find("h3").find("span", {"class":"lister-item-year"}).get_text().replace("(", "").replace(")", "") for series in list]
series_genre = [series.find("span", {"class":"genre"}).get_text().strip() for series in list]
series_reating = [series.find("div", {"class":"ratings-imdb-rating"}).get_text().strip() for series in list]

Series= pd.DataFrame({ "Series title": series_title,
                      "Series year": series_year,
                      "Series genre": series_genre,
                       "Series reating": series_reating})

data_exel= pd.ExcelWriter("Data.xlsx")
Series.to_excel(data_exel)
data_exel.save()