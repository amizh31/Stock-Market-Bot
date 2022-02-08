from time import sleep
from selenium import webdriver
from pyttsx3 import speak

PATH = "C:\\Program Files (x86)\\chromedriver.exe"
driver = webdriver.Chrome(PATH)

driver.get("https://invest.motilaloswal.com/Home/LoginPage")
sleep(5)

username = driver.find_element_by_class_name("form-control")
username.click()
sleep(1)
username.send_keys("8675983445")

password = driver.find_element_by_id("MainPassword")
password.click()
sleep(1)
password.send_keys("ILVJune06")

sleep(1)
enter = driver.find_element_by_id("btnLoginInv")
enter.click()

sleep(5)
watchlist = driver.find_element_by_link_text("Watchlist")
watchlist.click()

sleep(1)
buy = driver.find_element_by_link_text("BUY")
buy.click()

sleep(1)
close = driver.find_element_by_xpath("//*[@id='dvOrderDataForm']/img")
close.click()

sleep(1)
sell = driver.find_element_by_link_text("SELL")
sell.click()

sleep(1)
close = driver.find_element_by_xpath("//*[@id='dvOrderDataForm']/img")
close.click()

sleep(1)
openlogout = driver.find_element_by_xpath("/html/body/header/nav/div/div/div[2]/div[2]/ul/li[5]/a")
openlogout.click()
dropdown = driver.find_element_by_link_text("Logout")
dropdown.click()

speak("Hello Over")
