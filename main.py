from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService

# Update the path to your msedgedriver.exe
service = EdgeService(executable_path='C:/Users/robot/Downloads/Programs/zillowscrap/Driver/msedgedriver.exe')

# Pass the Service object to the Edge driver
driver = webdriver.Edge(service=service)
driver.get('https://google.com')
