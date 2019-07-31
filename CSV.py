from selenium import webdriver

print("Enter the URL you wish to search:")

inp=input()

url = webdriver.Firefox(executable_path='./geckodriver')
 
url.get("https://google.com/")

search = url.find_element_by_name('q')

search.send_keys(inp)

search.submit()
