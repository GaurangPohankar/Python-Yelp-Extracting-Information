import xlrd
import time
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import re
import csv
from urllib.parse import urlparse

#NA.xlsx
#name = input("Please Enter the File Name:")
name = "NA.xlsx"

#search_offset = input("Please Enter Search Offset:")
search_offset = 3

#time_offset = input("Please Enter Time Offset:")
time_offset = 10


workbook = xlrd.open_workbook(name)
sheet = workbook.sheet_by_index(0)


fields = ['Find', 'Near','Link','Name','Rating','No_of_reviews','Type','Address','Phone','Website','Timings','Price','Health_Score','Hours','Username','Location_of_Reviewer','Rating','Review_Date','Comment']

out_file = open('data.csv','w')
csvwriter = csv.DictWriter(out_file, delimiter=',', fieldnames=fields)
dict_service = {}

dict_service['Find'] = 'Find'
dict_service['Near'] = 'Near'
dict_service['Link'] = 'Link'
dict_service['Name'] = 'Name'
dict_service['Rating'] = 'Rating'
dict_service['No_of_reviews'] = '# of Rating'
dict_service['Type'] = 'Type'
dict_service['Address'] = 'Address'
dict_service['Phone'] = 'Phone'
dict_service['Website'] = 'Website'
dict_service['Timings'] = 'Timings'
dict_service['Price'] = 'Price'
dict_service['Health_Score'] = 'Health_Score'
dict_service['Hours'] = 'Hours'
dict_service['Username'] = 'Username'
dict_service['Location_of_Reviewer'] = 'Location_of_Reviewer'
dict_service['Rating'] = 'Rating'
dict_service['Review_Date'] = 'Review Date'
dict_service['Comment'] = 'Comment'

with open('data.csv', 'a') as csvfile:
     filewriter = csv.DictWriter(csvfile, delimiter=',', fieldnames=fields)
     filewriter.writerow(dict_service)
     csvfile.close()
     #Write row to CSV
     csvwriter.writerow(dict_service)

data = [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

driver = webdriver.Firefox(executable_path='./geckodriver')
driver.set_page_load_timeout(50)    
driver.maximize_window()

log_url = "https://www.google.com/"
lik=[]

for i in range(len(data)):
     try:
          print(data[i][0])
          driver.get(log_url)
          
          outlet_name = data[i+2][1]
          outlet_loc =  data[i+2][2]

          search = driver.find_element_by_name('q')
          search.send_keys('Yelp.com'+' '+outlet_name+' '+ outlet_loc)
          search.submit()
          
          time.sleep(int(time_offset))
          
          page = driver.page_source
          soup = BeautifulSoup(page, "html.parser")
          
          services = soup.find_all('div', {'class': 'r'})
          
          for sermbl in services:
               ser_ho = str(sermbl)
               links = re.findall(r'(http.*?)"', ser_ho)
               o = urlparse(str(links[0]))

               if o.netloc == "www.yelp.com":
                    lik.append(links[0])
                    
          for i in range(search_offset):
               driver.get(lik[i])
               page2 = driver.page_source
               soup2 = BeautifulSoup(page2 , "html.parser")

               try:
                    reviews = soup2.find_all('div', {'class': 'rating-info clearfix'})
                    #finding the rating
                    rating = reviews[0].find_all('img', {'class': 'offscreen'})
                    rating = rating[0]['alt']
                    rating =  ' '.join(rating.split())
                    print(rating)
               except:
                   rating =' ' 

               try:
                    #finding the reviews
                    reviews = reviews[0].find_all('span', {'class': 'review-count rating-qualifier'})
                    total_reviews = reviews[0].text
                    total_reviews =  ' '.join(total_reviews.split())
                    print(total_reviews)
               except:
                    total_reviews= ' '

               try:
                    # which category food
                    category_str_list = soup2.find_all('span', {'class': 'category-str-list'})
                    category_str_list = category_str_list[0].text
                    category_str_list =  ' '.join(category_str_list.split())
                    print(category_str_list)
               except:
                    category_str_list = ' '

               try:
                    # Address
                    mapbox_text = soup2.find_all('div', {'class': 'mapbox-text'})
                    mapbox_text = mapbox_text[0].find_all('div', {'class': 'map-box-address u-space-l4'})
                    Address = mapbox_text[0].text
                    Address =  ' '.join(Address.split())
                    print(Address)
               except:
                    Address= ' '

               try:
                    # Phone   
                    mapbox_text = soup2.find_all('div', {'class': 'mapbox-text'})
                    biz_phone = mapbox_text[0].find_all('span', {'class': 'biz-phone'})
                    biz_phone = biz_phone[0].text
                    biz_phone =  ' '.join(biz_phone.split())
                    print(biz_phone)
               except:
                    biz_phone = ' '

               try:
                    # website   
                    mapbox_text = soup2.find_all('div', {'class': 'mapbox-text'})
                    biz_website = mapbox_text[0].find_all('span', {'class': 'biz-website js-biz-website js-add-url-tagging'})
                    biz_website = biz_website[0].findAll('a')
                    biz_website = biz_website[0].string
                    biz_website =  ' '.join(biz_website.split())
                    print(biz_website)
               except:
                    biz_website = ' ' 

               # other info
               try:
                    island_summary = soup2.find_all('div', {'class': 'island summary'})
               
                    biz_hours = island_summary[0].find_all('li', {'class': 'biz-hours iconed-list-item'})
                    biz_hours = biz_hours[0].text
                    biz_hours =  ' '.join(biz_hours.split())
                    print(biz_hours)

                    price_range = island_summary[0].find_all('dl', {'class': 'short-def-list'})
                    price_range = price_range[1].text
                    price_range =  ' '.join(price_range.split())
                    print(price_range)

                    health_score = island_summary[0].find_all('li', {'class': 'iconed-list-item health-score'})
                    health_score = health_score[0].text
                    health_score =  ' '.join(health_score.split())
                    print(health_score)
               except:
                    biz_hours = ' '
                    price_range = ' '
                    health_score = ' '

               try:   
                    # Hours   
                    ywidget = soup2.find_all('div', {'class': 'ywidget biz-hours'})
                    ywidget = ywidget[0].text
                    #ywidget =  ' '.join(ywidget.split())
                    print(ywidget)
               except:
                    ywidget = ' '

               try:
                    # Cutomer Reviews   
                    ywidget_sidebar = soup2.find_all('div', {'class': 'review review--with-sidebar'})
                    
                    user_name = ywidget_sidebar[1].find_all('li', {'class': 'user-name'})
                    user_name = user_name[0].text
                    user_name =  ' '.join(user_name.split())
                    print(user_name)

                    #location user
                    user_location = ywidget_sidebar[1].find_all('li', {'class': 'user-location responsive-hidden-small'})
                    user_location = user_location[0].text
                    user_location =  ' '.join(user_location.split())
                    print(user_location)

                    #date
                    date = ywidget_sidebar[1].find_all('span', {'class': 'rating-qualifier'})
                    date = date[0].text
                    date =  ' '.join(date.split())
                    print(date)

                    #content
                    content = ywidget_sidebar[1].find_all('p')
                    content = content[0].text
                    content =  ' '.join(content.split())
                    print(content)

                    # rating by user
                    rating_given = ywidget_sidebar[1].find_all('img', {'class': 'offscreen'})
                    rating_given = rating_given[0]['alt']
                    rating_given =  ' '.join(rating_given.split())
                    print(rating_given)
               except:
                    user_name = ' '
                    user_location = ' '
                    date = ' '
                    content = ' '
                    rating_given = ' '

               dict_service['Find'] = outlet_name
               dict_service['Near'] = outlet_loc
               dict_service['Link'] = lik[i]
               dict_service['Name'] = outlet_name
               dict_service['Rating'] = rating
               dict_service['No_of_reviews'] = total_reviews
               dict_service['Address'] = Address
               dict_service['Type'] = category_str_list
               dict_service['Phone'] = biz_phone
               dict_service['Website'] = biz_website
               dict_service['Timings'] = biz_hours
               dict_service['Price'] = price_range
               dict_service['Health_Score'] = health_score
               dict_service['Hours'] = ywidget
               dict_service['Username'] = user_name
               dict_service['Location_of_Reviewer'] = user_location
               dict_service['Rating'] = rating_given
               dict_service['Review_Date'] = date
               dict_service['Comment'] = content
                    
               with open('data.csv', 'a') as csvfile:
                    filewriter = csv.DictWriter(csvfile, delimiter=',', fieldnames=fields)
                    filewriter.writerow(dict_service)
                    csvfile.close()
                    #Write row to CSV
                    csvwriter.writerow(dict_service)
                    
          #resetting array      
          lik=[]
     except:
          print(" ")
     


     
