from bs4 import BeautifulSoup as bs
import requests
import re
import wget
import Variables
from Variables import source_folder

# Defining base domain names at different levels. Domain is just the basic domain name of the website, Domain2 is
# used to easily create the download link from the HREF tag (due to the way the HREF tag is formatted), URL_Full is the
# page to the end use database the main page to be searching outward from
DOMAIN = "https://oee.nrcan.gc.ca"
DOMAIN2 = "https://oee.nrcan.gc.ca/corporate/statistics/neud/dpa"
url_full = "https://oee.nrcan.gc.ca/corporate/statistics/neud/dpa/menus/trends/comprehensive_tables/list.cfm"


# Defining a simple function for grabbing urls
def get_soup(url):  # defines the get_soup function for later use
    return bs(requests.get(url).text, 'html.parser')


# Defines a variable for the main page of the search
response = requests.get(url_full)

# First check if the website is working then proceed to run code to find and download files
if response.status_code == 200:  # This checks if the website is working
    soup = bs(response.text, 'html.parser')  # This will create soup from the main URL (essentially takes all the
    # links and turns them into a txt files to be read at a later date
    for tag in soup.find_all(href=re.compile("/trends")):  # This filters out any links with a /trends in the URL
        # Defining variables for each link with the /trends found
        HREF = tag['href']
        TITLE = tag['title']
        MODTITLE = TITLE.replace("/", "-")  # Replaces / with - because of how file systems work
        FILENAME = MODTITLE + '.zip'  # Append the .zip to the filename otherwise it downloads as a blank file
        URL = DOMAIN2 + HREF[8:]  # Creates the URL for each specific download page the [8:] is there to remove the
        # ../../.. from the HREF code, so it can be appended onto the domain. Should note that this does not lead to
        # the zip link only the page where the link resides
        FileStructure = source_folder + "\\Temp\\" + FILENAME
        print("")
        print("")
        print(URL)
        print(FILENAME)
        print("")  # These prints are here as just checks/information they can be deleted without worry but will make
        # the output log of the exe file less clear
        link = get_soup(URL).find(title='Click here to download all of the tables in this menu').get("href")  # Gets
        # the download link from the variable 'URL' and creates a new variable pointing towards the .zip file
        Down_Link = (DOMAIN + link)  # Joins the download links together to turn relative URL into absolute URL
        wget.download(Down_Link,
                      out=FileStructure)  # Downloads the file and renames it to the FILENAME variable previously
        # specified (should not specify earlier in this python file not from the Variable python file (putting this
        # as a note because eventually variable python file should be imported to share variables across script)
print("")
print('Hopefully Successfully Downloaded All Files From The Comprehensive Energy Use Database')
