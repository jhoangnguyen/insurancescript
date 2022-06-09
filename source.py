from unittest import skip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

from webdriver_manager.chrome import ChromeDriverManager
from pathlib import Path
import csv
import os
import time

from bs4 import BeautifulSoup
from urllib.request import urlopen
import webbrowser

def crawl(driver):

    
    wait = WebDriverWait(driver, 10)
    driver.set_page_load_timeout(10)

    try:
        driver.get("http://mienthithucvk.mofa.gov.vn/%C4%90%C4%83ngk%C3%BD/Khaitr%E1%BB%B1ctuy%E1%BA%BFn/tabid/104/VE4NCommand/introduction/Default.aspx")
    except TimeoutError:
        driver.execute_script("window.stop();")
    init = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_btAccept")
    init.click()


    time.sleep(3)

    info = csvOpen()

    vnContact = "1966 Aborn Road, San Jose, CA, USA"
    vnPhone = "4082389840"

    allCountries = {
        "Afghanistan" : "AFG",
        "Åland Islands" : "ALA",
        "Albania": "ALB",
        "Algeria": "DZA",
        "American Samoa": "ASM" ,
        "Andorra": "AND",
        "Angola": "AGO",
        "Anguilla": "AIA",
        "Antarctica": "ATA",
        "Antigua and Barbuda": "ATG" ,
        "Argentina": "ARG",
        "Armenia": "ARM",
        "Austria": "AUT",
        "Aruba": "ABW",
        "Australia": "AUS",
        "Azerbaijan": "AZE",
        "Bahamas": "BHS",
        "Bahrain": "BBHR",
        "Bangladesh": "BGD",
        "Barbados": "BRB",
        "Belarus": "BLR",
        "Belgium": "BEK",
        "Bermuda": "BMU",
        "Belize": "BLZ",
        "Benin": "BEN",
        "Bhutan": "BTN",
        "Bolivia": "BOL",
        "Bosnia and Herzegovina": "BIH",
        "Botswana": "BWA",
        "Bouvet Island": "BVT",
        "Brazil": "BRA",
        "British Indian Ocean Territory": "IOT",
        "Brunei Darussalam": "BRM",
        "Bulgaria": "BGR",
        "Burkina Faso": "BFA",
        "Burundi": "BDI",
        "Cambodia": "KHM",
        "Cameroon": "CMR",
        "Canada": "CAN",
        "Cape Verde": "CPV",
        "Cayman Islands": "CYM",
        "Central African Republic": "CAF",
        "Chad": "TCD",
        "Chile": "CHL",
        "China": "CHN",
        "Christmas Island": "CXR" ,
        "Cocos (Keeling) Islands": "CCK" ,
        "Colombia": "COL",
        "Comoros": "COM",
        "Congo": "COG",
        "Congo,": "COD",
        "Cook Islands": "COK" ,
        "Costa Rica": "CRI",
        "Côte d'Ivoire": "CIV",
        "Croatia": "HRV",
        "Cuba": "CUB",
        "Cyprus": "CYP",
        "Czech Republic": "CZE" ,
        "Denmark": "DNK",
        "Djibouti": "DJI",
        "Dominica": "DMA",
        "Dominican Republic": "DOM" ,
        "Ecuador": "ECU",
        "Egypt": "EGY",
        "El Salvador": "SLV" ,
        "Equatorial Guinea": "GNQ" ,
        "Eritrea": "REI",
        "Estonia": "EST",
        "Ethiopia": "ETH",
        "Falkland Islands (Malvinas)": "FLK",
        "Faroe Islands": "FRO",
        "Fiji": "FJI",
        "Finland": "FIN",
        "France": "FRA",
        "French Guiana": "GUF" ,
        "French Polynesia": "PYF",
        "French Southern Territories": "ATF" ,
        "Gabon": "GAB",
        "Gambia": "GMB",
        "Georgia": "GEO",
        "Germany": "DEU",
        "Ghana": "GHA",
        "Gibraltar": "GIB",
        "Greece": "GRC",
        "Greenland": "GRL",
        "Grenada": "GRD",
        "Guadeloupe": "GLP",
        "Guam": "GUM",
        "Guatemala": "GTM" ,
        "Guernsey": "GGY",
        "Guinea": "JIN",
        "Guinea-Bissau": "GNB",
        "Guyana": "GUY",
        "Haiti": "HTI",
        "Heard Island and McDonald Islands": "HMD",
        "Holy See (Vatican City State)": "VAT",
        "Honduras": "HND",
        "Hong Kong": "HKG",
        "Hungary": "HUN",
        "Iceland": "ISL",
        "India": "IND",
        "Indonesia": "IDN",
        "Iran,": "IRN",
        "Islamic Republic of Iraq": "IRQ" ,
        "Ireland": "IRL",
        "Isle of Man": "IMN",
        "Israel": "ISR",
        "Italy": "ITA",
        "Jamaica": "JAM",
        "Japan": "JPN",
        "Jersey": "JEY",
        "Jordan": "JOR",
        "Kazakhstan": "KAZ",
        "Kenya": "KEN",
        "Kiribati": "KIR",
        "Korea,": "KOR",
        "Democratic People's Republic of Korea,": "PRK",
        "Republic of Kuwait": "KWT",
        "Kyrgyzstan": "KGZ",
        "Lao People's Democratic Republic": "LAO",
        "Latvia": "LVA",
        "Lebanon": "LBN",
        "Lesotho": "LSO",
        "Liberia": "LBR",
        "Libyan Arab Jamahiriya": "LBY",
        "Liechtenstein": "LIE",
        "Lithuania": "LTU",
        "Luxembourg": "LUX",
        "Macao": "MAC",
        "Macedonia, the former Yugoslav Republic of": "MKD",
        "Madagascar": "MDG",
        "Malawi": "MWI",
        "Malaysia": "MYS",
        "Maldives": "MDV",
        "Mali": "MLI",
        "Malta": "MLT",
        "Marshall Islands": "MHL",
        "Martinique": "MTQ",
        "Mauritania": "MRT",
        "Mauritius": "MUS",
        "Mayotte": "MYT",
        "Mexico": "MEX",
        "Micronesia, Federated States of": "FSM",
        "Moldova, Republic of": "MDA",
        "Monaco": "MCO",
        "Mongolia": "MNG",
        "Montenegro": "MNE",
        "Montserrat": "MSR",
        "Morocco": "MAR",
        "Mozambique": "MOZ",
        "Myanmar": "MMR",
        "Namibia": "NAM",
        "Nauru": "NRU",
        "Nepal": "NPL",
        "Netherlands": "NLD" ,
        "Netherlands Antilles": "ANT",
        "New Caledonia": "NCL",
        "New Zealand": "NZL",
        "Nicaragua": "NIC",
        "Niger": "NER",
        "Nigeria": "NGA",
        "Niue": "NIU",
        "Norfolk Island": "NFK" ,
        "Northern Mariana Islands": "MNP" ,
        "Norway": "NOR",
        "Oman": "OMN",
        "Pakistan": "PAK",
        "Palau": "PLW",
        "Palestinian Territory, Occupied": "PSE",
        "Panama": "PAN",
        "Papua New Guinea": "PNG" ,
        "Paraguay": "PRY",
        "Peru": "PER",
        "Philippines": "PHL" ,
        "Pitcairn": "PCN",
        "Poland": "POL",
        "Portugal": "PRT",
        "Puerto Rico": "PRI",
        "Qatar": "QAT",
        "Réunion": "REU",
        "Romania": "ROU",
        "Russian Federation": "RUS" ,
        "Rwanda": "RWA",
        "Saint Helena": "SHN" ,
        "Saint Kitts and Nevis": "KNA" ,
        "Saint Lucia": "LCA",
        "Saint Pierre and Miquelon": "SPM" ,
        "Saint Vincent and the Grenadines": "VCT" ,
        "Samoa": "WSM",
        "San Marino": "SMR",
        "Sao Tome and Principe": "STP" ,
        "Saudi Arabia": "SAU",
        "Senegal": "SEN",
        "Serbia": "SRB",
        "Seychelles": "SYC",
        "Sierra Leone": "SLE",
        "Singapore": "SGP",
        "Slovakia": "SVK",
        "Slovenia": "SVN",
        "Solomon Islands": "SLB",
        "Somalia": "SOM",
        "South Africa": "ZAF",
        "South Georgia and the South Sandwich Islands": "SGS" ,
        "Spain": "ESP",
        "Sri Lanka": "LKA",
        "Sudan": "SDN",
        "Suriname": "SUR",
        "Svalbard and Jan Mayen": "SJM",
        "Swaziland": "SWZ",
        "Sweden": "SWE",
        "Switzerland": "CHE",
        "Syrian Arab Republic": "SYR",
        "Taiwan": "TWN",
        "Tajikistan": "TJK",
        "Tanzania, United Republic of": "TZA",
        "Thailand": "THA",
        "Timor-Leste": "TLS",
        "Togo": "TGO",
        "Tokelau": "TKL",
        "Tonga": "TON",
        "Trinidad and Tobago": "TTO",
        "Tunisia": "TUN",
        "Turkey": "TUR",
        "Turkmenistan": "TKM",
        "Turks and Caicos Islands": "TCA",
        "Tuvalu": "TUV",
        "Uganda": "UGA",
        "Ukraine": "UKR",
        "United Arab Emirates": "ARE",
        "United Kingdom": "GBR",
        "United States": "USA",
        "United States Minor Outlying Islands": "UMI",
        "Uruguay": "URY",
        "Uzbekistan": "UZB",
        "Vanuatu": "VUT",
        "Venezuela": "VEN",
        "Viet Nam": "VNM",
        "Virgin Islands, British": "VGB",
        "Virgin Islands, U.S.": "VIR",
        "Wallis and Futuna": "WLF",
        "Western Sahara" : "ESH",
        "Yemen": "YEM",
        "Zambia" : "ZMB",
        "Zimbabwe" : "ZWE",
    }

    for i in range(len(info)):
        timeStamp = info[i][0]
        address = info[i][11]
        addressBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtAddressAtPresent")
        addressBox.send_keys(address)
        gender = info[i][1]
        firstName = info[i][2]
        middleName = info[i][3]
        lastName = info[i][4]
        passport = info[i][5]
        dob = info[i][6]
        parseDOB = dob.split("-")
        dobYr = parseDOB[0]
        dobMo = parseDOB[1]
        dobDy = parseDOB[2]
        doe = info[i][7]
        parseDOE = doe.split("-")
        doeYr = parseDOE[0]
        doeMo = parseDOE[1]
        doeDy = parseDOE[2]
        doecomplete = "/" + doeDy + doeMo + doeYr
        birthPlace = info[i][8]


        birthPlaceBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtPlaceOfBirth")
        birthPlaceBox.send_keys(birthPlace)


        nationality = info[i][9]
        dropDownBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_ddlNationalityAtPresent")
        dropButton = Select(dropDownBox)
        if nationality in allCountries:
            dropButton.select_by_value(allCountries[nationality])



        occupation = info[i][10]
        # address = info[i][11]
        # addressBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtAddressAtPresent")
        # addressBox.send_keys(address)
        phoneNumber = info[i][12]


        if gender == "Female":
            genderBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_chkSex_Female")
            genderBox.click()

        fNameBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtFirstName")
        fNameBox.send_keys(firstName)
        fVnFNameBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtFirstNameVN")
        fVnFNameBox.send_keys(firstName)
        fVnMNameBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtMiddleNameVN")
        fVnMNameBox.send_keys(middleName)
        fVnLNameBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtLastNameVN")
        fVnLNameBox.send_keys(lastName)
        mNameBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtMiddleName")
        mNameBox.send_keys(middleName)
        lNameBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtLastName")
        lNameBox.send_keys(lastName)
        # if nationality in allCountries:
        #     dropButton.select_by_value(allCountries[nationality])
        passportBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtPassportNo")
        passportBox.send_keys(passport)
        dobDyBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtDayOfBirth")
        dobDyBox.send_keys(dobDy)
        dobMoBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtMonthOfBirth")
        dobMoBox.send_keys(dobMo)
        dobYrBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtYearOfBirth")
        dobYrBox.send_keys(dobYr)
        doeBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_diPassportExpireDate_txtDateInput")
        doeBox.send_keys(doecomplete)








        currNationalityBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_ddlNationalityAtPresent")
        checkBox = Select(currNationalityBox)
        checkBox.select_by_value("USA")
        occupationBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtOccupation")
        occupationBox.send_keys(occupation)
        
        phoneNumBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtPhoneNumber")
        phoneNumBox.send_keys(phoneNumber)
        vnContactBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtContactAddressInVN")
        vnContactBox.send_keys(vnContact)
        vnPhoneBox = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_txtContactPhoneNumber")
        vnPhoneBox.send_keys(vnPhone)
        issuingAuth = driver.find_element(by=By.ID, value = "dnn_ctr408_Desktop_subctr_txtPassportIssuingAuthority")
        issuingAuth.send_keys("USA")

        other = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_cblDocuments1_12")
        other.click()
        other1 = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_chkPrintType_2")
        other1.click()
        other2 = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_ddlCountry")
        other2Button = Select(other2)
        other2Button.select_by_value("USA")
        confirm = driver.find_element(by=By.ID, value="chkConfirm")
        confirm.click()
        submit = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_cmdFinish")
        submit.click()

        time.sleep(12)

        latest_downloaded_file_renamed(lastName, middleName, firstName)

        time.sleep(12)


        try:
            driver.get("http://mienthithucvk.mofa.gov.vn/%C4%90%C4%83ngk%C3%BD/Khaitr%E1%BB%B1ctuy%E1%BA%BFn/tabid/104/VE4NCommand/introduction/Default.aspx")
        except TimeoutError:
            driver.execute_script("window.stop();")

        init = driver.find_element(by=By.ID, value="dnn_ctr408_Desktop_subctr_btAccept")
        init.click()

def latest_downloaded_file_renamed(lastName, middleName, firstName):
    path = str(Path.home() / "Downloads")
    os.chdir(path)
    files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
    recent = files[-1]
    os.rename(path + "\\" + recent, path + "\\" + lastName + middleName + firstName + ".pdf")



def csvOpen():
    file = open("Visa Exemption Form.csv")
    csvRead = csv.reader(file, delimiter=",", quotechar='"')

    header = []
    header = next(csvRead)

    rows = []
    for row in csvRead:
        rows.append(row)
    
    file.close()

    return rows





def main():
    pass

if __name__ == '__main__':
    driver = webdriver.Chrome(ChromeDriverManager().install())
    crawl(driver)
    driver.quit()

    print("done") 

    main()


