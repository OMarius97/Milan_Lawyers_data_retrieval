import pandas as pd
import selenium
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time


results = []

# extracting Lawyers name from Excel file
df = pd.read_excel ('Lawyers name.xlsx')
lawyers_name = pd.DataFrame(df, columns =['NAME'])



for i in range(0, len(lawyers_name)):

    s = Service('msedgedriver.exe')
    driver = webdriver.Edge(service=s)
    try:
        # access the website
        driver.get('https://sfera.sferabit.com/servizi/alboonlineBoot/index.php?id=1080')

        # search lawyers by name
        rag_soc = driver.find_element(by=By.NAME, value='filtroRagioneSociale')
        rag_soc.click()
        rag_soc.send_keys(lawyers_name.loc[i])
        cerca = driver.find_element(by=By.XPATH, value='/html/body/form/div[9]/div/button').click()
        time.sleep(4)

        # try to get the address
        try:
            entire_address = driver.find_element(by=By.XPATH,
                                      value='/html/body/div[1]/div/div/div[2]/div[2]/table[2]/tbody/tr/td[4]').text
        except:
            pass

        # Splitting the entire address suitable for final Excel file
        address = entire_address.split(" - ")
        street = address[0]
        zip_code = address[1]
        province = address[2]
        province = province.replace(')', '')
        province = province.split("(")
        city = province[0]
        province_code = province[1]


        driver.find_element(by=By.NAME, value='dettagli').click()
        time.sleep(2)

        # getting the full name
        name = driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[2]/table/tbody/tr[2]/td/b').text
        try:

            # removing AVV. title
            name = name.replace('Avv. ', '')
        except:
            pass

        # filtering the email address for and looking for the pec
        pec = driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[2]/table/tbody/tr[3]/td/a[2]').text
        if "pec" not in pec:
            try:
                pec = driver.find_element(by=By.XPATH,
                                      value='/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[2]/table/tbody/tr[3]/td/a[3]').text
                if "pec" not in pec:
                    try:
                        pec = driver.find_element(by=By.XPATH,
                                      value='/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[2]/table/tbody/tr[3]/td/a[3]').text
                    except:
                        pass
            except:
                pass

        # getting the National ID
        id = driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[2]/table/tbody/tr[3]/td/div').text
        id = id.replace('Codice fiscale: ', '')


        results.append(name)
        results.append(street)
        results.append(pec)
        results.append(zip_code)
        results.append(city)
        results.append(province_code)
        results.append(id)


        rf = pd.DataFrame()
        rf['NAME'] = results[0::7]
        rf['STREET'] = results[1::7]
        rf['PEC'] = results[2::7]
        rf['ZIP'] = results[3::7]
        rf['CITY'] = results[4::7]
        rf['PROVINCE CODE'] = results[5::7]
        rf['ID'] = results[6::7]
        rf.to_excel('results.xlsx', index=False)

    except:
        pass

