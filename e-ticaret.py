from selenium import webdriver
from selenium.webdriver.common.keys import Keys  #Enter işlemi için
from selenium.webdriver.common.by import By      #By sınıfını kullanabilmek için
from selenium.webdriver.support.ui import WebDriverWait  #Bekleme süresini ayarlar
from selenium.webdriver.support import expected_conditions as EC  #sayfanın görünür olmasını bekler
import pandas as pd   #verilen bilgileri dataframe kullanarak excel e yazdırmak için
import xlsxwriter
import time  #time modülünü kullanabilmek için


input_name = input("Item Name: ")  #terminal ekranından girdi alır

Driver = webdriver.Chrome()  #hangi tarayıcıda olacağını sağlar
Driver.maximize_window()  #tarayıcıyı tam ekran yapar hep yapılması uygundur

wait=WebDriverWait(Driver, 10)  #belirtilen süre boyunca sayfanın yüklenmesini bekler


# --Trendyol--
Driver.get("https://www.trendyol.com/")

# arama butonunu bulma ve girdiyi arama-tıklama işlemi
search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div/div/div[2]/div/div/div/input")))
search_button.send_keys(input_name)  #girdiyi arama kutusuna yazar
search_button.send_keys(Keys.ENTER)  #arkaplanda enter tuşuna basarak arama işlemi gerçekleşir 

# ürün adlarını ve fiyatlarını class adları yoluyla alır
item_name_trendyol = wait.until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "prdct-desc-cntnr-name")))
item_price_trendyol = wait.until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "prc-box-dscntd")))

# ürün ad ve fiyatlarını tutabilmek için boş liste tanımladık
trendyol_name_list = []
trendyol_price_list = []

# döngü ile adı ve fiyatları listelere ekler 
for name, price in zip(item_name_trendyol, item_price_trendyol):
    trendyol_name_list.append(name.text)
    price_trendyol = price.text.replace("\n", "").replace("TL", "").replace(".", "").replace(",", ".")
    trendyol_price_list.append(float(price_trendyol))

#oluşacak tabloda site adı, ürün adı ve fiyatının olacağı sütunların alacağı değerleri yerleştirir
df_trendyol = pd.DataFrame({'website': "Trendyol", 'Item Name': trendyol_name_list, 'Item Price': trendyol_price_list})


# --N11--
Driver.get("https://www.n11.com/")

search_bar_n11 = wait.until(EC.visibility_of_element_located((By.ID, "searchData")))
search_bar_n11.send_keys(input_name)
search_bar_n11.send_keys(Keys.ENTER)

item_name_n11 = wait.until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "productName")))
item_price_n11 = wait.until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "newPrice")))

n11_name_list = []
n11_price_list = []

for name, price in zip(item_name_n11, item_price_n11):
    n11_name_list.append(name.text)
    price_n11 = price.text.replace("TL", "").replace(".", "").replace(",", ".")
    n11_price_list.append(float(price_n11))

df_n11 = pd.DataFrame({'website': "n11", 'Item Name': n11_name_list, 'Item Price': n11_price_list})


# --DR--
Driver.get("https://www.dr.com.tr/")

search_bar_DR = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/header/div[2]/div/div/div[4]/div[2]/input")))
search_bar_DR.send_keys(input_name)
search_bar_DR.send_keys(Keys.ENTER)

item_name_DR = wait.until(EC.visibility_of_all_elements_located((By.XPATH, "/html/body/div[1]/div[2]/div/div/main/div[4]/div[1]/div[7]/div/div[2]/div[1]/a")))
item_price_DR = wait.until(EC.visibility_of_all_elements_located((By.CLASS_NAME, "prd-price")))

DR_name_list = []
DR_price_list = []

for name, price in zip(item_name_DR, item_price_DR):
    DR_name_list.append(name.text)
    price_DR = price.text.replace("TL", "").replace(".", "").replace(",", ".").strip()
    DR_price_list.append(float(price_DR))

df_DR = pd.DataFrame({'website': "DR", 'Item Name': DR_name_list, 'Item Price': DR_price_list})

# üç dataframe i birleştirip df adlı dataframe oluşturur ve sütunlarını birleştirerek düzgün bir tablo oluşturur
df = pd.concat([df_trendyol, df_n11, df_DR], ignore_index=True)
df_sorted = df.sort_values("Item Price")  # ürün fiyatlarını azdan çoğa olacak şekilde sıralar

# Excel'e yazma işlemi
writer = pd.ExcelWriter("tablo.xlsx", engine="openpyxl")
df_sorted.to_excel(writer, sheet_name="Sheet1", index=False)

# Excel sayfasına erişim
workbook = writer.book
worksheet = workbook["Sheet1"]

# Hücre genişliklerini ayarlama
worksheet.column_dimensions["B"].width = 30  # İtem Adı sütunu genişliği

# Excel dosyasını kaydet
writer._save()
writer.close()

Driver.quit()
