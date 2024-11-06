from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException,StaleElementReferenceException,ElementClickInterceptedException
import time
import re
import pandas as pd
import datetime as dt
def choose_items(item_cat,
                 item_link,
                 item_size,
                 item_qty):
                 main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
                 child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"shop-home").shadow_root
                 mouterwear = child_root.find_element(By.CSS_SELECTOR,"a[href='"+item_cat+"']")
                 time.sleep(2)
                 mouterwear.click()
                 time.sleep(2)
                 # Outerwear Page and Picking Item
                 main_shadow_root = driver.find_element(By.XPATH,"//shop-app").shadow_root
                 child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"iron-pages")
                 child_root1 = child_root.find_element(By.CSS_SELECTOR,"shop-list").shadow_root
                 choose_shirt1 = child_root1.find_element(By.CSS_SELECTOR,"a[href='"+item_link+"']")
                 choose_shirt1.click()
                 time.sleep(2)
                 # Item detail page, Changing Size and qty
                 main_shadow_root = driver.find_element(By.XPATH,"//shop-app").shadow_root
                 child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"iron-pages")
                 child_root1 = child_root.find_element(By.CSS_SELECTOR,"shop-detail").shadow_root
                 child_root2 = child_root1.find_element(By.ID,"content")
                 child_root3 = child_root2.find_element(By.CLASS_NAME,"detail")
                 child_root4 = child_root3.find_element(By.CLASS_NAME,"pickers")
                 size_select = child_root4.find_element(By.ID,"sizeSelect")
                 size_select.click()
                 time.sleep(2)
                 size_select.send_keys(item_size)
                 time.sleep(2)
                 qty_select = child_root4.find_element(By.ID,"quantitySelect")
                 qty_select.click()
                 time.sleep(2)
                 qty_select.send_keys(item_qty)
                 time.sleep(2)
                 #Cick Add to Cart
                 child_root4 = child_root3.find_element(By.TAG_NAME,"shop-button")
                 child_root4.click()
                 time.sleep(2)
                 #goto Chart
                 child_root5 = main_shadow_root.find_element(By.CSS_SELECTOR,"shop-cart-modal").shadow_root
                 child_root6 = child_root5.find_element(By.CSS_SELECTOR,"a[href='/cart']")
                 child_root6.click()
                 time.sleep(2)
                  #goto Homepage
                 main_shadow_root = driver.find_element(By.CSS_SELECTOR,"shop-app").shadow_root
                 child_root6 = main_shadow_root.find_element(By.CSS_SELECTOR,"app-header")
                 child_root7= child_root6.find_element(By.CSS_SELECTOR,"app-toolbar")
                 child_root8 = child_root7.find_element(By.CSS_SELECTOR,"a[href='/']")
                 child_root8.click()
                 time.sleep(2)
def view_basket():
    main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
    child_root = main_shadow_root.find_element(By.CLASS_NAME,"cart-btn-container")
    child_root1 = child_root.find_element(By.CSS_SELECTOR,"a[href='/cart']")
    child_root1.click()
    time.sleep(3)
def change_qty(item_name, qty):
    main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
    child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"shop-cart.iron-selected").shadow_root
    child_root1 = child_root.find_element(By.CSS_SELECTOR, "div.main-frame")
    child_root2 = child_root1.find_elements(By.TAG_NAME,"shop-cart-item")
    cnt = 0
    for i in child_root2:
        cnt = cnt + 1
        if re.search(item_name,i.text):
            child_root1 = child_root1.find_element(By.CSS_SELECTOR, "div.subsection > div.list")
            main_shadow_root = child_root1.find_element(By.CSS_SELECTOR,"div > div:nth-child(2) > div.list > shop-cart-item:nth-child("+str(cnt)+")").shadow_root
            select = Select(WebDriverWait(main_shadow_root, 10).until(EC.element_to_be_clickable((By.ID, "quantitySelect"))))
            select.select_by_visible_text(qty)
    time.sleep(3)
def check_out():
    try:
        main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
        child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"shop-cart.iron-selected").shadow_root
        child_root.find_element(By.CSS_SELECTOR, "div > div:nth-child(2) > div.checkout-box > shop-button").click()
        return True
    except ElementNotInteractableException:
        main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
        child_root = main_shadow_root.find_element(By.CLASS_NAME,"logo")
        child_root1 = child_root.find_element(By.CSS_SELECTOR,"a[href='/']")
        child_root1.click()
        time.sleep(3)
        return False
def get_checkout_dtls(file_name):
    file_data = pd.read_excel(file_name)
    print(file_data)
    dict = {}
    cnt = 0
    for i in file_data:
        dict[i] = file_data[i][cnt]
    return dict
def delete_from_basket(item_name):
    try:
        main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
        child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"shop-cart.iron-selected").shadow_root
        child_root1 = child_root.find_element(By.CSS_SELECTOR, "div.main-frame")
        child_root2 = child_root1.find_elements(By.TAG_NAME,"shop-cart-item")
        cnt = 0
        for i in child_root2:
            cnt = cnt + 1
            if re.search(item_name,i.text):
                child_root1 = child_root1.find_element(By.CSS_SELECTOR, "div.subsection > div.list")
                main_shadow_root = child_root1.find_element(By.CSS_SELECTOR,"div > div:nth-child(2) > div.list > shop-cart-item:nth-child("+str(cnt)+")").shadow_root
                main_shadow_root.find_element(By.CSS_SELECTOR,"div > div.detail > paper-icon-button").click()
                #select = Select(WebDriverWait(main_shadow_root, 10).until(EC.element_to_be_clickable((By.ID, "quantitySelect"))))
                #select.select_by_visible_text(qty)
        time.sleep(3)
    except StaleElementReferenceException:
        time.sleep(3)
def upload_dtls(dict):
    main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
    child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"shop-checkout.iron-selected").shadow_root
    child_root1 = child_root.find_element(By.CSS_SELECTOR, "div.main-frame")
    for i in dict:
        if i == "Email":
            child_root1.find_element(By.ID,"accountEmail").send_keys(dict[i])
        elif i == "Phone":
            child_root1.find_element(By.ID,"accountPhone").send_keys(int(dict[i]))
        elif i == "Address":
            child_root1.find_element(By.ID,"shipAddress").send_keys(dict[i])
        elif i == "City":
            child_root1.find_element(By.ID,"shipCity").send_keys(dict[i])
        elif i == "State":
            child_root1.find_element(By.ID,"shipState").send_keys(dict[i])
        elif i == "Zip":
            child_root1.find_element(By.ID,"shipZip").send_keys(int(dict[i]))
        elif i == "Country":
            #child_root1.find_elements(By.ID,"shipCountry").send_keys(dict[i])
            select = Select(WebDriverWait(child_root1, 10).until(EC.element_to_be_clickable((By.ID, "shipCountry"))))
            select.select_by_visible_text(dict[i])
        elif i == "CardholderName":
            child_root1.find_element(By.ID,"ccName").send_keys(dict[i])
        elif i == "Card Number":
            child_root1.find_element(By.ID,"ccNumber").send_keys(int(dict[i]))
        elif i == "Expiry":
            select_mth = Select(WebDriverWait(child_root1, 10).until(EC.element_to_be_clickable((By.ID, "ccExpMonth"))))
            select_yr = Select(WebDriverWait(child_root1, 10).until(EC.element_to_be_clickable((By.ID, "ccExpYear"))))
            year = int(str(dict[i]).split()[0].split('-')[0])
            month = int(str(dict[i]).split()[0].split('-')[1])
            day = int(str(dict[i]).split()[0].split('-')[2])
            l_dt = dt.datetime(year,month,day)
            l_month = l_dt.strftime('%b')
            l_year = l_dt.strftime('%Y')
            select_mth.select_by_visible_text(l_month)
            select_yr.select_by_visible_text(l_year)
            #child_root1.find_elements(By.ID,"accountPhone").send_keys(dict[i])
        elif i == "CVV":
            child_root1.find_element(By.ID,"ccCVV").send_keys(int(dict[i]))
    time.sleep(3)
def place_order_and_finish():
    try:
        main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
        child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"shop-checkout.iron-selected").shadow_root
        child_root1 = child_root.find_element(By.CSS_SELECTOR, "div.main-frame")
        child_root1.find_element(By.ID,"submitBox").click()
        time.sleep(3)
        main_shadow_root = driver.find_element(By.XPATH,"//shop-app    ").shadow_root
        child_root = main_shadow_root.find_element(By.CSS_SELECTOR,"shop-checkout.iron-selected").shadow_root
        child_root1 = child_root.find_element(By.CSS_SELECTOR, "div.main-frame")
        child_root1.find_element(By.CSS_SELECTOR,"a[href='/']").click()
    except ElementClickInterceptedException:
        child_root1.find_element(By.CSS_SELECTOR,"a[href='/']").click()
# Initializing the browser
chrome_options = Options()
chrome_options.add_experimental_option("detach",True)
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://shop.polymer-project.org/")
#Choose Items
#Choose Men's wear
outerwear = '/list/mens_outerwear'
item_link = '/detail/mens_outerwear/Men+s+Tech+Shell+Full-Zip'
item_size = 'XL'
item_qty = '2'
choose_items(outerwear,
             item_link,
             item_size,
             item_qty)
#Choose Ladies wear
outerwear = '/list/ladies_outerwear'
item_link = '/detail/ladies_outerwear/Ladies+Modern+Stretch+Full+Zip'
item_size = 'XS'
item_qty = '3'
choose_items(outerwear,
             item_link,
             item_size,
             item_qty)

outerwear = '/list/mens_outerwear'
item_link = '/detail/mens_outerwear/Anvil+L+S+Crew+Neck+-+Grey'
item_size = 'L'
item_qty = '3'
choose_items(outerwear,
            item_link,
            item_size,
            item_qty)

#Viw Basket
view_basket()
#Change qty
item_name = "Ladies Modern Stretch Full Zip"
new_qty = '1'
change_qty(item_name, new_qty)

item_name = "Anvil L/S Crew Neck - Grey"
new_qty = '5'
change_qty(item_name, new_qty)

#Delete item from cart
item_name = "Men's Tech Shell Full-Zip"
delete_from_basket(item_name)

#Checkout
if check_out():
    #Read Excel File
    dict = {}
    dict = get_checkout_dtls("Checkoutdetails.xlsx")

    #Upload details
    upload_dtls(dict)

    #Place Order and Finish
    place_order_and_finish()
