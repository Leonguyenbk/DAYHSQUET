import time, traceback, threading, sys, json, re, os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, ElementClickInterceptedException, JavascriptException,
    StaleElementReferenceException, NoSuchElementException, ElementNotInteractableException,
    NoSuchWindowException
)
import tempfile
from helpers import wait_query_done
from ultils import *


USERNAME = os.getenv("MPLIS_USERNAME", "")
PASSWORD = os.getenv("MPLIS_PASSWORD", "")
URL = "https://dla.mplis.gov.vn/dc/DonDangKy/KeKhaiDangKyV2"
MAXA = "24121"
SOTO = "1"
SOTHUA = "12"
LOAIDAT = "DGT"

options = Options()
options.add_argument("--start-maximized")
options.add_argument("--window-position=100,100")
options.add_argument("--window-size=1400,900")
driver_path = r"D:\Soft\CHROMEDRIVER\chromedriver-win64\chromedriver-win64\chromedriver.exe"
service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

driver.get(URL)
print(f"🌐 Mở trang: {URL}")

if not USERNAME or not PASSWORD:
    raise RuntimeError("Thiếu MPLIS_USERNAME hoặc MPLIS_PASSWORD trong biến môi trường.")

username_box, password_box = get_login_fields(wait)
username_box.send_keys(USERNAME)
password_box.send_keys(PASSWORD)
password_box.send_keys(Keys.ENTER)
print("🔐 Đăng nhập thành công!")
input("Nhấn Enter để tiếp tục…")
wait_query_done(driver)
chon_xa_va_mo_tra_cuu(driver, wait, MAXA)
so_ban_ghi = nhap_to_thua_va_tim_kiem(
    driver=driver,
    wait=wait,
    so_to=SOTO,
    so_thua=SOTHUA,
    timeout=60
)
if so_ban_ghi:
    chon_ban_ghi_dau_tien(driver)
    print(f"Đã chọn bản ghi đầu tiên.")
    wait_query_done(driver)
    modal_ho_so_quet = mo_ho_so_quet(driver, timeout=60)
    chon_ho_so_quet_dau_tien(driver, modal_ho_so_quet, timeout=10)
    modal_add_hsq = cap_nhat_ho_so_quet_dau_tien(
    driver=driver,
    modal_ho_so_quet=modal_ho_so_quet,
    timeout=30
    )
    xoa_file_dau_tien_trong_add_hosoquet(
        driver=driver,
        modal_add_hsq=modal_add_hsq,
        timeout=30
    )
    them_file_don_dang_ky_trong_add_hosoquet(
            driver=driver,
            modal_add_hsq=modal_add_hsq,
            maxa=MAXA,
            loaidat=LOAIDAT,
            timeout=30
        )

else:
    print("⚠️ Không tìm thấy bản ghi nào.")

input("Nhấn Enter để thoát…")
    
driver.quit()

