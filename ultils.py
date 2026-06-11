from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

from helpers import wait_query_done

# ====== HỖ TRỢ ĐĂNG NHẬP ======
def get_login_fields(wait):
    username_box = wait.until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "input[autocomplete='username'], input[name='username']",
            )
        )
    )
    password_box = wait.until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "input[autocomplete='current-password'], input[name='password']",
            )
        )
    )
    return username_box, password_box

# ====== PHẦN KÊ KHAI ĐĂNG KÝ ======
def wait_tracuu_module_ready(driver, timeout=60):
    """
    Chờ module tra cứu đơn đăng ký load xong.
    """

    # Chờ phần tử xuất hiện trong DOM
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#donDangKyTraCuuModule"))
    )

    # Chờ nó hiển thị
    WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#donDangKyTraCuuModule"))
    )

    # Chờ overlay loading biến mất
    loading_selectors = [
        ".jquery-loading-modal__bg",
        ".jquery-loading-modal_bg",
        "div.jquery-loading-modal_bg",
        "div.jquery-loading-modal__bg",
    ]

    for selector in loading_selectors:
        try:
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, selector))
            )
        except:
            pass

    # Chờ module render xong
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("""
            let el = document.querySelector("#donDangKyTraCuuModule");
            if (!el) return false;
            return el.offsetHeight > 0 && el.offsetWidth > 0;
        """)
    )

    print("✅ Module tra cứu (#donDangKyTraCuuModule) đã load xong!")
    return True

def chon_xa_va_mo_tra_cuu(driver, wait, ma_xa, logger=None):
    """
    Sau khi đăng nhập:
    1. Chọn xã theo mã xã
    2. Mở modal tra cứu đơn đăng ký
    3. Chờ module tra cứu load xong
    """

    def log(msg):
        if logger:
            logger.log(msg)
        else:
            print(msg)

    log(f"✅ Đăng nhập thành công. Bắt đầu chọn xã có mã: {ma_xa}")

    # Chờ combobox xã xuất hiện
    select_xa = wait.until(
        EC.presence_of_element_located((By.ID, "ddlPhuongXaKeKhai"))
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        select_xa
    )

    # Chọn xã bằng Select cho chắc
    Select(select_xa).select_by_value(ma_xa)

    # Bắn sự kiện change để web nhận
    driver.execute_script("""
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, select_xa)

    log(f"✅ Đã chọn xã có mã: {ma_xa}.")

    # Mở modal tra cứu
    log("🔎 Mở cửa sổ tra cứu…")

    tra_cuu_button = wait.until(
        EC.element_to_be_clickable((By.ID, "btnChonDonDangKy"))
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        tra_cuu_button
    )

    driver.execute_script("arguments[0].click();", tra_cuu_button)

    log("✅ Đã bấm nút mở cửa sổ tra cứu.")

    # Chờ module tra cứu load xong
    wait_tracuu_module_ready(driver, timeout=60)

    log("✅ Cửa sổ tra cứu đã sẵn sàng.")

    return True

def wait_tracuu_section_ready(driver, timeout=60):
    selector = "#donDangKyTraCuuModule > div.panel-body > div > div:nth-child(3)"

    # 1) Chờ xuất hiện trong DOM
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
    )

    # 2) Chờ nó visible thật sự
    WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
    )

    # 3) Chờ overlay biến mất (nếu có)
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".jquery-loading-modal__bg"))
        )
    except:
        pass

    # 4) Chờ height/width > 0 (DOM render xong)
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("""
            let el = document.querySelector(arguments[0]);
            if (!el) return false;
            let rect = el.getBoundingClientRect();
            return rect.width > 0 && rect.height > 0;
        """, selector)
    )

    print("✅ Vùng tra cứu (div:nth-child(3)) đã load xong!")

def wait_and_count_tblTraCuu(driver, timeout=60):
    table_selector = "#tblTraCuuTinhHinhDangKy"

    # 1) Chờ bảng xuất hiện
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, table_selector))
    )

    # 2) Chờ overlay MPLIS biến mất
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".jquery-loading-modal__bg"))
        )
    except:
        pass

    # 3) Chờ DataTables ngừng processing
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("""
            let p = document.querySelector("#tblTraCuuTinhHinhDangKy_processing");
            if (p && p.offsetParent !== null) return false;  // đang loading
            return true;
        """)
    )

    # 4) Chờ tbody xuất hiện
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("""
            let tb = document.querySelector("#tblTraCuuTinhHinhDangKy tbody");
            return tb && tb.children.length >= 0;
        """)
    )

    # 5) Đếm số bản ghi thật
    count = driver.execute_script("""
        let table = document.querySelector("#tblTraCuuTinhHinhDangKy");
        if (!table) return -1;

        let rows = table.querySelectorAll("tbody tr");
        if (!rows || rows.length === 0) return 0;

        let count = 0;
        rows.forEach(r => {
            let td = r.querySelector("td");
            if (td && td.classList.contains("dataTables_empty")) return; 
            count++;
        });

        return count;
    """)

    print("➡️ Số bản ghi:", count)
    return count

def nhap_to_thua_va_tim_kiem(driver, wait, so_to, so_thua, timeout=60):
    """
    Nhập Số tờ, Số thửa trong modal Tra cứu đơn đăng ký
    và thực hiện tìm kiếm.

    Return:
        so_ban_ghi: số bản ghi tìm được
        False: nếu lỗi hoặc không tìm thấy bản ghi
    """

    try:
        # --- Tìm modal Tra cứu đơn đăng ký đang mở ---
        modal = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((
                By.CSS_SELECTOR,
                "div[id^='mdlTraCuuDonDangKy-'].modal.in, "
                "div[id^='mdlTraCuuDonDangKy-'].modal.show"
            ))
        )

        # --- Tìm vùng tra cứu trong modal ---
        tra_cuu_box = modal.find_element(
            By.CSS_SELECTOR,
            "#dvTraCuuTinhHinhDangKyChiTiet"
        )

        # --- Tìm ô Số thửa ---
        so_thua_input = WebDriverWait(driver, timeout).until(
            lambda d: tra_cuu_box.find_element(
                By.CSS_SELECTOR,
                "input[name='soThuTuThua']"
            )
        )

        # --- Tìm ô Số tờ ---
        so_to_input = WebDriverWait(driver, timeout).until(
            lambda d: tra_cuu_box.find_element(
                By.CSS_SELECTOR,
                "input[name='soHieuToBanDo']"
            )
        )

        # --- Cuộn tới vùng nhập ---
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            so_thua_input
        )

        # --- Clear và nhập Số thửa ---
        so_thua_input.click()
        so_thua_input.send_keys(Keys.CONTROL, "a")
        so_thua_input.send_keys(Keys.BACKSPACE)
        so_thua_input.send_keys(str(so_thua))

        # --- Clear và nhập Số tờ ---
        so_to_input.click()
        so_to_input.send_keys(Keys.CONTROL, "a")
        so_to_input.send_keys(Keys.BACKSPACE)
        so_to_input.send_keys(str(so_to))

        # --- Nhấn Enter để tìm ---
        so_thua_input.send_keys(Keys.ENTER)

        # --- Chờ query/load xong lần 1 ---
        wait_query_done(driver, timeout=timeout)

        # --- Chờ bảng kết quả xuất hiện ---
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((
                By.ID,
                "tblTraCuuTinhHinhDangKy_info"
            ))
        )

        # --- Đếm bản ghi ---
        so_ban_ghi = wait_and_count_tblTraCuu(driver)

        print(
            f"✅ Đã nhập Số tờ: {so_to}, Số thửa: {so_thua}. "
            f"Số bản ghi tìm được: {so_ban_ghi}."
        )

        if so_ban_ghi == 0:
            print("❌ Không tìm thấy bản ghi nào. Tìm thửa tiếp theo...")
            return False

        return so_ban_ghi

    except Exception as e:
        print(f"❌ Lỗi khi nhập Số tờ/Số thửa và tìm kiếm: {e}")
        return False
    
# === Chọn bản ghi đầu tiên nếu có nhiều hơn 1 bản ghi ===
def chon_ban_ghi_dau_tien(driver, timeout=30):
    wait = WebDriverWait(driver, timeout)

    # 1. Chờ có ít nhất 1 dòng trong bảng
    first_row = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "#tblTraCuuTinhHinhDangKy tbody tr")
        )
    )

    # Trường hợp không có bản ghi nào
    if "Không tìm thấy" in first_row.text:
        return False

    # 2. Tìm ô checkbox
    checkbox = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#tblTraCuuTinhHinhDangKy tbody tr:nth-child(1) td.select-checkbox")
        )
    )

    checkbox.click()

    # 3. Chờ DataTables thêm class 'selected'
    wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "#tblTraCuuTinhHinhDangKy tbody tr.selected")
        )
    )

    # 4. Nhấn nút "Chọn"
    btn_chon = wait.until(
        EC.element_to_be_clickable((By.ID, "btnLuuChonTinhHinhDangKy"))
    )
    btn_chon.click()

    # 5. Chờ modal đóng (panel ẩn đi)
    wait.until(
        EC.invisibility_of_element_located((By.ID, "donDangKyTraCuuModule"))
    )

    return True

# == MỞ HỒ SƠ QUÉT ==
def mo_ho_so_quet(driver, timeout=60):
    try:
        wait = WebDriverWait(driver, timeout)

        btn_ho_so_quet = wait.until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR,
                "#updateDonDangKyModule #btnHoSoQuet"
            ))
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            btn_ho_so_quet
        )

        try:
            wait.until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "#updateDonDangKyModule #btnHoSoQuet"
                ))
            ).click()
        except:
            driver.execute_script("arguments[0].click();", btn_ho_so_quet)

        print("✅ Đã nhấn nút Hồ sơ quét.")

        modal_ho_so_quet = wait.until(
            EC.visibility_of_element_located((
                By.CSS_SELECTOR,
                "div[id^='mdlHoSoQuet-'].in, "
                "div[id^='mdlHoSoQuet-'].show, "
                "div[id^='mdlHoSoQuet-'].modal.in, "
                "div[id^='mdlHoSoQuet-'].modal.show"
            ))
        )

        wait.until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR,
                "div[id^='mdlHoSoQuet-'] #lstHoSoQuet"
            ))
        )

        print("✅ Modal Hồ sơ quét đã mở và danh sách hồ sơ quét đã load.")
        return modal_ho_so_quet

    except Exception as e:
        print(f"❌ Lỗi khi mở Hồ sơ quét: {e}")
        return False

# == CHỌN HỒ SƠ QUÉT ĐẦU TIÊN NẾU CÓ NHIỀU HỒ SƠ ==
def chon_ho_so_quet_dau_tien(driver, modal_ho_so_quet, timeout=10):
    """
    Kiểm tra đơn đăng ký đầu tiên đã selected chưa.
    Nếu rồi thì chọn hồ sơ quét đầu tiên cho đến khi selected.

    Return:
        True: chọn thành công
        False: lỗi hoặc chưa selected
    """

    try:
        wait = WebDriverWait(driver, timeout)

        # --- Lấy ul đầu tiên trong danh sách đơn đăng ký ---
        ul_don_dang_ky = wait.until(
            lambda d: modal_ho_so_quet.find_element(
                By.CSS_SELECTOR,
                "#vModuleDanhSachDangKy #lstDonDangKy ul.vbd-search-item"
            )
        )

        class_don = ul_don_dang_ky.get_attribute("class") or ""

        if "selected" not in class_don.split():
            print("❌ Đơn đăng ký đầu tiên chưa selected.")
            return False

        print("✅ Đơn đăng ký đầu tiên đã selected.")

        # --- Lấy ul đầu tiên trong danh sách hồ sơ quét ---
        ul_ho_so_quet = wait.until(
            lambda d: modal_ho_so_quet.find_element(
                By.CSS_SELECTOR,
                "#vModuleHoSoQuetChon #lstHoSoQuet ul.hosoquet-item"
            )
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            ul_ho_so_quet
        )

        try:
            ul_ho_so_quet.click()
        except:
            driver.execute_script("arguments[0].click();", ul_ho_so_quet)

        print("✅ Đã click hồ sơ quét đầu tiên.")

        # --- Chờ ul hồ sơ quét đầu tiên có class selected ---
        wait.until(
            lambda d: "selected" in (
                modal_ho_so_quet.find_element(
                    By.CSS_SELECTOR,
                    "#vModuleHoSoQuetChon #lstHoSoQuet ul.hosoquet-item"
                ).get_attribute("class") or ""
            ).split()
        )

        print("✅ Hồ sơ quét đầu tiên đã selected.")
        return True

    except Exception as e:
        print(f"❌ Lỗi khi chọn hồ sơ quét đầu tiên: {e}")
        return False
    
# === CHỈNH SỬA HỒ SƠ QUÉT ===
def cap_nhat_ho_so_quet_dau_tien(driver, modal_ho_so_quet, timeout=30):
    """
    Trong modal Hồ sơ quét:
    - Tìm ul hồ sơ quét đầu tiên đang selected
    - Click nút Cập nhật .btnUpdate trong li.actions
    - Chờ modal thêm/cập nhật hồ sơ quét #mdlAddHoSoQuet-* mở ra

    Return:
        modal_add_hsq: WebElement nếu mở thành công
        False: nếu lỗi
    """

    try:
        wait = WebDriverWait(driver, timeout)

        # --- Tìm hồ sơ quét đầu tiên đang selected ---
        ul_hsq_selected = wait.until(
            lambda d: modal_ho_so_quet.find_element(
                By.CSS_SELECTOR,
                "#vModuleHoSoQuetChon #lstHoSoQuet ul.hosoquet-item.selected"
            )
        )

        print("✅ Đã tìm thấy hồ sơ quét đầu tiên đang selected.")

        # --- Tìm nút Cập nhật trong ul selected ---
        btn_update = wait.until(
            lambda d: ul_hsq_selected.find_element(
                By.CSS_SELECTOR,
                "li.actions a.btnUpdate"
            )
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            btn_update
        )

        # --- Click nút Cập nhật ---
        try:
            btn_update.click()
        except:
            driver.execute_script("arguments[0].click();", btn_update)

        print("✅ Đã click nút Cập nhật hồ sơ quét.")

        # --- Chờ modal Add/Cập nhật Hồ sơ quét mở ---
        modal_add_hsq = wait.until(
            EC.visibility_of_element_located((
                By.CSS_SELECTOR,
                "div[id^='mdlAddHoSoQuet-'].popup.in, "
                "div[id^='mdlAddHoSoQuet-'].modal.in, "
                "div[id^='mdlAddHoSoQuet-'].modal.show, "
                "div[id^='mdlAddHoSoQuet-'].in, "
                "div[id^='mdlAddHoSoQuet-'].show"
            ))
        )
        wait_query_done(driver)
        print("✅ Modal Cập nhật hồ sơ quét đã mở.")
        return modal_add_hsq

    except Exception as e:
        print(f"❌ Lỗi khi click Cập nhật hồ sơ quét: {e}")
        return False
    
# === XÓA HỒ SƠ QUÉT ĐẦU TIÊN ===
def xoa_file_dau_tien_trong_add_hosoquet(driver, modal_add_hsq, timeout=30):
    """
    Trong modal AddHoSoQuet:
    - Tìm dòng tr đầu tiên trong bảng #tbDanhSachFile
    - Nhấn nút #btnRemoveRow trong dòng đó
    - Không có hộp xác nhận
    - Chờ dòng đó biến mất hoặc số dòng giảm

    Return:
        True nếu xóa thành công
        False nếu lỗi
    """

    try:
        wait = WebDriverWait(driver, timeout)

        # --- Chờ bảng danh sách file xuất hiện ---
        table = wait.until(
            lambda d: modal_add_hsq.find_element(
                By.CSS_SELECTOR,
                "#tbDanhSachFile"
            )
        )

        # --- Lấy danh sách dòng hiện có ---
        rows_before = table.find_elements(By.CSS_SELECTOR, "tbody tr")
        if not rows_before:
            print("⚠️ Không có file nào trong bảng #tbDanhSachFile để xóa.")
            return False

        so_dong_truoc = len(rows_before)
        tr_first = rows_before[0]

        print(f"✅ Số dòng file trước khi xóa: {so_dong_truoc}")

        # --- Tìm nút xóa đúng là #btnRemoveRow trong dòng đầu tiên ---
        btn_remove = tr_first.find_element(
            By.CSS_SELECTOR,
            "#btnRemoveRow"
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            btn_remove
        )

        # --- Click nút xóa ---
        try:
            btn_remove.click()
        except:
            driver.execute_script("arguments[0].click();", btn_remove)

        print("✅ Đã nhấn nút #btnRemoveRow của file đầu tiên.")
    except Exception as e:
        print(f"❌ Lỗi khi xóa file đầu tiên bằng #btnRemoveRow: {e}")
        return False
    
# === CẬP NHẬT HỒ SƠ QUÉT ĐẦU TIÊN ===
def them_file_don_dang_ky_trong_add_hosoquet(
    driver,
    modal_add_hsq,
    maxa,
    loaidat,
    timeout=30
):
    """
    Trong modal AddHoSoQuet:
    - Nhấn #btnAddFileHoSoQuet trong #hoSoQuet-*
    - Chờ modal #mdlChiTietHoSoQuet-* mở
    - Chọn loại hồ sơ quét value=2: Đơn đăng ký
    - Nhập mô tả: CHUACOGIAY_{MAXA}_{LOAIDAT}-DDK
    - Nhấn #btnLuuHoSoQuet

    Return:
        True nếu thêm thành công
        False nếu lỗi
    """

    try:
        wait = WebDriverWait(driver, timeout)

        mo_ta = f"CHUACOGIAY_{maxa}_{loaidat}-DDK"

        # --- Tìm vùng hoSoQuet-* trong modal AddHoSoQuet ---
        ho_so_quet_box = wait.until(
            lambda d: modal_add_hsq.find_element(
                By.CSS_SELECTOR,
                "div[id^='hoSoQuet-']"
            )
        )

        # --- Tìm nút Thêm file hồ sơ quét ---
        btn_add_file = wait.until(
            lambda d: ho_so_quet_box.find_element(
                By.CSS_SELECTOR,
                "#btnAddFileHoSoQuet"
            )
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            btn_add_file
        )

        try:
            btn_add_file.click()
        except:
            driver.execute_script("arguments[0].click();", btn_add_file)

        print("✅ Đã nhấn nút thêm tập tin hồ sơ quét.")

        # --- Chờ modal chi tiết hồ sơ quét mở ---
        modal_chi_tiet = wait.until(
            EC.visibility_of_element_located((
                By.CSS_SELECTOR,
                "div[id^='mdlChiTietHoSoQuet-'].popup.in, "
                "div[id^='mdlChiTietHoSoQuet-'].modal.in, "
                "div[id^='mdlChiTietHoSoQuet-'].modal.show, "
                "div[id^='mdlChiTietHoSoQuet-'].in, "
                "div[id^='mdlChiTietHoSoQuet-'].show"
            ))
        )

        print("✅ Modal chi tiết hồ sơ quét đã mở.")
        # --- Chọn loại hồ sơ quét: value = 2 Đơn đăng ký ---
        select_loai = wait.until(
            lambda d: modal_chi_tiet.find_element(
                By.CSS_SELECTOR,
                "select[name='loaiHoSoQuet']"
            )
        )

        Select(select_loai).select_by_value("2")
        print("✅ Đã chọn loại hồ sơ quét: Đơn đăng ký.")

        # --- Nhập mô tả ---
        input_mota = wait.until(
            lambda d: modal_chi_tiet.find_element(
                By.CSS_SELECTOR,
                "input[name='moTa']"
            )
        )

        input_mota.click()
        input_mota.send_keys(Keys.CONTROL, "a")
        input_mota.send_keys(Keys.BACKSPACE)
        input_mota.send_keys(mo_ta)

        print(f"✅ Đã nhập mô tả: {mo_ta}")

        # --- Nhấn lưu ---
        btn_luu = wait.until(
            lambda d: modal_chi_tiet.find_element(
                By.CSS_SELECTOR,
                "#btnLuuHoSoQuet"
            )
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            btn_luu
        )

        try:
            btn_luu.click()
        except:
            driver.execute_script("arguments[0].click();", btn_luu)

        print("✅ Đã nhấn nút Lưu hồ sơ quét.")

        # --- Chờ modal chi tiết đóng lại hoặc biến mất ---
        wait.until(
            EC.invisibility_of_element_located((
                By.CSS_SELECTOR,
                "div[id^='mdlChiTietHoSoQuet-']"
            ))
        )

        print("✅ Đã lưu và đóng modal chi tiết hồ sơ quét.")
        return True

    except Exception as e:
        print(f"❌ Lỗi khi thêm file Đơn đăng ký trong AddHoSoQuet: {e}")
        return False