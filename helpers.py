from selenium.webdriver.support.ui import WebDriverWait
import time

def wait_query_done(driver, timeout=30, ajax_wait=5):
    end_time = time.time() + timeout
    try:
        WebDriverWait(driver, 5).until(
            lambda d: d.execute_script("return window.jQuery !== undefined;")
        )
    except Exception:
        return

    phase1_end = time.time() + ajax_wait
    saw_ajax = False
    while time.time() < phase1_end:
        try:
            active = driver.execute_script("return jQuery.active;")
            if active > 0:
                saw_ajax = True
                break
        except Exception:
            break

    if not saw_ajax:
        return

    while time.time() < end_time:
        try:
            active = driver.execute_script("return jQuery.active;")
            if active == 0:
                return
        except Exception:
            return

def wait_query_xoadon(driver, timeout=30, ajax_wait=5, max_after_first_ajax=10):
    try:
        WebDriverWait(driver, 3).until(
            lambda d: d.execute_script("return window.jQuery !== undefined;")
        )
    except Exception:
        return

    phase1_end = time.time() + ajax_wait
    saw_ajax = False
    while time.time() < phase1_end:
        try:
            active = driver.execute_script("return jQuery.active;")
            if active > 0:
                saw_ajax = True
                break
        except Exception:
            return

    if not saw_ajax:
        return

    phase2_end = time.time() + max_after_first_ajax
    THRESH = 1
    stable_required = 5
    stable_count = 0

    while time.time() < phase2_end:
        try:
            active = driver.execute_script("return jQuery.active;")
        except Exception:
            return

        if active <= THRESH:
            stable_count += 1
            if stable_count >= stable_required:
                return
        else:
            stable_count = 0
    return