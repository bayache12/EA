from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains


def viable_collate(driver):
    to_return = []
    collates = ["POST Collate and Sort", "PRE Collate and Sort", "MID Collate and Sort"]
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//label[text()="Ranges"]')))
    for i in driver.find_elements(By.XPATH, value="//label"):
        if i.text in collates:
            to_return.append(i.text)
    return to_return[::-1]


def refresh_search(driver):
    element_before_refresh = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "analyse_academic_year")))
    toggle_search(driver)
    WebDriverWait(driver, 10).until(EC.staleness_of(element_before_refresh))
    toggle_search(driver)


def login(driver, username, password):
    WebDriverWait(driver, 10).until(
        EC.visibility_of_all_elements_located((By.CLASS_NAME, "login_buttons")))[1].click()
    driver.find_element(by=By.ID, value="username").send_keys(username)
    driver.find_element(by=By.ID, value="password").send_keys(password)
    driver.find_element(by=By.ID, value="sign_in").click()
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "teacher_logout")))
    driver.get("https://www.essentialassessment.com.au/quiz/")
    try:
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "settings_wrapper")))
        body_element = driver.find_element(by=By.TAG_NAME, value="body")
        ActionChains(driver).move_to_element(body_element).click().perform()
    except TimeoutException:
        pass


def toggle_search(driver):
    wait_for_click(driver, By.CLASS_NAME, "switch.search_switch")


def has_general_all(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.visibility_of_any_elements_located((By.CLASS_NAME, 'analyse_dropdown_style')))
        elements = driver.find_elements(By.CLASS_NAME, value='analyse_dropdown_style')

        for element in elements:
            if element.text.strip() == "General All":
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[text()="General All"]'))).click()
                return True
        return False

    except TimeoutException:
        return False


def wait_for_click(driver, element, value):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((element, value))).click()


def get_viable_paths(academic_year, driver):
    to_populate = {
    }
    wait_for_click(driver, By.ID, "analyse_academic_year")
    wait_for_click(driver, By.XPATH, f'//button[text()="{academic_year}"]')
    wait_for_click(driver, By.XPATH, '//button[text()="Strand"]')
    for strand in [i.text for i in
                   driver.find_elements(by=By.CLASS_NAME, value="analyse_dropdown_style")
                   if i.text.strip() != ""]:
        wait_for_click(driver, By.ID, "analyse_academic_year")
        wait_for_click(driver, By.XPATH, f'//button[text()="{academic_year}"]')
        wait_for_click(driver, By.XPATH, '//button[text()="Strand"]')
        wait_for_click(driver, By.XPATH, f'//button[text()="{strand}"]')
        wait_for_click(driver, By.XPATH, '//button[text()="Substrand"]')
        if has_general_all(driver):
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#search_wrapper_3_1.search_disabled")))
            wait_for_click(driver, By.XPATH, '//button[text()="Year"]')
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'analyse_dropdown_style.analyse_dropdown_style_2')))
            to_populate[strand] = [
                i.text for i in
                driver.find_elements(By.CLASS_NAME, value='analyse_dropdown_style.analyse_dropdown_style_2')
                if i.text.strip() != ""
            ]
        refresh_search(driver)
    print('done')
    return to_populate


def get_strand_year_files(academic_year, subject, year, driver):
    refresh_search(driver)
    instructions_one = [
        [By.ID, "analyse_academic_year"],
        [By.XPATH, f'//button[text()="{academic_year}"]'],
        [By.XPATH, '//button[text()="Strand"]'],
        [By.XPATH, f'//button[text()="{subject}"]'],
        [By.XPATH, '//button[text()="Substrand"]'],
        [By.XPATH, '//button[text()="General All"]']
    ]
    instructions_two = [
        [By.XPATH, '//button[text()="Year"]'],
        [By.XPATH, f'//button[text()={year}]'],
        [By.XPATH, '//button[text()="Level"]'],
        [By.XPATH, '//button[text()="All Levels"]'],
        [By.ID, "search_button"]
    ]

    for instruction in instructions_one:
        wait_for_click(driver, instruction[0], instruction[1])
    WebDriverWait(driver, 10).until(
        EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#search_wrapper_3_1.search_disabled")))
    for instruction in instructions_two:
        wait_for_click(driver, instruction[0], instruction[1])
    for collate in viable_collate(driver):
        try:
            wait_for_click(driver, By.XPATH, f'//label[text()="{collate}"]')
            wait_for_click(driver, By.ID, "export_results")
            wait_for_click(driver, By.XPATH, '//button[text()="Export A.S.R (XLSX)"]')
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, "pdf_animator_all")))
            print(f"{year} - {subject} - {collate}")
        except:
            continue


def main():
    year = 2022
    script_settings = {"driver path": "chromedriver.exe"}
    state = {"Username": "", "Password": ""}

    driver = webdriver.Chrome(service=Service(script_settings["driver path"]))
    url = 'https://www.essentialassessment.com.au/'
    driver.get(url)
    driver.maximize_window()
    login(driver, state["Username"], state["Password"])
    wait_for_click(driver, By.ID, "export_data_main_text")
    WebDriverWait(driver, 10).until(
        EC.invisibility_of_element_located((By.ID, "pdf_animator_all")))
    toggle_search(driver)
    viable_paths = get_viable_paths(year, driver)
    for key in viable_paths.keys():
        for academic_year in viable_paths[key]:
            get_strand_year_files(year, key, academic_year, driver)


main()
