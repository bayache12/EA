import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, \
    ElementClickInterceptedException, ElementNotInteractableException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.remote.webelement import WebElement


class EAScraper:

    def __init__(self, username, school_name,
                 password, sleep_length, academic_year,
                 driver_path, download_path):

        self.full_path = f'{download_path}/{school_name}/{academic_year}'
        self.temp_folder = rf'{download_path}\{school_name}\{academic_year}\temp'
        self.missed_files = []
        self.viable_paths = {}
        self.url = "https://www.essentialassessment.com.au/"
        self.username = username
        self.password = password
        self.sleep_length = sleep_length
        self.academic_year = academic_year
        self.download_path = download_path
        self.school_name = school_name
        chromeOptions = webdriver.ChromeOptions()
        prefs = {"download.default_directory": rf'{download_path}\{school_name}\{academic_year}\temp',
                 "safebrowsing.disable_download_protection": True}
        print(prefs)
        chromeOptions.add_experimental_option("prefs", prefs)
        self.driver = webdriver.Chrome(service=Service(driver_path),
                                       options=chromeOptions)

    def click_on_element(self, element, value):
        clicked = False
        while not clicked:
            try:
                WebDriverWait(self.driver, self.sleep_length).until(
                    EC.element_to_be_clickable((element, value))).click()
                clicked = True
            except (ElementClickInterceptedException, ElementNotInteractableException) as e:

                print(e)
                print(f"---- Something was blocking button: {value} trying again in 1 second.")
                time.sleep(1)

    def create_directory(self):
        try:
            if self.school_name not in os.listdir(self.download_path):
                print("---- School Directory not existent in this path, creating")
                os.mkdir(f'{self.download_path}/{self.school_name}')
                os.mkdir(self.full_path)
                os.mkdir(f'{self.full_path}/temp')
            elif self.academic_year not in os.listdir(f'{self.download_path}/{self.school_name}'):
                print("---- Academic year directory not existent for this school, creating")
                os.mkdir(self.full_path)
                os.mkdir(f'{self.full_path}/temp')
            else:
                print("---- School directory already exists")
        except Exception as e:
            print(f"---- Error creating directory for file storage, Please fix and rerun...")
            time.sleep(10)

    def rename_file(self, new_name):
        try:
            for file in os.listdir(self.temp_folder):
                if file.endswith('.xlsx'):
                    os.rename(f'{self.temp_folder}/{file}',
                              f'{self.full_path}/{new_name}.xlsx')
        except FileExistsError:
            print(f"X File {new_name} failed to be renamed and sorted")

    def wait_for_download(self):
        print(self.temp_folder)
        while not [file for file in os.listdir(self.temp_folder) if file.endswith('.xlsx')]:
            print("------ Waiting for download")
            time.sleep(2)
        if len([file for file in os.listdir(self.temp_folder) if file.endswith('.xlsx')]) != 1:
            input("Multiple Excel Files Detected In school temp file Directory, Please Fix and then press any key: ")

    def toggle_search_button(self):
        self.click_on_element(By.CLASS_NAME, "switch.search_switch")

    def refresh_search(self):
        self.toggle_search_button()
        WebDriverWait(self.driver, self.sleep_length).until(
            EC.visibility_of_element_located((By.ID, "export_data_main_text")))
        self.toggle_search_button()

    def close_ea_trial_button(self):  # <- new pop up section after closing side pop up
        self.click_on_element(By.ID, "ea_trial_video_close")

    def click_button_bottom(self, element, value):

        WebDriverWait(self.driver, self.sleep_length).until(
            EC.element_to_be_clickable((element, value)))

        button = self.driver.find_element(by=element, value=value)

        button_width = button.size['width']
        button_height = button.size['height']

        offset_x = button_width // 2
        offset_y = button_height // 3

        actions = ActionChains(self.driver)
        actions.move_to_element(button).move_by_offset(offset_x, offset_y).click().perform()

    def login(self):

        self.driver.get(self.url)
        self.driver.maximize_window()
        WebDriverWait(self.driver, self.sleep_length).until(
            EC.visibility_of_all_elements_located((By.CLASS_NAME, "login_buttons")))[1].click()
        self.driver.find_element(by=By.ID, value="username").send_keys(self.username)
        self.driver.find_element(by=By.ID, value="password").send_keys(self.password)
        self.driver.find_element(by=By.ID, value="sign_in").click()

    def get_export(self):
        WebDriverWait(self.driver, self.sleep_length).until(
            EC.visibility_of_element_located((By.ID, "teacher_logout")))
        self.driver.get("https://www.essentialassessment.com.au/quiz/")

        try:

            WebDriverWait(self.driver, self.sleep_length).until(
                EC.visibility_of_element_located((By.ID, "settings_wrapper")))
            body_element = self.driver.find_element(by=By.TAG_NAME, value="body")
            ActionChains(self.driver).move_to_element(body_element).click().perform()
            self.close_ea_trial_button()

        except TimeoutException:
            print("Either or pop up not-detected - if first undetected might crash")

        finally:
            WebDriverWait(self.driver, self.sleep_length).until(
                EC.invisibility_of_element_located((By.ID, "ea_intro_nt")))
            self.click_on_element(By.ID, "export_data_main_text")
            print("Clicked")
            self.click_on_element(By.XPATH, f'//button[text()="Export {self.academic_year} Data"]')
            WebDriverWait(self.driver, self.sleep_length).until(
                EC.invisibility_of_element_located((By.ID, "pdf_animator_all")))
            self.wait_for_download()
            self.rename_file("Main Export")
            print(f'---- Main Export Successfully Downloaded')
            self.toggle_search_button()

    def has_general_all(self):
        time.sleep(1)
        try:
            if not self.driver.find_elements(by=By.XPATH,value='//button[text()="No Completed Sub-Strands"]'):
                WebDriverWait(self.driver, self.sleep_length).until(
                    EC.visibility_of_any_elements_located((By.CLASS_NAME, 'analyse_dropdown_style')))
                elements = self.driver.find_elements(By.CLASS_NAME, value='analyse_dropdown_style')

                for element in elements:
                    if element.text.strip() == "General All":
                        self.click_on_element(By.XPATH, '//button[text()="General All"]')
                        return True
            return False

        except TimeoutException:
            return False

    def has_viable_collate(self):
        viable_collates = []
        try:
            WebDriverWait(self.driver, self.sleep_length).until(
                EC.presence_of_element_located((By.XPATH, '//label[text()="Ranges"]')))
            for element in self.driver.find_elements(By.XPATH, value="//label"):
                if element.text in ("POST Collate and Sort",
                                    "PRE Collate and Sort",
                                    "MID Collate and Sort"):
                    viable_collates.append(element.text)
        except TimeoutException:
            print("X Ranges Button Not Found: Increase sleep times if strand was skipped X")
        finally:
            return viable_collates

    def populate_viable_paths(self):
        self.click_button_bottom(By.ID, "analyse_academic_year")
        self.click_on_element(By.XPATH, f'//button[text()="{self.academic_year}"]')
        self.click_on_element(By.XPATH, '//button[text()="Strand"]')

        for strand in [i.text for i in
                       self.driver.find_elements(by=By.CLASS_NAME, value="analyse_dropdown_style")
                       if i.text.strip()]:
            try:
                self.click_button_bottom(By.ID, "analyse_academic_year")
                self.click_on_element(By.XPATH, f'//button[text()="{self.academic_year}"]')
                self.click_on_element(By.XPATH, '//button[text()="Strand"]')
                self.click_on_element(By.XPATH, f'//button[text()="{strand}"]')
                self.click_on_element(By.XPATH, '//button[text()="Substrand"]')
                if self.has_general_all():
                    WebDriverWait(self.driver, self.sleep_length).until(
                        EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#search_wrapper_3_1.search_disabled")))
                    self.click_on_element(By.XPATH, '//button[text()="Year"]')
                    WebDriverWait(self.driver, self.sleep_length).until(
                        EC.visibility_of_element_located(
                            (By.CLASS_NAME, 'analyse_dropdown_style.analyse_dropdown_style_2')))
                    self.viable_paths[strand] = [
                        i.text for i in
                        self.driver.find_elements(By.CLASS_NAME,
                                                  value='analyse_dropdown_style.analyse_dropdown_style_2')
                        if i.text and i.text.isnumeric()
                    ]
            except TimeoutException:
                print(f"X Failed to include {strand} in viable paths, might be missing substrand")

            finally:
                self.refresh_search()

    def get_file(self, subject, classroom_year):

        try:

            buttons_to_click = (
                (By.XPATH, f'//button[text()="{self.academic_year}"]'),
                (By.XPATH, '//button[text()="Strand"]'),
                (By.XPATH, f'//button[text()="{subject}"]'),
                (By.XPATH, '//button[text()="Substrand"]'),
                (By.XPATH, '//button[text()="General All"]')
            )

            self.click_button_bottom(By.ID, "analyse_academic_year")

            for button in buttons_to_click:
                self.click_on_element(button[0], button[1])

            WebDriverWait(self.driver, self.sleep_length).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#search_wrapper_3_1.search_disabled")))

            self.click_on_element(By.XPATH, '//button[text()="Year"]')
            self.click_on_element(By.XPATH, f'//button[text()={classroom_year}]')

            WebDriverWait(self.driver, self.sleep_length).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#search_wrapper_2_1.search_disabled")))
            time.sleep(1)
            self.click_on_element(By.XPATH, '//button[text()="Level"]')
            time.sleep(1)
            WebDriverWait(self.driver, self.sleep_length).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#analyse_search_asr_wrapper_2.search_disabled")))

            WebDriverWait(self.driver, self.sleep_length).until(
                EC.visibility_of_element_located((By.XPATH, '//button[text()="All Levels"]')))
            self.click_on_element(By.XPATH, '//button[text()="All Levels"]')

            WebDriverWait(self.driver, self.sleep_length).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#analyse_button_wrapper.search_disabled")))

            self.click_on_element(By.ID, "search_button")

            for collate in self.has_viable_collate():
                try:
                    self.click_on_element(By.XPATH, '//label[text()="Ranges"]')
                    self.click_on_element(By.XPATH, f'//label[text()="{collate}"]')
                    self.click_on_element(By.ID, "export_results")
                    self.click_on_element(By.XPATH, '//button[text()="Export A.S.R (XLSX)"]')
                    WebDriverWait(self.driver, self.sleep_length).until(
                        EC.invisibility_of_element_located((By.ID, "pdf_animator_all")))
                    self.wait_for_download()
                    self.rename_file(f'Year {classroom_year} {subject} {collate}')
                    print(f'------ Successfully downloaded Year-{classroom_year} {subject} {collate}')
                except StaleElementReferenceException:
                    print(f'X Missed {collate} - disappearing Element - Try increasing wait time')

        except TimeoutException:
            self.missed_files.append(f'{subject}-{classroom_year}')
        finally:
            self.refresh_search()

    def begin_process(self):
        self.login()
        self.create_directory()
        self.get_export()
        self.populate_viable_paths()
        for subject in self.viable_paths:
            for classroom_year in self.viable_paths[subject]:
                print(f'---- Attempting to begin downloads on Year-{classroom_year} {subject}')
                self.get_file(subject, classroom_year)

        if self.missed_files:
            for file in self.missed_files:
                print(f'X {file} not found/downloaded, if file exists please report issue.')
        else:
            print("All Files Successfully Downloaded")


EAScraper(
    username="",
    password="",
    school_name="",
    sleep_length=40,
    academic_year=2022,
    driver_path="chromedriver.exe",
    download_path=r"C:\Users\test\OneDrive\Desktop\test",
).begin_process()
