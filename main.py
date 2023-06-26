from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains


class EAScraper:

    def __init__(self, url, username,
                 password, sleep_length, academic_year, driver_path):
        self.missed_files = []
        self.viable_paths = {}
        self.url = url
        self.username = username
        self.password = password
        self.sleep_length = sleep_length
        self.academic_year = academic_year
        self.driver = webdriver.Chrome(service=Service(driver_path))

    def click_on_element(self, element, value):
        WebDriverWait(self.driver, self.sleep_length).until(
            EC.element_to_be_clickable((element, value))).click()

    def toggle_search_button(self):
        self.click_on_element(By.CLASS_NAME, "switch.search_switch")

    def refresh_search(self):
        element_before_refresh = WebDriverWait(self.driver, self.sleep_length).until(
            EC.visibility_of_element_located((By.ID, "analyse_academic_year")))
        self.toggle_search_button()
        WebDriverWait(self.driver, self.sleep_length).until(
            EC.staleness_of(element_before_refresh))
        self.toggle_search_button()

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

        except TimeoutException:
            print("No pop up detected")

        finally:
            self.click_on_element(By.ID, "export_data_main_text")
            WebDriverWait(self.driver, self.sleep_length).until(
                EC.invisibility_of_element_located((By.ID, "pdf_animator_all")))
            self.toggle_search_button()

    def has_general_all(self):
        try:
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
        WebDriverWait(self.driver, self.sleep_length).until(
            EC.presence_of_element_located((By.XPATH, '//label[text()="Ranges"]')))
        for element in self.driver.find_elements(By.XPATH, value="//label"):
            if element.text in ("POST Collate and Sort",
                                "PRE Collate and Sort",
                                "MID Collate and Sort"):
                viable_collates.append(element.text)
        return viable_collates

    def populate_viable_paths(self):
        self.click_on_element(By.ID, "analyse_academic_year")
        self.click_on_element(By.XPATH, f'//button[text()="{self.academic_year}"]')
        self.click_on_element(By.XPATH, '//button[text()="Strand"]')

        for strand in [i.text for i in
                       self.driver.find_elements(by=By.CLASS_NAME, value="analyse_dropdown_style")
                       if i.text.strip()]:
            self.click_on_element(By.ID, "analyse_academic_year")
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
                    self.driver.find_elements(By.CLASS_NAME, value='analyse_dropdown_style.analyse_dropdown_style_2')
                    if i.text and i.text.isnumeric()
                ]
            self.refresh_search()

    def get_file(self, subject, classroom_year):

        try:

            buttons_to_click = (
                (By.ID, "analyse_academic_year"),
                (By.XPATH, f'//button[text()="{self.academic_year}"]'),
                (By.XPATH, '//button[text()="Strand"]'),
                (By.XPATH, f'//button[text()="{subject}"]'),
                (By.XPATH, '//button[text()="Substrand"]'),
                (By.XPATH, '//button[text()="General All"]')
            )

            for button in buttons_to_click:
                self.click_on_element(button[0], button[1])

            WebDriverWait(self.driver, self.sleep_length).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#search_wrapper_3_1.search_disabled")))

            self.click_on_element(By.XPATH, '//button[text()="Year"]')
            self.click_on_element(By.XPATH, f'//button[text()={classroom_year}]')
            self.click_on_element(By.XPATH, '//button[text()="Level"]')
            self.click_on_element(By.XPATH, '//button[text()="All Levels"]')
            self.click_on_element(By.ID, "search_button")

            for collate in self.has_viable_collate():
                try:
                    self.click_on_element(By.XPATH, f'//label[text()="{collate}"]')
                    self.click_on_element(By.ID, "export_results")
                    self.click_on_element(By.XPATH, '//button[text()="Export A.S.R (XLSX)"]')
                    WebDriverWait(self.driver, self.sleep_length).until(
                        EC.invisibility_of_element_located((By.ID, "pdf_animator_all")))
                except TimeoutException:
                    self.missed_files.append(f'{subject}-{classroom_year}-{collate} '
                                             f'not found, likely disappearing bug')
        except TimeoutException:
            self.missed_files.append(f'Files likely dont exist: '
                                     f'{subject}-{classroom_year}')
        self.refresh_search()

    def begin_process(self):
        self.login()
        self.get_export()
        self.populate_viable_paths()
        for subject in self.viable_paths:
            for classroom_year in self.viable_paths[subject]:
                self.get_file(subject, classroom_year)

        print(self.missed_files)


st_anthony = EAScraper(
    url="https://www.essentialassessment.com.au/",
    username="",
    password="",
    sleep_length=20,
    academic_year=2022,
    driver_path="chromedriver.exe"
)
st_anthony.begin_process()
