# %%
import os
import time

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from collections import defaultdict
from selenium.webdriver.support.ui import WebDriverWait
import pickle
from pages.useful_functions import clear_text_box

# %%

# %%
class DiemRenLuyen:
    def __init__(self):
        self.answers = None
        self.path = 'driver/chromedriver.exe'
        self.driver = webdriver.Chrome(executable_path=self.path)
        self.baseURL = 'https://forms.office.com/pages/responsepage.aspx?id=n7jxBugHT0a0COwbRXA_McksYyCtpodMvSxbOvdX-BFUOVlHQlhRVjQyOE1SSktWU08zUjhZRjhCSi4u&fbclid=IwAR20EiRsRh2sE6avyFRmd18H2MfK-LpXJ-LJJ4c9aScn9O-kMSOgt-I5o8g'
        self.answerURL = 'https://forms.office.com/Pages/ResponseDetailPage.aspx?id=n7jxBugHT0a0COwbRXA_McksYyCtpodMvSxbOvdX-BFUOVlHQlhRVjQyOE1SSktWU08zUjhZRjhCSi4u&rid=11044&GetResponseToken=9_haQELQTahXoOPEFyxuNX_mwFHJ0cIOpOtePFUXyzM'

    def save_answer(self):
        answers_elements = self.driver.find_elements_by_xpath(
            "//input[@aria-checked='true']")
        answers = []
        for answer_element in answers_elements:
            answers.append([answer_element.get_attribute(
                'value'), answer_element.get_attribute('aria-posinset')])
        self.answers = answers
        return answers

    def answers_to_pkl(self, filename):
        with open(filename, 'wb') as f:
            pickle.dump(self.answers, f)

    def pkl_to_answers(self, filename):
        with open(filename, 'rb') as f:
            self.answers = pickle.load(f)

    def lam_bai(self, auto_submit = False):
        self.driver.get(self.baseURL)
        test = WebDriverWait(self.driver, 100).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[@class='office-form-title heading-1 ']"))
        )
        if self.answers is not None:
            for answer in self.answers:
                if answer[1] is not None:
                    elements = self.driver.find_elements_by_xpath(
                        f"//input[@value='{answer[0]}' and @aria-posinset='{answer[1]}']")
                else:
                    elements = self.driver.find_elements_by_xpath(
                        f"//input[@value='{answer[0]}']")
                try:
                    if isinstance(elements, list):
                        for element in elements:
                            element.click()
                except:
                    print(elements, "can't be clicked")
        else:
            print("self.answers is None")

      

    def input_answer(self):
        self.driver.get(self.answerURL)
        waiting = True
        while waiting:
            try:
                self.save_answer()
                waiting = False
            except NoSuchElementException:
                waiting = True

    def auto_log_in(self, username, password):
        EMAILFIELD = (By.ID, "i0116")
        PASSWORDFIELD = (By.ID, "i0118")
        NEXTBUTTON = (By.ID, "idSIButton9")
        SUBMIT = (By.ID, "submitButton")
        NO_STAY = (By.ID, "idBtn_Back")
        self.driver.quit()
        self.driver = webdriver.Chrome(executable_path=self.path)
        self.driver.get(self.baseURL)
        self._set_input_by_locator(EMAILFIELD, username)
        self._click_by_locator(NEXTBUTTON)
        self._set_input_by_id("passwordInput", password)
        self._click_by_locator(SUBMIT)
        self._click_by_locator(NO_STAY)

    ################################################################################################
    #                                                                                              #
    #                                   useful function                                            #
    #                                                                                              #
    #################################################################################################

    def _click_button_by_XPATH(self, button_XPATH):
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, button_XPATH))).click()


    def _set_input_by_id(self, input_box_id, input_value):
        input_box = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.ID, input_box_id))
        )
        clear_text_box(input_box)
        input_box.send_keys(input_value)

    def _set_input_by_XPATH(self, input_box_XPATH, input_value):
        input_box = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, input_box_XPATH))
        )
        clear_text_box(input_box)
        input_box.send_keys(input_value)

    def _click_by_link_text(self, text):
        WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.LINK_TEXT, text))
        ).click()

    def set_select(self, select_ID, value):
        box = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.ID, select_ID)))

        option = Select(box)
        option.select_by_value(value)

    def set_select_by_class(self, class_name, value):
        box = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, class_name)))

        option = Select(box)
        option.select_by_value(value)


    def _set_input_by_locator(self, locator, input_value):
        input_box = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located(locator)
        )
        clear_text_box(input_box)
        input_box.send_keys(input_value)

    def _click_by_locator(self, locator):
        WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable(locator)
        ).click()

    def _click_button_by_ID(self, button_ID):
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.ID, button_ID))).click()



drl = DiemRenLuyen()
drl.pkl_to_answers("answer/ketqua.pkl")


# %%
email = ''
pwd = ''
drl.auto_log_in(email, pwd)
drl.lam_bai()
