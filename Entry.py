from utils.config import *
import numpy as np
from datetime import datetime
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options as Chrome_options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import warnings, time, operator, gc, sys, os, pyautogui, pyperclip
from selenium.webdriver.support.ui import Select
import pandas as pd

warnings.filterwarnings('ignore')


class CustomerObjection:
    """Current chromedriver version is 83. In order to keep this driver compatible with running chromebrowser version,
    you can download newer version of chromedriver or disable chromebrowser's auto update function.
    If chromebrowser version is ahead of the installed chromedriver version, chromedriver cannot manipulate chromebrowser.

    Chromedriver download : https://chromedriver.chromium.org/downloads
    How to disable chromedriver auto update(KOR) : https://rgy0409.tistory.com/2301

    Since admin blocked access to chromedriver download page, disabling chromedriver is the easy way to go.

    """

    GC_DRIVER = 'driver/chromedriver.exe'
    CHROME_OPTIONS = Chrome_options()
    CHROME_OPTIONS.add_argument("--start-maximized")

    # CHROME_OPTIONS.add_argument("--disable-extensions")
    # CHROME_OPTIONS.add_argument("--incognito")

    def __init__(self):
        objections = Pipeline()
        self.log = None
        self.df = objections.df
        self.tot_seq = len(objections.storage)
        self.objset = objections.objection_generator()
        self.delimiter = objections.delimiter
        self.filters = objections.filters
        self.customer = None
        self.driver = webdriver.Chrome(self.GC_DRIVER, options=self.CHROME_OPTIONS)
        self.driver.get(URL)
        self.sequence = 0
        self.length = 0
        self.amount = 0

    def click_element_id(self, ID, sec):
        try:
            element = WebDriverWait(self.driver, sec).until(
                EC.element_to_be_clickable((By.ID, ID)))
            element.click()
            return element
        except TimeoutException:
            pass

    def log_in(self):
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_3_form_ImageViewer00', 3)
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_3_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_2_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_1_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_0_form_ImageViewer00', 1)
        id_box = self.driver.find_element_by_id('mainframe_VFrameSet_LoginFrame_form_div_logo_edt_userId_input')
        login_button = self.driver.find_element_by_id('mainframe_VFrameSet_LoginFrame_form_div_logo_btn_login')
        id_box.send_keys(GSCM_ID)
        ActionChains(self.driver).send_keys(Keys.TAB).send_keys(GSCM_PW).perform()
        login_button.click()

    def inner_remove_noti(self):
        # self.click_element_id("mainframe_VFrameSet_TopFrame_CommLgds010P_4_form_ImageViewer00", 3)
        self.click_element_id('mainframe_VFrameSet_TopFrame_CommLgds010P_4_form_ImageViewer00', 3)  # todo : consider commenting out
        self.click_element_id('mainframe_VFrameSet_TopFrame_CommLgds010P_3_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_TopFrame_CommLgds010P_2_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_TopFrame_CommLgds010P_1_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_TopFrame_CommLgds010P_0_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_TopFrame_gfn_alert_form_alert_bottom_bg_btn_confirm', 1)

    def export_page(self):
        self.log_in()
        self.inner_remove_noti()
        dropdown = self.click_element_id('mainframe_VFrameSet_TopFrame_form_cbo_auth_sub_dropbutton', 5)
        dropdown.send_keys(Keys.DOWN)
        dropdown.send_keys(Keys.RETURN)

    def get_workfloor(self):
        self.export_page()
        pyautogui.moveTo(COORDINATES['QAMENU'])
        s_menu = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), '변제관리')]")))
        s_menu.click()
        time.sleep(3)
        self.driver.switch_to.window(self.driver.window_handles[1])

    def input_year_month(self, issuemonth, wait):
        """Redesigned on 2020.03.30"""
        seq_m = self.click_element_id('K_RPYM_WNTI_MM', 3)
        seq_m.send_keys(Keys.DELETE)
        seq_y = self.click_element_id('K_RPYM_WNTI_YY', 3)
        seq_y.send_keys(issuemonth[0:4])
        seq_m.send_keys(issuemonth[4:6])
        time.sleep(wait)

    def setting(self, feed):
        issuemonth = ''.join(feed[5])[0:6]
        self.input_year_month(issuemonth, 1)
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, "//select[@name='K_PRDN_CORP_CD']/option[text()='{}']".format(feed[0])))).click()
        while True:
            try:
                time.sleep(1)
                WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
                        By.XPATH, "//select[@name='K_RPYM_WNTI_NO']/option[@value='{}']".format(
                            feed[5][0] + feed[5][1] + '-' + feed[5][2] + '-' + feed[1][0])))).click()
                break
            except:
                print('TIME OUT')
                continue

    def creation_loop(self, feed):
        """Redesigned on 2020.03.25"""
        index = [feed[0] + self.delimiter + feed[1] + self.delimiter + feed[2] + self.delimiter + str(
            feed[3]) + self.delimiter + feed[4]]
        if 'Created' in self.df.loc[index]['Result'].tolist():
            return
        noc = 1
        self.setting(feed)
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="K_MKOB_TYPE_CD"]/option[@value="{}"]'.format(feed[4])))).click()
        chained = False
        print('\nSEQUENCE : {}/{}'.format(self.sequence, self.tot_seq), '\nFEED : {}'.format(feed),
              '\nLength : {}'.format(len(feed[6])), '\nAmount : {}'.format(feed[-1]), '\nIssueMonth : {}'.format(
                feed[5][0] + feed[5][1] + '-' + feed[5][2] + '-' + feed[1][0]))
        for n in range(len(feed[6])):
            if not chained:
                front = self.click_element_id('K_FIRM_INFM_SN_F', 3)
                front.send_keys(Keys.DELETE)
                front.send_keys(str(feed[6][n]))
            if n != len(feed[6]) - 1 and feed[6][n] - feed[6][n + 1] == -1 and noc < 500:
                noc += 1
                chained = True
                print(feed[6][n], end='=>')
                continue
            else:
                noc = 1
                chained = False
                tail = self.click_element_id('K_FIRM_INFM_SN_T', 3)
                tail.send_keys(Keys.DELETE)
                tail.send_keys(str(feed[6][n]))
                print(feed[6][n])
                self.click_element_id('searchBtn', 3)
                time.sleep(0.7)
                self.click_element_id('groupRejectBtn', 3)
                self.update_hist(feed, 'Created')
                continue
        self.logging(feed, 'Created')

    def get_searched_data_information(self):
        while True:
            try: WebDriverWait(self.driver, 1).until(EC.text_to_be_present_in_element((
                        By.XPATH, "// *[ @ id = 'statusArea']"), "You"))
            except:
                try: WebDriverWait(self.driver, 1).until(EC.text_to_be_present_in_element((
                        By.XPATH, "// *[ @ id = 'statusArea']"), "Theres"))
                except:
                    pass
                else:
                    text = None
                    return text
            else:
                text = WebDriverWait(self.driver, 1).until(EC.presence_of_element_located((
                            By.XPATH, "// *[ @ id = 'statusArea']"))).text
                return text.split(' ')[2]

    def register(self, feed):
        WebDriverWait(self.driver, 7).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menu"]/div[3]/ul/li[4]/a'))).click()
        self.setting(feed)
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, "//*[@id='K_GUBN']/option[@value='N']"))).click()
        time.sleep(1)
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, "//*[@id='K_MKOB_TYPE_CD']/option[@value='{}']".format(feed[4])))).click()
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="searchBtn"]'))).click()
        pyperclip.copy(feed[3])

        text = self.get_searched_data_information()

        if text != str(len(feed[6])):
            if self.sequence == 1:
                pyautogui.hotkey('alt', 'tab', 'left', interval=0.1)
            ans = pyautogui.confirm(text=f'{self.sequence}/{self.tot_seq} 건수가 불일치합니다. 교류클레임 여부 확인하세요. '
                                         f'\n건수: {len(feed[6])}, 금액: {feed[-1]}. \n사유: {feed[3]}.', title='등록확인',
                                    buttons=['OK', 'NO'])
        else :
            ans = 'OK'
            time.sleep(1)

        if ans.upper() == 'OK':
            self.driver.switch_to.window(self.driver.window_handles[1])
            self.length += len(feed[6])
            self.amount += feed[-1]
            WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
                By.XPATH, "//*[@id='K_MKOB_TYPE_CD']/option[@value='{}']".format(feed[4])))).click()
            WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
                By.XPATH, '//*[@id="insertBtn"]'))).click()
            time.sleep(0.7)
            pyautogui.press('enter')
            time.sleep(0.5)
            letter_box_1 = self.click_element_id('D_MKOB_TITL_NM', 5)
            letter_box_1.send_keys(str(feed[3]))
            letter_box_2 = self.click_element_id('D_MKOB_DTL_SBC', 5)
            letter_box_2.send_keys(str(feed[3]))
            self.click_element_id('dlg_AppealSaveBtn', 5)
            time.sleep(0.7)
            pyautogui.press('enter')
            time.sleep(0.5)

            objection_no = Select(WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
                By.XPATH, '// *[ @ id = "K_MKOB_OFDC_NO"]')))).options[-1].get_attribute('innerText')

            self.logging(feed, 'Registered')
            self.update_hist(feed, 'Registered', objection_no)
        else:
            self.driver.switch_to.window(self.driver.window_handles[1])
        WebDriverWait(self.driver, 7).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menu"]/div[3]/ul/li[3]/a'))).click()

    def request(self):
        """Redesigned on 2020.03.30"""
        min_year = sorted([date[0:6] for date in self.df['ISSUE NO'].tolist()])[0]
        WebDriverWait(self.driver, 7).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menu"]/div[3]/ul/li[5]/a'))).click()
        print('\nCustomer : {}, Starting month : {}, Length : {}, Amount : {}'.format(self.customer, min(min_year),
                                                                                      self.length, self.amount))
        self.input_year_month(str(min_year), 1)
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, "//select[@name='K_PRDN_CORP_CD']/option[text()='{}']".format(self.customer)))).click()
        self.click_element_id('searchBtn', 3)

    def mainloop(self):
        with open('Cookies_objection/log.txt', 'a+') as txt:
            txt.write(f'\n{datetime.now().strftime("%Y/%m/%d %H:%M:%S")} Initiated\n')
        start = time.time()
        self.get_workfloor()
        WebDriverWait(self.driver, 7).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menu"]/div[3]/ul/li[3]/a'))).click()

        for feed in self.objset:
            self.customer = feed[0].upper()
            self.sequence += 1
            self.creation_loop(feed)
            self.register(feed)

        elapsed = 'Elapsed {0:02d}:{1:02d}'.format(*divmod(int(time.time() - start), 60))
        print(f'Elapsed : {elapsed}')
        with open('Cookies_objection/log.txt', 'a+') as txt:
            txt.write(f'{datetime.now().strftime("%Y/%m/%d %H:%M:%S")}, {elapsed} Registration Finished\n')

        self.request()
        pyautogui.alert(
            text=f'Customer : {self.customer}, Length : {self.length}, Amount : {self.amount}, '
                 f'\n소요시간 : {elapsed} \n금액, 건수 검증하고 이의제기 의뢰하세요. \n의뢰 한 후 브라우저를 닫으세요.',
            title='프로세스종료알림', button='OK')

        self.save_df()
        input('Press <ENTER> to terminate...')
        self.close()

    def logging(self, feed, stage):
        list_no = ', '.join([str(i) for i in feed[6]])
        self.log = datetime.now().strftime("%Y/%m/%d %H:%M:%S") + ' ' + feed[0] + ' ' + feed[5][3] + '-' + feed[1][0] + ', ' + feed[
            3] + f', ({list_no}), {stage}\n'
        with open('Cookies_objection/log.txt', 'a+') as txt:
            txt.write(self.log)

    def update_hist(self, feed, stage, objection_no=None):
        index = feed[0] + self.delimiter + feed[1] + self.delimiter + feed[2] + self.delimiter + feed[3] + \
                self.delimiter + feed[4]
        if objection_no:
            self.df.at[index, 'Result'] = objection_no
        else :
            self.df.at[index, 'Result'] = stage

    def save_df(self):
        self.df['Result'] = np.where(self.filters, self.df['Result'], '')
        self.df.to_excel(r'Cookies_objection\objection.xls', index=False)
        os.startfile(r'Cookies_objection\objection.xls')

    @classmethod
    def run(cls):
        """ 2020.9.28 This is the main execution method of this class """

        obj = cls()
        try:
            obj.mainloop()
        except Exception as e:
            # raise
            print(f'Error occurred! {e}')
            obj.save_df()
            obj.close()

    def close(self):
        try:
            self.driver.delete_all_cookies()
            self.driver.quit()
        except Exception as e:
            print(f'Closing error occurred! {e}')
        sys.exit()


class Pipeline:
    def __init__(self):
        self.filename = r'Cookies_objection\objection.xls'
        with open(self.filename, 'rb') as file:
            self.df = pd.read_excel(file)
            self.df.fillna('', inplace=True)

        self.customer = self.df.iloc[0]['고객사']
        self.delimiter = ',,'
        self.df['KEY'] = self.df['고객사'] + self.delimiter + self.df['E/D'] + self.delimiter + self.df['ISSUE NO'] + \
                         self.delimiter + self.df['OBJECTION_'].apply(str) + self.delimiter + self.df['유형']
        self.df.set_index('KEY', inplace=True)
        if 'Result' not in self.df.columns:
            self.df['Result'] = ['' for _ in range(len(self.df))]

        self.filters = (
                ~(self.df['Customer Reivew_'].str.contains('reject')) &
                ~(self.df['Customer Reivew_'].str.contains('wait')) &
                ~(self.df['Customer Reivew_'].str.contains('pending')) &
                ~(self.df['Customer Reivew_'].str.contains('denied')) &
                ~(self.df['Customer Reivew_'].str.contains('wrong')) &
                ~(self.df['Customer Reivew_'].str.contains('deny')) &
                ~(self.df['Customer Reivew_'].str.contains('done')) &
                ~(self.df['Customer Reivew_'].str.contains('cancel')) &
                ~(self.df['Customer Reivew_'].str.contains('기각'))
        )

        self.df_ = self.df[self.filters & ~(self.df['Result'].str.contains('SEF9'))]
        self.keys = set(self.df_.index)
        self.storage = list()
        self.data = self.isolate()
        try:
            self.print_example()
        except IndexError as e:
            print(f'No data is available. {e}')
            sys.exit()

    def isolate(self):
        for key in self.keys:
            val_list = key.split(self.delimiter)
            val_list.append([val_list[2][0:4], val_list[2][4:6], val_list[2][6:],
                             val_list[2][0:4] + '-' + val_list[2][4:6] + '-' + val_list[2][6:]])
            partial_df = self.df_.loc[key]
            try:
                val_list.append(sorted(partial_df['LIST'].to_numpy().tolist()))
            except AttributeError:
                val_list.append([partial_df['LIST']])
            val_list.append(round(partial_df['전체'].sum(), 2))
            self.storage.append(val_list)
        self.storage = sorted(self.storage, key=operator.itemgetter(0, 1, 2, 3, 6))
        return self.storage

    def print_example(self):
        example = list(enumerate(self.storage))[0][1]
        print('\nNumber of Objections : {}, Total Amount : {:,.2f}, Total Sequence : {}'.format(
            len(self.df_), round(self.df_['전체'].sum(), 2), len(self.storage)))
        print('\nFeed Format : {}'.format(example))
        for i, item in enumerate(example):
            print('Index {} : {}'.format(i, item))

    def objection_generator(self):
        for datum in self.data:
            yield datum


if __name__ == '__main__':
    CustomerObjection.run()
