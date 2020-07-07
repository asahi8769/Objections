from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options as Chrome_options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import warnings, time, operator, gc, sys, os, pyautogui
import pandas as pd
from datetime import datetime

URL = "https://partner.hyundai.com/gscm/"
GSCM_ID = os.environ.get ('GSCM_ID').upper()
GSCM_PW = os.environ.get ('GSCM_PW')

GC_DRIVER = r'driver/chromedriver.exe'
COORDINATES = {'QAMENU' : (309, 178), 'YYMMCOORD' : (670, 204), 'MM_COORD' : (716, 204), 'ROW_COORD' : (308, 380),
               'ISSUENO' : (680, 238)}

CHROME_OPTIONS = Chrome_options()
CHROME_OPTIONS.add_argument("--start-maximized")
CHROME_OPTIONS.add_argument("--disable-extensions")
CHROME_OPTIONS.add_argument("--incognito")
warnings.filterwarnings ('ignore')


class CustomerObjection():
    def __init__(self, stop=False):
        """FOR IE, USE  webdriver.Ie(GI_DRIVER) """
        super().__init__()
        self.objset = Pipeline.run()
        self.customer = self.objset[0][0].upper()
        self.driver = webdriver.Chrome(GC_DRIVER, options=CHROME_OPTIONS)
        self.driver.get(URL)
        self.sequence = 0
        self.stop = stop
        self.length = 0
        self.amount = 0
        self.tot_seq = len(self.objset)
        self.log = None

    def click_element_id(self, ID, sec):
        try:
            element = WebDriverWait (self.driver, sec).until(
                EC.element_to_be_clickable ((By.ID, ID)))
            element.click()
            return element
        except TimeoutException:
            pass

    def log_in(self):
        self.click_element_id ('mainframe_VFrameSet_LoginFrame_CommLgds010_0_form_ImageViewer00', 7)
        id_box = self.driver.find_element_by_id('mainframe_VFrameSet_LoginFrame_form_div_logo_edt_userId_input')
        login_button = self.driver.find_element_by_id('mainframe_VFrameSet_LoginFrame_form_div_logo_btn_login')
        id_box.send_keys(GSCM_ID)
        ActionChains (self.driver).send_keys (Keys.TAB).send_keys (GSCM_PW).perform ()
        login_button.click ()

    def inner_remove_noti(self):
        self.click_element_id ('mainframe_VFrameSet_TopFrame_CommLgds010P_4_form_ImageViewer00', 3)  # todo : consider commenting out
        self.click_element_id ('mainframe_VFrameSet_TopFrame_CommLgds010P_3_form_ImageViewer00', 1)
        self.click_element_id ('mainframe_VFrameSet_TopFrame_CommLgds010P_2_form_ImageViewer00', 1)
        self.click_element_id ('mainframe_VFrameSet_TopFrame_CommLgds010P_1_form_ImageViewer00', 1)
        self.click_element_id ('mainframe_VFrameSet_TopFrame_CommLgds010P_0_form_ImageViewer00', 1)
        self.click_element_id ('mainframe_VFrameSet_TopFrame_gfn_alert_form_alert_bottom_bg_btn_confirm', 1)

    def export_page(self):
        self.log_in()
        self.inner_remove_noti()
        dropdown = self.click_element_id ('mainframe_VFrameSet_TopFrame_form_cbo_auth_sub_dropbutton', 5)
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

    @staticmethod
    def feeder(objset):
        for feed in objset:
            yield feed

    def input_year_month(self, issuemonth, wait):
        """Redesigned on 2020.03.30"""
        seq_m = self.click_element_id('K_RPYM_WNTI_MM', 3)
        seq_m.send_keys(Keys.DELETE)
        seq_y = self.click_element_id('K_RPYM_WNTI_YY',3)
        seq_y.send_keys(issuemonth[0:4])
        seq_m.send_keys(issuemonth[4:6])
        time.sleep(wait)

    def stopper(self, num):
        self.sequence += 1
        if self.stop:
            if self.sequence == num+1:
                self.close()
                print('TERMINATED BY SCRIPT')
                sys.exit()
        else:
            pass

    def setting(self, feed):
        issuemonth = ''.join(feed[5])[0:6]
        self.input_year_month(issuemonth, 1)
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, "//select[@name='K_PRDN_CORP_CD']/option[text()='{}']".format(feed[0])))).click()
        while True:
            try:
                time.sleep(1)
                WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((
                        By.XPATH, "//select[@name='K_RPYM_WNTI_NO']/option[@value='{}']".format(
                            feed[5][0] + feed[5][1] + '-' + feed[5][2] + '-' + feed[1][0])))).click()
                break
            except :
                print('TIME OUT')
                continue

    def creation_loop(self, feed):
        """Redesigned on 2020.03.25"""
        noc = 1
        self.setting(feed)
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="K_MKOB_TYPE_CD"]/option[@value="{}"]'.format(feed[4])))).click()
        chained = False
        print('\nSEQUENCE : {}/{}'.format(self.sequence, self.tot_seq), '\nFEED : {}'.format(feed),
              '\nLength : {}'.format(len(feed[6])), '\nAmount : {}'.format(feed[-1]), '\nIssueMonth : {}'.format(
                feed[5][0] + feed[5][1] + '-' + feed[5][2] + '-' + feed[1][0]))
        for n in range(len(feed[6])):
            if not chained :
                front = self.click_element_id ('K_FIRM_INFM_SN_F', 3)
                front.send_keys(Keys.DELETE)
                front.send_keys(str(feed[6][n]))
            if n != len(feed[6]) - 1 and feed[6][n] - feed[6][n + 1] == -1 and noc < 500:
                noc += 1
                chained = True
                print(feed[6][n], end='=>')
                continue
            else :
                noc = 1
                chained = False
                tail = self.click_element_id('K_FIRM_INFM_SN_T', 3)
                tail.send_keys(Keys.DELETE)
                tail.send_keys(str(feed[6][n]))
                print(feed[6][n])
                self.click_element_id('searchBtn', 3)
                time.sleep(0.7)
                self.click_element_id('groupRejectBtn', 3)
                continue
        self.logging(feed, 'Created')

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
        # pyautogui.hotkey('alt', 'tab', interval=0.1)
        # ans = input('Verify your objections. Register?[Y/N, Default: Y] : ')
        ans = pyautogui.confirm(text=f'{self.sequence}/{self.tot_seq} 이의제기 등록합니다. 교류클레임 여부 확인하세요. \n건수: {len(feed[6])}, 금액: {feed[-1]}.', title='등록확인', buttons=['OK', 'NO'])
        if ans.lower() != 'NO':
            pyautogui.hotkey('alt', 'tab', interval=0.1)
            self.length += len(feed[6])
            self.amount += feed[-1]
            WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
                By.XPATH, "//*[@id='K_MKOB_TYPE_CD']/option[@value='{}']".format(feed[4])))).click()
            WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
                By.XPATH, '//*[@id="insertBtn"]'))).click()
            time.sleep(0.7)
            pyautogui.press('enter')
            time.sleep(0.5)
            letter_box_1 = self.click_element_id ('D_MKOB_TITL_NM', 5)
            letter_box_1.send_keys(str(feed[3]))
            letter_box_2 = self.click_element_id('D_MKOB_DTL_SBC', 5)
            letter_box_2.send_keys(str(feed[3]))
            self.click_element_id('dlg_AppealSaveBtn', 5)
            time.sleep(0.7)
            pyautogui.press('enter')
            time.sleep(0.5)
            self.logging (feed, 'Registered')
        else:
            pyautogui.hotkey('alt', 'tab', interval=0.1)
        WebDriverWait(self.driver, 7).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menu"]/div[3]/ul/li[3]/a'))).click()

    def request(self, min_year):
        """Redesigned on 2020.03.30"""
        WebDriverWait(self.driver, 7).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menu"]/div[3]/ul/li[5]/a'))).click()
        print ('\nCustomer : {}, Starting month : {}, Length : {}, Amount : {}'.format(self.customer, min(min_year), self.length, self.amount))
        self.input_year_month(str(min(min_year)), 1)
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((
            By.XPATH, "//select[@name='K_PRDN_CORP_CD']/option[text()='{}']".format(self.customer)))).click()
        self.click_element_id('searchBtn', 3)
        # pyautogui.hotkey('alt', 'tab', interval=0.1)
        # input ('Request your objections. Press <ENTER> to terminate...')
        pyautogui.alert(text=f'Customer : {self.customer}, Length : {self.length}, Amount : {self.amount} \n금액, 건수 검증하고 이의제기 의뢰하세요. \n프로그램 종료합니다.', title='프로세스종료알림', button='OK')

    def mainloop(self):
        now = datetime.now ()
        dt_string = now.strftime ("%Y/%m/%d")
        with open ('Cookies_objection/log.txt', 'a+') as txt:
            txt.write (f'\n{dt_string} Initiated\n')
        start = time.time()
        self.get_workfloor()
        WebDriverWait(self.driver, 7).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="menu"]/div[3]/ul/li[3]/a'))).click()
        min_year = list()
        for feed in self.feeder(self.objset):
            min_year.append (int (feed[5][0] + feed[5][1]))
            self.stopper(3)
            self.creation_loop(feed)
            self.register(feed)
        end = time.time()
        elapsed = 'Elapsed {0:02d}:{1:02d}'.format(*divmod(int(end-start), 60))
        print(elapsed)
        now = datetime.now()
        dt_string = now.strftime("%Y/%m/%d")
        with open('Cookies_objection/log.txt', 'a+') as txt:
            txt.write(f'{dt_string}, {elapsed} Registration Finished\n')
        self.request(min_year)  # todo : verify usability add lines if needed
        self.close()

    def logging(self, feed, stage):
        now = datetime.now()
        dt_string = now.strftime("%Y/%m/%d %H:%M:%S")
        list_no = ', '.join([str(i) for i in feed[6]])
        self.log = dt_string+' '+ feed[0]+' '+ feed[5][3]+'-'+ feed[1][0]+f', ({list_no})' +f', {stage}\n'
        with open('Cookies_objection/log.txt', 'a+') as txt:
            txt.write(self.log)

    def close(self):
        self.driver.delete_all_cookies ()
        self.driver.quit ()

    def __del__(self):
        os.system("taskkill /f /im chromedriver.exe /T")
        gc.collect()


class Pipeline:
    def __init__(self, filename):
        self.filename = filename
        with open ('Cookies_objection/objection.xls', 'rb') as file:
            self.df = pd.read_excel (file)
            self.df.fillna('', inplace=True)
        self.df = self.df[(self.df['Customer Reivew_'] != 'reject')&(self.df['Customer Reivew_'] != 'wait')&(
                self.df['Customer Reivew_'] != 'done')]
        self.customer = self.df.iloc[0]['고객사']
        delimiter = ',,'
        self.df['KEY'] = self.df['고객사'] + delimiter + self.df['E/D'] + delimiter + self.df['ISSUE NO'] +\
                         delimiter + self.df['OBJECTION_'].apply(str) + delimiter + self.df['유형']
        self.keys = set (self.df['KEY'])
        # print(self.keys)
        self.df.set_index ('KEY', inplace=True)
        self.storage = list()

    def isolate(self):
        for key in self.keys:
            val_list = key.split (',,')
            val_list.append ([val_list[2][0:4], val_list[2][4:6], val_list[2][6:],
                              val_list[2][0:4] + '-' + val_list[2][4:6] + '-' + val_list[2][6:]])
            partial_df = self.df.loc[key]
            try:
                val_list.append (sorted (partial_df['LIST'].to_numpy ().tolist ()))
            except AttributeError:
                val_list.append ([partial_df['LIST']])
            val_list.append (round (partial_df['전체'].sum (), 2))
            self.storage.append(val_list)
        self.storage = sorted(self.storage, key = operator.itemgetter(0,1,2,6))
        return self.storage

    def print_example(self):
        example = list(enumerate (self.storage))[0][1]
        print('\nNumber of Objections : {}, Total Amount : {:,.2f}, Total Sequence : {}'.format(
            len(self.df), round(self.df['전체'].sum(),2), len(self.storage)))
        print ('\nFeed Format : {}'.format(example))
        for i, item in enumerate (example):
            print ('Index {} : {}'.format (i, item))

    @staticmethod
    def run():
        filename = 'objection'
        obj = Pipeline (filename)
        data = obj.isolate ()
        obj.print_example ()
        return data


if __name__ == '__main__':
    obj = CustomerObjection(stop=False)
    obj.mainloop()
