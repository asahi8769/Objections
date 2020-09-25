from utils.config import *
from utils.functions import path_find
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
import pandas as pd


warnings.filterwarnings ('ignore')


class CustomerObjection:
    GC_DRIVER = os.path.join(os.getcwd(), 'driver', 'chromedriver.exe')
    CHROME_OPTIONS = Chrome_options()
    CHROME_OPTIONS.add_argument("--start-maximized")
    CHROME_OPTIONS.add_argument("--disable-extensions")
    CHROME_OPTIONS.add_argument("--incognito")

    def __init__(self, stop=False):
        objections = Pipeline()
        # self.idx = []
        self.log = None
        self.df = objections.df
        self.tot_seq = len(objections.df)
        self.objset = objections.objection_generator()
        self.customer = None
        self.driver = webdriver.Chrome(self.GC_DRIVER, options=self.CHROME_OPTIONS)
        self.driver.get(URL)
        self.sequence = 0
        self.stop = stop
        self.length = 0
        self.amount = 0

    def click_element_id(self, ID, sec):
        try:
            element = WebDriverWait (self.driver, sec).until(
                EC.element_to_be_clickable ((By.ID, ID)))
            element.click()
            return element
        except TimeoutException:
            pass

    def log_in(self):
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_3_form_ImageViewer00', 3)
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_3_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_2_form_ImageViewer00', 1)
        self.click_element_id('mainframe_VFrameSet_LoginFrame_CommLgds010_1_form_ImageViewer00', 1)
        self.click_element_id ('mainframe_VFrameSet_LoginFrame_CommLgds010_0_form_ImageViewer00', 1)
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

    def input_year_month(self, issuemonth, wait):
        """Redesigned on 2020.03.30"""
        seq_m = self.click_element_id('K_RPYM_WNTI_MM', 3)
        seq_m.send_keys(Keys.DELETE)
        seq_y = self.click_element_id('K_RPYM_WNTI_YY',3)
        seq_y.send_keys(issuemonth[0:4])
        seq_m.send_keys(issuemonth[4:6])
        time.sleep(wait)

    def counter(self, num):
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
        if self.df.loc[feed[0] + ',,' +feed[1]+ ',,'+feed[2]+ ',,' +feed[3]+ ',,' +feed[4]].iloc[0]['Customer Reivew_'] == 'Created':
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
        self.update_hist(feed, 'Created')

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
        if self.sequence == 1:
            pyautogui.hotkey('alt', 'tab', 'left')
        ans = pyautogui.confirm(text=f'{self.sequence}/{self.tot_seq} 이의제기 등록합니다. 교류클레임 여부 확인하세요. \n건수: {len(feed[6])}, 금액: {feed[-1]}. \n사유: {feed[3]}.', title='등록확인', buttons=['OK', 'NO'])
        if ans.upper() != 'NO':
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
            letter_box_1 = self.click_element_id ('D_MKOB_TITL_NM', 5)
            letter_box_1.send_keys(str(feed[3]))
            letter_box_2 = self.click_element_id('D_MKOB_DTL_SBC', 5)
            letter_box_2.send_keys(str(feed[3]))
            self.click_element_id('dlg_AppealSaveBtn', 5)
            time.sleep(0.7)
            pyautogui.press('enter')
            time.sleep(0.5)
            self.logging (feed, 'Registered')
            self.update_hist(feed, 'Registered')
        else:
            self.driver.switch_to.window(self.driver.window_handles[1])
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
        for feed in self.objset:
            self.customer = feed[0].upper()
            min_year.append (int (feed[5][0] + feed[5][1]))
            self.counter(3)
            self.creation_loop(feed)
            self.register(feed)
        end = time.time()
        elapsed = 'Elapsed {0:02d}:{1:02d}'.format(*divmod(int(end-start), 60))
        print(elapsed)
        now = datetime.now()
        dt_string = now.strftime("%Y/%m/%d")
        with open('Cookies_objection/log.txt', 'a+') as txt:
            txt.write(f'{dt_string}, {elapsed} Registration Finished\n')
        self.request(min_year)
        pyautogui.alert(
            text=f'Customer : {self.customer}, Length : {self.length}, Amount : {self.amount}, \n소요시간 : {elapsed} \n금액, 건수 검증하고 이의제기 의뢰하세요. \n의뢰 한 후 브라우저를 닫으세요.',
            title='프로세스종료알림', button='OK')
        self.df.to_excel('Cookies_objection\objection.xls', index=False)
        os.startfile('Cookies_objection\objection.xls')
        input ('Press <ENTER> to terminate...')
        self.close()

    def logging(self, feed, stage):
        now = datetime.now()
        dt_string = now.strftime("%Y/%m/%d %H:%M:%S")
        list_no = ', '.join([str(i) for i in feed[6]])
        self.log = dt_string+' '+ feed[0]+' '+ feed[5][3]+'-'+ feed[1][0]+', '+ feed[3]+f', ({list_no})' +f', {stage}\n'
        with open('Cookies_objection/log.txt', 'a+') as txt:
            txt.write(self.log)

    def update_hist(self, feed, stage):
        self.df.at[feed[0] + ',,' +feed[1]+ ',,'+feed[2]+ ',,' +feed[3]+ ',,' +feed[4], 'Customer Reivew_'] = stage

    @classmethod
    def run(cls):
        obj = cls(stop=False)
        try :
            obj.mainloop()
        except Exception as e :
            obj.df.to_excel('Cookies_objection\objection.xls', index=False)
            os.startfile('Cookies_objection\objection.xls')
            obj.close()

    def close(self):
        try:
            self.driver.delete_all_cookies ()
            self.driver.quit ()
        except Exception as e:
            pass

    def __del__(self):
        try :
            self.driver.delete_all_cookies()
        except Exception as e:
            print('Screen terminated according to nominal procedure.')
        os.system("taskkill /f /im chromedriver.exe /T")
        gc.collect()


class Pipeline:
    def __init__(self):
        self.filename = path_find('objection.xls', os.path.abspath(os.pardir))
        with open (self.filename, 'rb') as file:
            self.df = pd.read_excel (file)
            self.df.fillna('', inplace=True)

        self.customer = self.df.iloc[0]['고객사']
        delimiter = ',,'
        self.df['KEY'] = self.df['고객사'] + delimiter + self.df['E/D'] + delimiter + self.df['ISSUE NO'] +\
                         delimiter + self.df['OBJECTION_'].apply(str) + delimiter + self.df['유형']
        self.df.set_index('KEY', inplace=True)
        self.df_ = self.df[~(self.df['Customer Reivew_'].str.contains('reject')) &
                          ~(self.df['Customer Reivew_'].str.contains('wait')) &
                          ~(self.df['Customer Reivew_'].str.contains('pending')) &
                          ~(self.df['Customer Reivew_'].str.contains('denied')) &
                          ~(self.df['Customer Reivew_'].str.contains('wrong')) &
                          ~(self.df['Customer Reivew_'].str.contains('deny')) &
                          ~(self.df['Customer Reivew_'].str.contains('done')) &
                          ~(self.df['Customer Reivew_'].str.contains('cancel')) &
                          ~(self.df['Customer Reivew_'].str.contains('기각')) &
                          ~(self.df['Customer Reivew_'].str.contains('Registered'))
        ]
        self.keys = set (self.df_.index)
        self.storage = list()
        self.data = self.isolate()
        self.print_example()

    def isolate(self):
        for key in self.keys:
            # print(key)
            val_list = key.split (',,')
            val_list.append ([val_list[2][0:4], val_list[2][4:6], val_list[2][6:],
                              val_list[2][0:4] + '-' + val_list[2][4:6] + '-' + val_list[2][6:]])
            partial_df = self.df_.loc[key]
            try:
                val_list.append (sorted (partial_df['LIST'].to_numpy ().tolist ()))
            except AttributeError:
                val_list.append ([partial_df['LIST']])
            val_list.append (round (partial_df['전체'].sum (), 2))
            self.storage.append(val_list)
        self.storage = sorted(self.storage, key = operator.itemgetter(0,1,2,3,6))
        return self.storage

    def print_example(self):
        example = list(enumerate (self.storage))[0][1]
        print('\nNumber of Objections : {}, Total Amount : {:,.2f}, Total Sequence : {}'.format(
            len(self.df_), round(self.df_['전체'].sum(),2), len(self.storage)))
        print ('\nFeed Format : {}'.format(example))
        for i, item in enumerate (example):
            print ('Index {} : {}'.format (i, item))

    def objection_generator(self):
        for datum in self.data:
            yield datum


if __name__ == '__main__':
    # CustomerObjection.run()
    # data = Pipeline()
    # df = data.df
    # print(df)
    # keys = ['HAOS,,EXP,,2020081WW,,Test1,,A', 'HAOS,,EXP,,2020081WW,,Test3,,C']
    # df.at[keys, 'Customer Reivew_'] = 'Registered'
    # print(df.loc[keys].iloc[0]['Customer Reivew_'])
    # print(df.loc[keys]['VENDORCODE'])
    print(os.path.abspath(os.pardir))
