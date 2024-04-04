import atexit
import datetime
import logging
import os
import shutil
import string
import tempfile
import time
import configparser

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import TimeoutException

config = configparser.RawConfigParser()
config.read('default.ini')
config.read('properties.ini')

module_logger = logging.getLogger('addison-oneclick-downloader')
selenium_logger = logging.getLogger('selenium')
log_folder = config['debug']['log_folder']
module_log_level = config['debug']['log_level']
selenium_log_level = config['debug']['selenium_log_level']

dry_run = config['debug']['dry_run'] == 'True'

url = config['login']['url']
client_number = config['login']['client_number']
username = config['login']['username']
password = config['login']['password']

download_folder = tempfile.mkdtemp()
file_destination = config['file']['file_destination']
file_rename = config['file']['rename'] == 'True'
file_search_period = config['file']['search_period']
file_search_period_command = 'command_week' if file_search_period == '7' else 'command_month'

wait_timeout = int(config['settings']['wait_timeout'])
download_timeout = int(config['settings']['download_timeout'])
headless = config['settings']['headless'] == 'True'

chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory': download_folder}
chrome_options.add_experimental_option('prefs', prefs)
if headless:
    chrome_options.add_argument('--headless')
driver = webdriver.Chrome(options=chrome_options)


def wait_for_element(xpath):
    try:
        WebDriverWait(driver, timeout=wait_timeout).until(
            ec.visibility_of_element_located((By.XPATH, xpath)))
    except TimeoutException:
        module_logger.error(xpath + ' not loaded after waiting for ' + str(wait_timeout) + ' seconds')
        quit(1)


def login():
    module_logger.info('login')
    driver.find_element(by=By.ID, value='ClientNumber').send_keys(client_number)
    driver.find_element(by=By.ID, value='Username').send_keys(username)
    driver.find_element(by=By.ID, value='Password').send_keys(password)
    driver.find_element(by=By.XPATH, value="//button[@value='login']").click()


def select_week_filter():
    # in my documents view wait for table to render
    wait_for_element("//md-part-modulemenu[@data-part='menu']")

    # get filter selection
    filter_text = driver.find_element(by=By.XPATH, value="//ul[@class='nav navbar-nav']/li[2]/a[@data-toggle='dropdown'][1]/div/span[@class='dropdown-toggle-label']").text
    if file_search_period not in filter_text:
        # reset filter
        close_button = driver.find_element(by=By.XPATH, value="//ul[@class='nav navbar-nav']/li[2]/a[@data-toggle='dropdown'][1]/div/span[@title='close']")
        if close_button.is_displayed():
            close_button.click()

        # open filter dropdown
        driver.find_element(by=By.XPATH, value="//ul[@class='nav navbar-nav']/li[2]/a[@data-toggle='dropdown'][1]").click()

        # select last 7 days
        driver.find_element(by=By.XPATH, value="//ul[@class='nav navbar-nav']/li[2]/ul/li/a[@data-i18n='app:modules.documentstorage.%s']" % file_search_period_command).click()


def download_payslip(index):
    index += 1
    # wait for table to load
    wait_for_element("//tbody[@role='rowgroup']/tr[%d]" % index)

    # click first row
    driver.find_element(by=By.XPATH, value="//tbody[@role='rowgroup']/tr[%d]" % index).click()

    # open three-dot menu
    wait_for_element("//tbody[@role='rowgroup']/tr[1]/td/div/wk-dropdown/button[@type='button']")
    driver.find_element(by=By.XPATH, value="//tbody[@role='rowgroup']/tr[1]/td/div/wk-dropdown/button[@type='button']").click()
    # click download
    if not dry_run:
        driver.find_element(by=By.XPATH, value="//tbody[@role='rowgroup']/tr[1]/td/div/wk-dropdown/wk-dropdown-menu/render-target/div/slot/dl/dd[3]/button").click()
        module_logger.info('downloading file to ' + download_folder)
        i = 0
        while len(os.listdir(download_folder)) == 0:
            time.sleep(1)
            if i < download_timeout:
                i += 1
            else:
                if len(os.listdir(download_folder)) != 0:
                    module_logger.error('timeout while downloading')
                    exit(1)
    else:
        module_logger.debug('dry-run enabled: skipped downloading file')
    # navigate back to overview
    driver.find_element(by=By.XPATH, value="//nav[@class='wk-breadcrumb wk-breadcrumb-full']/ol/li[1]/a").click()


def copy_and_rename_payslip_document():
    if dry_run:
        return
    file_path = max([download_folder + "\\" + f for f in os.listdir(download_folder)], key=os.path.getctime)
    file_name = os.path.split(file_path)[1].split('_')

    output_file_path = file_destination
    if file_rename:
        accounting_period = file_name[4]
        accounting_period_datetime = datetime.datetime(int(accounting_period[0:4]), int(accounting_period[4:6]), 1)
        create_date = file_name[5]
        create_time = file_name[6]
        create_date_time = datetime.datetime(int(create_date[0:4]), int(create_date[4:6]), int(create_date[6:]), int(create_time[0:2]), int(create_time[2:4]), int(create_time[4:]))

        placeholders = {'accounting_period': accounting_period_datetime, 'create_date': create_date_time}

        field_names = [name for text, name, spec, conv in string.Formatter().parse(file_destination) if name is not None]
        for placeholder in field_names:
            placeholder_name = placeholders[placeholder.split('%', 1)[0]]
            placeholder_date = '%' + placeholder.split('%', 1)[1]
            output_file_path = output_file_path.replace('{' + placeholder + '}', placeholder_name.strftime(placeholder_date))
    copy_payslip_document(file_path, output_file_path)
    module_logger.info("created file: " + output_file_path)


def copy_payslip_document(origin, destination):
    module_logger.info('copy payslip from ' + origin + ' to ' + destination)
    os.makedirs(os.path.dirname(destination), exist_ok=True)
    shutil.copy(origin, destination)


def download_payslips():
    documents_list = driver.find_elements(by=By.XPATH, value="//div[@class='k-grid-content k-auto-scrollable']/table/tbody/tr")
    for i, doc in enumerate(documents_list):
        download_payslip(i)
        copy_and_rename_payslip_document()


def register_logger(logger, logger_name, folder_path, log_level, formatter=None):
    os.makedirs(log_folder, exist_ok=True)

    module_file_handler = logging.FileHandler(filename=('{0}' + os.sep + '{1}.log').format(folder_path, logger_name + '_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S')))
    if formatter is not None:
        module_file_handler.setFormatter(formatter)
    logger.setLevel(logging.getLevelName(level=log_level))
    logger.addHandler(module_file_handler)


def setup():
    if config['debug']['logging_enabled'] == 'True':
        formatter = logging.Formatter('%(asctime)s %(levelname)s %(lineno)d:%(filename)s(%(process)d) - %(message)s')
        register_logger(module_logger, 'addison_oneclick_downloader', log_folder, module_log_level, formatter)
    if config['debug']['selenium_logging_enabled'] == 'True':
        register_logger(selenium_logger, 'selenium', log_folder, selenium_log_level)
    atexit.register(cleanup)


def cleanup():
    module_logger.debug('remove temp folder ' + download_folder)
    shutil.rmtree(download_folder)


setup()
driver.get(url)
login()
select_week_filter()
download_payslips()
driver.close()
