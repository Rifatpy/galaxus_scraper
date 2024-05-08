"""
www.galaxus.ch webstite scraper.
Description: Read product ids from a list, check them in galaxus website and extract some required data. Then store them in a single csv file.
Date: 01/05/2024
client: Arif
"""

import sys
import time
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QLineEdit, QTextBrowser, QVBoxLayout, QWidget, QMessageBox
from PyQt5.QtCore import QThreadPool, QRunnable, QObject, pyqtSignal, QMutex, QMutexLocker
from selenium.common.exceptions import ElementClickInterceptedException
import random
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import csv

class Communicate(QObject):
    message_signal = pyqtSignal(str)

class SharedCounter:
    def __init__(self):
        self.value = 1
        self.lock = QMutex()

    def increment(self, amount=1):
        with QMutexLocker(self.lock):
            self.value += amount

    def get_value(self):
        with QMutexLocker(self.lock):
            return self.value

def chunk_data(data, num_chunks):
    chunk_size = len(data) // num_chunks
    chunks = [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]
    return chunks

# Removing checked mail from all mails
def mail_removal(array1, array2):
    result = list(set(array1) - set(array2))
    return result

def unchecked_mail_write(unchecked_hotmails):
    with open('UNCHECKED_DATA.txt', 'a') as file:
        for hotmail in unchecked_hotmails:
            file.write(hotmail)
            file.write('\n')

def worker_function(args):
    thread_number, chunk, communicate, counter = args
    checked_mails = []
    try:
        # Set up Chrome options
        options = webdriver.ChromeOptions()
        #options.add_argument("--headless=new")
        #options.add_argument(f"user-agent={my_user_agent}")
        prefs = {"credentials_enable_service": False,
                "profile.password_manager_enabled": False}
        options.add_experimental_option("prefs", prefs)
        # Adding argument to disable the AutomationControlled flag 
        options.add_argument("--disable-blink-features=AutomationControlled") 
        # Exclude the collection of enable-automation switches 
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        # Turn-off userAutomationExtension 
        options.add_experimental_option("useAutomationExtension", False) 
        driver = webdriver.Chrome(options=options)
        #driver.maximize_window()
        # Changing the property of the navigator value for webdriver to undefined 
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        actions = ActionChains(driver)
        # Perform the required work here
        for single_data in chunk:
            if communicate.stopped:  # Check if the stop event is set
                unchecked_mails = mail_removal(chunk, checked_mails)
                unchecked_mail_write(unchecked_mails)
                return
            main_work(thread_number, single_data, communicate, driver, actions)
            checked_mails.append(single_data)

            # Increment the counter
            counter.increment()

    except Exception as e:
        # Write unchecked mails to a file before closing the thread
        unchecked_mails = mail_removal(chunk, checked_mails)
        unchecked_mail_write(unchecked_mails)
        # Handle any exceptions that may occur during Selenium actions
        """error_message = f'[THREAD-{thread_number+1}] : Error - {str(e)}'
        communicate.message_signal.emit(error_message)"""

    finally:
        # Make sure to properly close the WebDriver instance
        if 'driver' in locals():
            driver.quit()

    

def main_work(thread_number, single_data, communicate, driver, actions):
    driver.set_window_size(800, 1000)
    driver.get(f'https://www.galaxus.ch/de/search?q={single_data}')
    try:

        try:
            WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.XPATH, "//article[@class='sc-d9cbca0f-1 gFbGSC']")))
        except:
            WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.XPATH, "//article[@class='sc-b3dc936d-1 fOimcc']")))

        try:
            search_results = driver.find_element(By.XPATH, "//article[@class='sc-d9cbca0f-1 gFbGSC']")
        except:
            search_results = driver.find_element(By.XPATH, "//article[@class='sc-b3dc936d-1 fOimcc']")
        product_link = search_results.find_element(By.TAG_NAME, 'a').get_attribute('href')

        driver.get(product_link)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//span[@class='sc-d963cb62-0 btKhJv']")))
        time.sleep(2)
        driver.execute_script("document.body.style.zoom='70%'")

        # Extracting data
        # Product Name
        product_name = driver.find_element(By.XPATH, "//span[@class='sc-d963cb62-0 btKhJv']").text
        # Producttype, katagory, hubkatagory
        category_list = driver.find_elements(By.XPATH, "//li[@class='sc-f40471c7-3 eXdWaR']")
        
        try:
            product_type = category_list[-2].text
        except:
            product_type = ' '
        
        try:
            katagory = category_list[-3].text
        except:
            katagory = ' '

        try:
            hubkatagory = category_list[-4].text
        except:
            hubkatagory = ' '


        # Description
        try:
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-test='ShowMoreToggleButton-description']")))
                show_more_btn = driver.find_element(By.XPATH, "//button[@data-test='ShowMoreToggleButton-description']")
                driver.execute_script("arguments[0].scrollIntoView();", show_more_btn)
                time.sleep(1)
                driver.execute_script("arguments[0].click()", show_more_btn)
                time.sleep(2)
            except:
                pass
            descriptions = driver.find_elements(By.XPATH, "//div[@class='sc-8e9bd633-0 ciyGDx']")
            for des in descriptions:
                description = des.text
                if description.find("Beschreibung")!=-1:
                    break
            
            try:
                description = description.replace("Beschreibung", "")
            except:
                pass

            try:
                description = description.replace("\n", " ")
            except:
                pass

            try:
                description = description.replace("Weniger anzeigen", "")
            except:
                pass
        except Exception as e:
            communicate.message_signal.emit(str(e))

    

        # Farbe, Material, Gewitch, Lange, Breite,  Hohe
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-test='showMoreButton-specifications']")))
        show_more_btn2 = driver.find_element(By.XPATH, "//button[@data-test='showMoreButton-specifications']")
        driver.execute_script("arguments[0].scrollIntoView();", show_more_btn2)
        driver.execute_script("arguments[0].click()", show_more_btn2)
        time.sleep(2)
        all_table = driver.find_elements(By.XPATH, "//div[@class='sc-36b31da0-3 ifiYZe']")
        farbe = ' '
        material = ' '
        gewicht = ' '
        lange = ' '
        breite = ' '
        hohe = ' '
        for table in all_table:
            table_data = table.text

            # Gewicht
            if table_data.find("Produktdimensionen")!=-1 or table_data.find("Verpackungsdimensionen")!=-1:
                driver.execute_script("arguments[0].scrollIntoView();", table)
                time.sleep(1)
                if len(gewicht)<2:
                    gewicht_data_range = table_data.split("\n")
                    for i in range(0, len(gewicht_data_range)):
                        g_data = gewicht_data_range[i]
                        if g_data.find("Gewicht")!=-1:
                            gewicht = gewicht_data_range[i+1]
                            break

            # Farbe
            if table_data.find("Farbe")!=-1:
                if len(farbe)<=2:
                    gewicht_data_range = table_data.split("\n")
                    farbe = gewicht_data_range[2]
            # Material
            table_heading = table.find_element(By.TAG_NAME, 'h3').text
            if table_heading.find("Material")!=-1:
                gewicht_data_range = table_data.split("\n")
                material = gewicht_data_range[2]

            
            # Länge
            if table_data.find("Verpackungsdimensionen")!=-1 or table_data.find("Produktdimensionen")!=-1:
                if len(lange)<=1:
                    gewicht_data_range = table_data.split("\n")
                    for i in range(0, len(gewicht_data_range)):
                        g_data = gewicht_data_range[i]
                        if g_data.find("Länge")!=-1:
                            lange = gewicht_data_range[i+1]
                            break

    
            # Breite
            if table_data.find("Verpackungsdimensionen")!=-1 or table_data.find("Produktdimensionen")!=-1:
                if len(breite)<=1:
                    gewicht_data_range = table_data.split("\n")
                    for i in range(0, len(gewicht_data_range)):
                        g_data = gewicht_data_range[i]
                        if g_data.find("Breite")!=-1:
                            breite = gewicht_data_range[i+1]
                            break

            
            # Höhe
            if table_data.find("Verpackungsdimensionen")!=-1 or table_data.find("Produktdimensionen")!=-1:
                if len(hohe)<=1:
                    gewicht_data_range = table_data.split("\n")
                    for i in range(0, len(gewicht_data_range)):
                        g_data = gewicht_data_range[i]
                        if g_data.find("Höhe")!=-1:
                            hohe = gewicht_data_range[i+1]
                            break

            
        info = [single_data, product_name, product_type, katagory, hubkatagory, description, farbe, material, gewicht, lange, breite, hohe]
        message = f'Thread {thread_number+1} - {info}'
        communicate.message_signal.emit(str(message))
        write_to_file(info)

    except:
        message = f'Thread {thread_number+1} - {single_data}: NOT FOUND'
        communicate.message_signal.emit(str(message))
        info = [single_data, "", "", "", "", "", "", "", "", "", "", ""]
        write_to_file(info)

def write_to_file(data):
    with open('OUTPUT.csv', 'a', encoding='UTF8') as file:
        writer = csv.writer(file)
        writer.writerow(data)
        
class Worker(QRunnable):
    def __init__(self, args):
        super().__init__()
        self.args = args

    def run(self):
        worker_function(self.args)

class CommunicateWithStop(Communicate):
    stopped = False

class EmailCheckerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.upload_mail_file_button = QPushButton('UPLOAD DATA FILE', self)
        self.threads_label = QLabel('THREADS:', self)
        self.threads_input = QLineEdit(self)
        self.start_button = QPushButton('START', self)
        self.stop_button = QPushButton('STOP', self)

        self.output_screen = QTextBrowser(self)

        self.mail_file_data = []

        self.is_start_clicked = False

        self.init_ui()

        self.communicate = CommunicateWithStop()
        self.communicate.message_signal.connect(self.update_output_screen)

        self.thread_pool = QThreadPool.globalInstance()
        self.counter = SharedCounter()

        # Set your expiration date here
        self.expiry_date = datetime(2024, 12, 31)

        # Check if the expiration date has passed
        if datetime.now() > self.expiry_date:
            self.show_expiry_message()
            self.disable_buttons()

    def show_expiry_message(self):
        QMessageBox.warning(self, "Expired", "Your software has expired. Please renew your license.")

    def disable_buttons(self):
        self.upload_mail_file_button.setEnabled(False)
        self.threads_input.setEnabled(False)
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.output_screen.setEnabled(False)


    def init_ui(self):
        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)

        layout.addWidget(self.upload_mail_file_button)
        layout.addWidget(self.threads_label)
        layout.addWidget(self.threads_input)
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_button)
        layout.addWidget(self.output_screen)

        self.setCentralWidget(central_widget)

        self.upload_mail_file_button.clicked.connect(self.upload_mail_file)
        self.start_button.clicked.connect(self.start_processing)
        self.stop_button.clicked.connect(self.stop_processing)

        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(False)

        # Set colors
        self.setStyleSheet("background-color: #b2b2b2; color: black;")
        self.output_screen.setStyleSheet("background-color: black; color: white;")

        self.setGeometry(100, 100, 600, 400)
        self.setWindowTitle('Galaxus.ch/de Scraper')
        self.show()

    def upload_mail_file(self):
        file_dialog = QFileDialog.getOpenFileName(self, 'Open Mail File', '', 'Text Files (*.txt)')
        if file_dialog[0]:
            if file_dialog[0].endswith('.txt'):
                with open(file_dialog[0], 'r') as file:
                    self.mail_file_data = file.read().splitlines()
                    self.output_screen.append('Data File Uploaded')
                    self.output_screen.append(f'Total Data: {len(self.mail_file_data)}')
                    self.start_button.setEnabled(True)
            else:
                self.output_screen.append('Error: Please select a .txt file for Mail File')


    def start_processing(self):
        if not self.is_start_clicked:
            try:
                num_threads = int(self.threads_input.text())
                if num_threads <= 0:
                    raise ValueError("Number of threads must be a positive integer.")

                self.is_start_clicked = True
                self.start_button.setEnabled(False)
                self.stop_button.setEnabled(True)

                # Reset the counter
                self.counter.value = 0

                # Create your own thread pool with the desired maximum thread count
                thread_pool = QThreadPool.globalInstance()
                thread_pool.setMaxThreadCount(num_threads)

                # Prepare arguments for worker_function
                args_list = [
                    (i, chunk, self.communicate, self.counter)
                    for i, chunk in enumerate(chunk_data(self.mail_file_data, num_threads))
                ]

                # Submit tasks to the thread pool
                for args in args_list:
                    worker = Worker(args)
                    thread_pool.start(worker)

            except ValueError as e:
                self.output_screen.append(f'Error: {str(e)}')

    def stop_processing(self):
        if self.is_start_clicked:
            # Set the stop event for all threads
            self.communicate.stopped = True

            # Reset the flag and buttons
            self.is_start_clicked = False
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)

    def update_output_screen(self, message):
        total_checked = self.counter.get_value()
        if message.find("NOT EXIST"):
            self.output_screen.append(f"[{total_checked}] {message}")
        else:
            self.output_screen.append(f"[{total_checked}] {message}")



    def closeEvent(self, event):
        # Wait for all threads to finish before closing the application
        self.thread_pool.waitForDone()
        event.accept()


    


if __name__ == '__main__':
    # Writing header to output.csv file
    with open('OUTPUT.csv', 'w', encoding='UTF8') as file:
        writer = csv.writer(file)
        header = ["Gtin", "Product Name", "Producttyp", "Kategorie", "HUbkatagorie", "Beschreibung", "Farbe", "Material", "Gewicht", "Länge", "Breite", "Höhe"]
        writer.writerow(header)
        file.close()

    app = QApplication(sys.argv)
    email_checker_app = EmailCheckerApp()
    sys.exit(app.exec_())
#pyinstaller --onefile --noconsole your_script.py