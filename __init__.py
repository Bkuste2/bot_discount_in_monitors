import os
import smtplib
import pandas as pd
from colorama import Fore
from dotenv import load_dotenv
from selenium import webdriver
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


load_dotenv()
email_password = os.getenv('email_password')
email_user = os.getenv('email_user')

class Scrappy:

    def start(self):
        self.data_scraping()

    def configure_and_open_web_page(self):
        print(Fore.YELLOW, "Configuring Driver... \n")

        url = 'https://www.kabum.com.br/computadores/monitores?page_number=1&page_size=20&facet_filters=eyJoYXNfb2ZmZXIiOlsidHJ1ZSJdfQ==&sort=price'
        service = Service(ChromeDriverManager().install())
        options = Options()
        options.add_argument('--headless')

        driver = webdriver.Chrome(service=service, options=options)
        driver.get(url)

        return driver

    def convert_selenium_type_to_text(self, list, ):
        list_text = [item.text for item in list]
        return list_text

    def data_scraping(self):
        driver = self.configure_and_open_web_page()

        print(Fore.LIGHTGREEN_EX, "Scraping Data... \n")

        titles = driver.find_elements('xpath', '//main/div/a/div/button/div/h2/span')
        links = driver.find_elements('xpath', '//main/div/a')
        regular_prices = driver.find_elements('xpath', '//main/div/a/div/div/span[1]')
        sale_prices = driver.find_elements('xpath', '//main/div/a/div/div/span[2]')

        titles_text = self.convert_selenium_type_to_text(titles)
        links_text = [item.get_attribute('href') for item in links]
        regular_prices_text = self.convert_selenium_type_to_text(regular_prices)
        sale_prices_text = self.convert_selenium_type_to_text(sale_prices)

        self.write_csv_file(titles_text, regular_prices_text, sale_prices_text, links_text)
        self.send_mail()

    def write_csv_file(self, titles, regular_prices, sale_prices, link):
        print(Fore.BLUE, "Creating Sheets")

        df = pd.DataFrame({
            'Title': titles,
            'Preço Normal': regular_prices,
            'Preço em Promoção': sale_prices,
            'Link Produto': link
        })

        df.to_csv('dados.csv', index=False)
        print(Fore.BLUE, "Sheets created with Success!")

    def send_mail(self):
        print(Fore.CYAN, 'Sending Mail...')
        msg = MIMEMultipart()

        message = "Hey, I'm sending you an attached spreadsheet that contains all the discounted monitors " \
                  "on the kabum website, attached are the promotions with the best prices"

        # setup the parameters of the message
        password = email_password
        msg['From'] = email_user
        msg['To'] = input('Who would you like to send this email to? ')
        msg['Subject'] = "Hey bro, look at all the monitors with promotion on the kabum website"
        msg.attach(MIMEText(message, 'plain'))

        # Add the attachment
        with open('./dados.csv', 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='csv')
            attachment.add_header('Content-Disposition', 'csv', filename='monitor_with_discount_in_kabum.csv')
            msg.attach(attachment)

        # Connect to the SMTP server and send the email
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(msg['From'], password)
            server.sendmail(msg['From'], msg['To'], msg.as_string())
            server.quit()

        print("successfully sent email to %s:" % (msg['To']))


monitors_with_discount = Scrappy()
monitors_with_discount.start()
