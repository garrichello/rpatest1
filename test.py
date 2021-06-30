from selenium import webdriver
from selenium.webdriver.common.by import By 
from selenium.common.exceptions import  TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from datetime import datetime
import xlsxwriter
import math
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from getpass import getpass

def wait_and_click(browser, loc, timeout=30):
    WebDriverWait(browser, timeout).until(
        EC.element_to_be_clickable((By.XPATH, loc))
    )
    browser.find_element_by_xpath(loc).click()

def get_table_data(browser):
    # Get table rows
    loc = f'//table[@class="{data_table_class}"]/tbody/tr'
    WebDriverWait(browser, 30).until(
        EC.visibility_of_element_located((By.XPATH, loc))
    )
    trs = browser.find_elements_by_xpath(loc)[2:]

    # Extract data
    data = []
    for tr in trs:
        tds = tr.find_elements_by_xpath('td')
        current_date = datetime.strptime(tds[0].text, '%d.%m.%Y')
        try:  # Get only rows with valid values. 
            previous_rate = float(tds[1].text.replace(',', '.'))
            current_rate = float(tds[3].text.replace(',', '.'))
            data.append([current_date, current_rate, current_rate-previous_rate])
        except ValueError:  # Rows with bad values such as '-' in a cell are NOT included.
            pass            
    return data

def plural_form(n_lines):
    n = n_lines - n_lines // 10 * 10  # Rightmost digit
    if (n == 1):
        return 'у'  # Одну строк_у_
    elif ((n > 1) and (n < 5)):
        return 'и'  # Две/три/четыре строк_и_
    else:
        return ''  # Много строк_


url = 'http://www.moex.com/'
menu_button_class = 'js-menu-dropdown-button'
menu_element_text = 'Срочный рынок'
agree_button_text = 'Согласен'
ind_rates_text = 'Индикативные курсы'
currency_select_id = 'ctl00_PageContent_CurrencySelect'
usd_rub_option_text = 'USD/RUB - Доллар США к российскому рублю'
eur_rub_option_text = 'EUR/RUB - Евро к российскому рублю'
data_table_class = 'tablels'
xlsx_file_name = 'test.xlsx'

today = datetime.today()
first_day_of_month = datetime(today.year, today.month, 1)

browser = webdriver.Firefox()  # Create browser window

browser.get(url)  # Go to the url

# Menu
loc = f'//a[contains(@class, "{menu_button_class}")]'
wait_and_click(browser, loc)

# Derivatives market
loc = f'//div[@class="item"]/a[contains(text(), "{menu_element_text}")]'
wait_and_click(browser, loc)

# Disclamer window
loc = f'//a[contains(text(), "{agree_button_text}")]'
try:
    wait_and_click(browser, loc)
    WebDriverWait(browser, 10).until(
        EC.invisibility_of_element_located((By.ID, 'disclaimer-modal'))
    )
except TimeoutException:
    pass

# Indicative rate
loc = f'//div[contains(text(), "{ind_rates_text}")]'
wait_and_click(browser, loc)

# Select currency USD
WebDriverWait(browser, 30).until(
    EC.element_to_be_clickable((By.ID, currency_select_id))
)
currency_select = Select(browser.find_element_by_id(currency_select_id))
currency_select.select_by_visible_text(usd_rub_option_text)

# Get USD table data
data_usd = get_table_data(browser)

# Select currency EUR
currency_select.select_by_visible_text(eur_rub_option_text)

# Get EUR table data
data_eur = get_table_data(browser)

browser.close()

# Prepare Excel-file
workbook = xlsxwriter.Workbook(xlsx_file_name)
worksheet = workbook.add_worksheet()
date_format = workbook.add_format()
date_format.set_num_format(14)
finance_format = workbook.add_format()
finance_format.set_num_format(44)
cross_format = workbook.add_format()
cross_format.set_num_format('0.0000')
header_format = workbook.add_format()
header_format.set_align('center')

# Write header of the table, adapting columns width
worksheet.write('A1', 'Дата', header_format)
worksheet.set_column('A:A', 10)
worksheet.write('B1', 'Курс', header_format)
int_part_len = math.ceil(math.log10((max([e[1] for e in data_usd]))))  # Max value int part length
max_usd_width = int_part_len + 7  # Decimal point + 2 digits + sign + currency symbol + 2 spaces
worksheet.set_column('B:B', max_usd_width)
worksheet.write('C1', 'Изменение', header_format)
worksheet.set_column('C:C', 10)
worksheet.write('D1', 'Дата', header_format)
worksheet.set_column('D:D', 10)
worksheet.write('E1', 'Курс', header_format)
int_part_len = math.ceil(math.log10((max([e[1] for e in data_eur]))))  # Max value int part length
max_eur_width = int_part_len + 7  # Decimal point + 2 digits + sign + currency symbol + 2 spaces
worksheet.set_column('B:B', max_eur_width)
worksheet.write('F1', 'Изменение', header_format)
worksheet.set_column('F:F', 10)
worksheet.write('G1', 'Кросс-курс', header_format)
worksheet.set_column('G:G', 10)

# Write data to the table
row = 1
col = 0
for usd, eur in zip(data_usd, data_eur):
    if (usd[0] >= first_day_of_month):
        worksheet.write(row, col,   usd[0], date_format)
        worksheet.write(row, col+1, usd[1], finance_format)
        worksheet.write(row, col+2, usd[2], finance_format)
        worksheet.write(row, col+3, eur[0], date_format)
        worksheet.write(row, col+4, eur[1], finance_format)
        worksheet.write(row, col+5, eur[2], finance_format)
        worksheet.write(row, col+6, eur[1] / usd[1], cross_format)
        row += 1

workbook.close()

# Send E-mail message
subject = 'Отчет о курсах валют'
body = f'Отчет содержит {row} строк' + plural_form(row) + ', включая заголовок.'
sender_email = 'oig@sibmail.com'
receiver_email = 'oig@sibmail.com'
password = getpass('E-mail password:')

msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = subject
msg.attach(MIMEText(body, 'plain'))

with open(xlsx_file_name, 'rb') as file:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(file.read())

encoders.encode_base64(part)

part.add_header(
    'Content-Disposition',
    f'attachment; filename={xlsx_file_name}',
)

msg.attach(part)
text = msg.as_string()

with smtplib.SMTP('smtp.sibmail.com', 25) as server:
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, text)
    print('E-mail sent.')
