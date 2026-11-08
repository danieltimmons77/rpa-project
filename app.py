# INSERT ANSWER FOR QUESTION 1 HERE
required_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:    
    directory = single_input("Please tell me the directory where the files are uploaded.", mask=False, width=400, height=200)
    if directory[-1] != '\\':
        directory = directory + '\\'


    if os.path.isdir(directory):
        files_in_dir = os.listdir(directory)
        missing = {f for f in required_files if f not in files_in_dir}
    if not missing:
        break
        show_message('Error', 'Missing files: ' + ', '.join(missing) + '. Pick another directory.')
    else:
        show_message("Error", "Something is not right. Please re-enter the directory.")

print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
descriptions_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}
for index, row in descriptions_df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']
prices_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
product_prices = {}
for index, row in prices_df.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

print(product_descriptions)
print(product_prices)

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    workbook = openpyxl.load_workbook(directory + filename, data_only=True)
    sheet = workbook['OrderForm']

    date = sheet['C9'].value.strftime('%m/%d/%Y')
    email = sheet['C13'].value
    name = sheet['C14'].value
    company = sheet['C15'].value
    address = sheet['C16'].value
    city = sheet['C17'].value
    state = sheet['C18'].value
    zipcode = str(sheet['C19'].value)

    order_data = {
    'InvoiceNum': str(invoice_num),
    'Date': date,
    'CustomerName': name,
    'CustomerCompany': company,
    'CustomerAddress1': address,
    'CustomerAddress2': city + ', ' + state + ' ' + zipcode,
    }

    total = 0
    row = 24
    item = 0
    while sheet['B' + str(row)].value is not None:
        unit = sheet['B' + str(row)].value
        quantity = sheet['C' + str(row)].value
        price = product_prices[unit]
        amount = price * quantity
        total += amount

        order_data['Description' + str(item)] = product_descriptions[unit]
        order_data['Quantity' + str(item)] = str(quantity)
        order_data['Price' + str(item)] = f'${price:,.2f}'
        order_data['Amount' + str(item)] = f'${amount:,.2f}'

        row += 1
        item += 1

    order_data['Total'] = f'${total:,.2f}'
    return order_data, name, email, date


order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

# Test the function
create_invoice(8515, order_data, directory)


def send_email(invoice_num, name, email, date, directory):
    gmail_user = 'danieltimmons77@gmail.com'
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},

    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()


# Test the function
send_email(8515, 'Thomas Holt', 'danieltimmons77@gmail.com', '11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames.sort()

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        print(f'Sent Invoice {invoice_num} to {name}')
        invoice_num += 1
