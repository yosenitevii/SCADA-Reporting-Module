import pyodbc
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, date
import matplotlib.pyplot as plt
import base64
from io import BytesIO
import pandas as pd
from sqlalchemy import create_engine
import locale
import math

# SMTP Settings
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAILS = [
    'gesA@gmail.com',
    'gesB@gmail.com',
    'gesC@gmail.com']
PASSWORDS = [
    '**** **** **** ****',
    '**** **** **** ****',
    '**** **** **** ****']

# VARIABLES
DESKTOP_A = r'localhost\SQLEXPRESS01'

DESKTOP_C_DB = 'table_A'
DEFAULT_DB = 'table_B'

tesis_nam = [
    'Plant A',
    'Plant B',
    'Plant C']
mail_adr = [
    'example@example.com',
    'example01@example01.com',
    'example02@example02.com']
test_mail = [
    'ulas.colak@outlook.com']

current_date = date.today()    # Mevcut tarihi al
test_date_str = '2025-01-22'
test_date_pd = pd.to_datetime(test_date_str).date()  # Tarihi datetime.date formatına çeviriyoruz
test_date = datetime.strptime(test_date_str, '%Y-%m-%d').date()

# Settings related to solar power plant!!!!!!!------------------------------------------------------------------------------------------------------------------
tesis_no = 1                            # [0] Plant A, [1] Plant B, [2] Plant C
mail_list = test_mail                   # test_mail veya mail_adr
SERVER = DESKTOP_B                      # DESKTOP_A / DESKTOP_B / DESKTOP_C / DESKTOP_D
DATABASE = DESKTOP_C_DB                 # DESKTOP_C_DB / DEFAULT_DB
DATE = current_date                     # test_date / current_date
filter_hour = '18:00:00'                # hh:mm:ss
TM_SUM = 6                              # Toplam trafo merkezi sayısı
INV_PER_TM = 19                         # Her trafo merkezindeki inverter adedi

EMAIL = EMAILS[tesis_no]
PASSWORD = PASSWORDS[tesis_no]
SANTRAL = tesis_nam[tesis_no]          
# Settings related to solar power plant!!!!!!!------------------------------------------------------------------------------------------------------------------

# Functions related to column diagram ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
tm_column_names = []
inv_column_names = []
for tm_no in range(1, TM_SUM+1):
    column_name = f'TM{tm_no}'
    tm_column_names.append(column_name)
for inv_no in range(1, INV_PER_TM+1):
    column_name = f'INV{inv_no}'
    inv_column_names.append(column_name)

# Comma for decimals
locale.setlocale(locale.LC_ALL, 'tr_TR.UTF-8')

# Fonction: Formating the value
def format_value(value, aftercom):
    # Format the value (comma for decimals)
    value = locale.format_string(f"%0.{aftercom}f", value, grouping=True)
    formatted_value = value.replace(',', 'X').replace('.', ',').replace('X', '.')
    return formatted_value

def create_bar_chart(rows, column_names):
    # Seperate the column names and values
    values = []
    for row in rows:
        for cell in row:
            # Checking for NAN
            if cell is None or (isinstance(cell, float) and math.isnan(cell)):
                values.append(0.0)
            else:
                values.append(float(cell))
    # Creating the graph
    plt.figure(figsize=(6, 3))
    plt.bar(column_names[:len(values)], values, color='skyblue')
    plt.xlabel("Üretim Noktaları")
    plt.ylabel("Üretim (kWh)")

    # Y axis limit
    max_value = max(values)
    min_value = min(values)
    plt.ylim(min_value - (min_value * 0.1), max_value + (max_value * 0.1))  # Maksimum değerin üzerine %10 ekleyerek sınırı ayarlayın

    plt.tight_layout()

    # Converting graph image to BASE64
    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)
    image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
    buffer.close()

    return f'<img src="data:image/png;base64,{image_base64}" alt="Günlük Üretim Grafiği">'

# MAIL Related functions ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
def send_email(to_email, subject, body):            # The function sending mail
    try:
        # Creating e-mail content
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = EMAIL
        msg['To'] = to_email

        # injecting HTML content
        msg.attach(MIMEText(body, 'html'))

        # Connecting Gmail server
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # TLS şifrelemesini başlat
            server.login(EMAIL, PASSWORD)  # Giriş yap
            server.sendmail(EMAIL, to_email, msg.as_string())  # send e-mail

        print(f"E-posta başarıyla gönderildi -> {to_email}")
    except Exception as e:
        print(f"E-posta gönderim hatası: {e}")

def get_db_conn(server_name, database_name, type=None):  # the function connecting to the MSSQL database
    print(f"connecting to {database_name} database")
    if type == 'engine':
        engine = create_engine(
            f"mssql+pyodbc://@{server_name}/{database_name}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
        )
        return engine
    else :
        conn = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={server_name};"
            f"DATABASE={database_name};"
            f"Trusted_Connection=yes;",
            autocommit=True
        )
        return conn

def fetch_energy_data(type='html', table='default'):
    # Setting connection
    engine = get_db_conn(SERVER, DATABASE, type='engine')
    # Tüm sütunları seç
    sql_query = "SELECT * FROM [dbo].[ENERJI];"
    # Pulling to the dataframe
    data = pd.read_sql_query(sql_query, engine)
    # converting LocalTimeCol column to datetime format
    data['LocalTimeCol'] = pd.to_datetime(data['LocalTimeCol'])

    if table == 'default':
        # Filter only data with a time stamp of 20 o'clock today
        filter_date = DATE
        filtered_data = data[
            (data['LocalTimeCol'].dt.date == filter_date) &
            (data['LocalTimeCol'].dt.time == pd.Timestamp(filter_hour).time())
        ]
        # Filter substation columns
        allowed_columns = []
        for tm_col in tm_column_names:
            column_name = f'{tm_col}_TOTAL_GUNLUK_ENERJI'
            allowed_columns.append(column_name)

        filtered_data = filtered_data[allowed_columns]
        filtered_dict = filtered_data.to_dict()
        rows = tuple(filtered_data.itertuples(index=False, name=None)) 
        
        second_column_total = 0
        # Creating HTML table
        html_table = "<table border='1' style='border-collapse: collapse; width: 100%;'>"

        for tm_col in tm_column_names:
            try:
                value = list(filtered_dict[f'{tm_col}_TOTAL_GUNLUK_ENERJI'].values())[0]
            except IndexError as e:
                value = 0
                print(f'IndexError: {e}')
            if value == None:
                    value = "0"
            else:
                second_column_total += value
                value = format_value(float(value),0)

            # Getting the column name from the custom_column_names array
            html_table += f"<td style='padding: 8px; text-align: left; width: 50%; background-color: #f2f2f2;'><b>{tm_col}</b></td>"
            html_table += f"""
                            <td style='padding: 8px; text-align: left; width: 50%;'>
                                <span style="text-align: left;">{value}</span>
                                <span style="text-align: right; margin-left: auto;">kWh</span>
                            </td>"""
            html_table += "</tr>"
        
        # Adding the Sum row
        second_column_total = format_value(float(second_column_total),0)
        html_table += "<tr style='background-color: #e0e0e0;'>"
        html_table += f"<td style='padding: 8px; text-align: left;'><b>TOPLAM</b></td>"
        html_table += f"<td style='padding: 8px; text-align: left;'>{second_column_total} kWh</td>"
        html_table += "</tr>"
        html_table += "</table>"

    if type == 'text':
        # Create plain text table
        second_column_total = 0         # Calculate second colummn sum
        text_table = "🔋Günlük Üretim Raporu\n\n"
        for row in rows:
            for i, cell in enumerate(row):
                text_table += f"{tm_column_names[i]:<30}: {cell}\n"
                # # Calculate second colummn sum
                try:
                    second_column_total += float(cell)
                except ValueError:
                    pass

        # Addin sum row
        text_table += f"\n{'TOPLAM (kWh)':<30}: {second_column_total}\n"

    if table == 'W_table':
        # Creating W_table
        W_dict = {}
        W_columns = [ 
            'WPNEG','WPPOS',
            'WQNEG','WQPOS']
        units = {
            'WPNEG': 'kWh',
            'WPPOS': 'kWh',
            'WQNEG': 'kVArh',
            'WQPOS': 'kVArh'
        }
        filter_date = DATE
        W_data = data[(data['LocalTimeCol'].dt.date == filter_date)]
        W_data = W_data[W_columns]
        for column in W_data.columns:
            # Filtering NAN values
            non_zero_values = W_data[column][W_data[column] != 0]
            
            # Calculating avg, max and min 
            max_value = non_zero_values.max()
            min_value = non_zero_values.min()  
            value = max_value - min_value
            W_dict[column] = value

        W_table = "<table border='1' style='border-collapse: collapse; width: 100%;'>"
        for col_name in W_columns:
            try:
                Wvalue = W_dict[col_name]
            except IndexError as e:
                Wvalue = 0
                print(f'IndexError: {e}')
            if Wvalue == None:
                    value = "0"
            else:
                value = format_value(float(value),0)

            W_table += f"<td style='padding: 8px; text-align: left; width: 50%; background-color: #f2f2f2;'><b>{col_name}</b></td>"
            W_table += f"""
                            <td style='padding: 8px; text-align: left; width: 50%;'>
                                <span style="text-align: left;">{Wvalue}</span>
                                <span style="text-align: right; margin-left: auto;">{units[col_name]}</span>
                            </td>"""
            W_table += "</tr>"
        W_table += "</table>"

    # return conditions
    if type == 'html' and table == 'default':
        return html_table
    elif type == 'html' and table == 'W_table':
        return W_table
    elif type == 'rows':
        return rows
    elif type == 'text':
        return text_table


def fetch_total_kayit_data():
    ortalama_dict = {}
    max_dict = {}
    min_dict = {}
    allowed_columns = [
        'IA', 'IB', 'IC', 'P', 'Q', 'S', 'VAB', 'VAN', 'VBC', 'VBN', 'VCA', 'VCN', 
        'Isinim', 'PR_DEGER', 'Hucre_Sicaklik', 'Ortam_Sicaklik', 'Ruzgar_Hiz', 'Ruzgar_Yon'
    ]
    units = {
    'IA': 'A',              # Akım (Amper)
    'IB': 'A',              # Akım (Amper)
    'IC': 'A',              # Akım (Amper)
    'P': 'kW',              # Güç (Kilowatt)
    'Q': 'kVAr',            # Reaktif Güç (Kilovolt-Amper Reaktif)
    'S': 'kVA',             # Görünen Güç (Kilovolt-Amper)
    'VAB': 'V',             # Voltaj (Volt)
    'VAN': 'V',             # Voltaj (Volt)
    'VBC': 'V',             # Voltaj (Volt)
    'VBN': 'V',             # Voltaj (Volt)
    'VCA': 'V',             # Voltaj (Volt)
    'VCN': 'V',             # Voltaj (Volt)
    'Isinim': 'W/m²',     # Isınma (Santigrat Derece)
    'PR_DEGER': '%',        # Performans Oranı (%)
    'Hucre_Sicaklik': '°C', # Hücre Sıcaklığı (Santigrat Derece)
    'Ortam_Sicaklik': '°C', # Ortam Sıcaklığı (Santigrat Derece)
    'Ruzgar_Hiz': 'm/s',    # Rüzgar Hızı (Metre/Saniye)
    'Ruzgar_Yon': '°'       # Rüzgar Yönü (Derece)
}
    engine = get_db_conn(SERVER, DATABASE, type='engine')
    sql_query = 'SELECT * FROM [dbo].[TOTAL_KAYIT]'
    data = pd.read_sql_query(sql_query, engine)

    filter_date = DATE
    filtered_data = data[(data['LocalTimeCol'].dt.date == filter_date)]
    filtered_data = filtered_data[allowed_columns]
    filtered_data['Isinim'] = filtered_data['Isinim']/10
    filtered_data['Hucre_Sicaklik'] = filtered_data['Hucre_Sicaklik']/10
    filtered_data['Ortam_Sicaklik'] = filtered_data['Ortam_Sicaklik']/10
    filtered_data['PR_DEGER'] = filtered_data['PR_DEGER'].where(filtered_data['PR_DEGER'] > 0)

    for column in filtered_data.columns:
        # Filtering NAN values
        non_zero_values = filtered_data[column][filtered_data[column] != 0]
        
        # Calculating avg, max and min 
        ortalama = non_zero_values.mean() if not non_zero_values.empty else 0
        max_value = non_zero_values.max() if not non_zero_values.empty else 0
        min_value = non_zero_values.min() if not non_zero_values.empty else 0
        
        # Formating and saving to the dictionaries
        ortalama_dict[column] = format_value(ortalama,1)
        max_dict[column] = format_value(max_value,1)
        min_dict[column] = format_value(min_value,1)
    
    # Creating HTML table
    html_table = """
    <table border="1" style="border-collapse: collapse; width: 100%;">
    <thead>
        <tr style="background-color: #f2f2f2; text-align: center;">
            <th style="text-align: left;">Sütun Adı</th>
            <th>Ortalama</th>
            <th>Maksimum</th>
            <th>Minimum</th>
        </tr>
    </thead>
    <tbody>
    """
    # Adding data from dictionaries to table via loop
    for key in ortalama_dict.keys():
        unit = units.get(key, '-')
        html_table += f"""
            <tr>
                <td style="font-weight: bold; background-color: #e0e0e0; text-align: left;">{key}</td>
                <td style="padding: 5px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="text-align: left;">{ortalama_dict[key]}</span>
                        <span style="text-align: right; margin-left: auto;">{unit}</span>
                    </div>
                </td>
                <td style="padding: 5px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="text-align: left;">{max_dict[key]}</span>
                        <span style="text-align: right; margin-left: auto;">{unit}</span>
                    </div>
                </td>
                <td style="padding: 5px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="text-align: left;">{min_dict[key]}</span>
                        <span style="text-align: right; margin-left: auto;">{unit}</span>
                    </div>
                </td>
            </tr>
        """

    # Closing the table
    html_table += """
    </tbody>
    </table>
    """

    return html_table

def fetch_inv_energy(type='default'):
    # Generating the list of columns to be filtered
    allowed_columns = []
    # Substations numbers
    for tm_no in range(1, TM_SUM+1):
        # Inverter numbers for per substation
        for inv_no in range(1, INV_PER_TM+1):
            # Creating column names formatted as TM{tm_no}_INV{inv_no}_GUNLUK_ENERJI
            column_name = f'TM{tm_no}_INV{inv_no}_GUNLUK_ENERJI'
            allowed_columns.append(column_name)

    engine = get_db_conn(SERVER, DATABASE, type='engine')
    sql_query = 'SELECT * FROM INV_ENERJI'
    data = pd.read_sql_query(sql_query, engine)

    filter_date = DATE
    filtered_data = data[
        (data['LocalTimeCol'].dt.date == filter_date) &
        (data['LocalTimeCol'].dt.time == pd.Timestamp(filter_hour).time())]
    try: 
        filtered_data = filtered_data[allowed_columns]
    except KeyError as e:
        print(f'Keyerror hatası: {e}')

        # Changing NAN values with zero
    filtered_data = filtered_data.fillna(0)

    filtered_dict = filtered_data.to_dict()
    rows = []

    # HTML Table...
    html_table = """
    <div style="border: 1px solid #ccc;">
        <table border="1" style="border-collapse: collapse; text-align: left; width: 100%;">
            <thead>
                <tr style="background-color: #f2f2f2; text-align: center;">
                    <th></th>
    """
    # First row
    for tm_col in tm_column_names:
        html_table += f"""
            <th style="font-weight: bold; background-color: #e0e0e0; text-align: center;">{tm_col.split('_')[0]}</th>
            """

    html_table += "</tr></thead><tbody>"
    

    # adding rows per inverter
    inv_sum = 0
    for inv_no in range(1, INV_PER_TM+1): 
        html_table += "<tr>"
        html_table += f"""<td style="font-weight: bold; background-color: #e0e0e0; text-align: left;">INV {inv_no}</td>"""  # first row for inverter number

        # Adding energy amount for per substation
        for tm_col in tm_column_names:
            try:
                value = list(filtered_dict[f'{tm_col}_INV{inv_no}_GUNLUK_ENERJI'].values())[0]
            except KeyError as e:
                print(f'Keyerror hatası: {e}')
                value = 0
            except IndexError as e:
                print(f'IndexError hatası: {e}')
                value = 0

            if value is not None and not math.isnan(float(value)):
                value = float(value) / 1000
                formatted_value = format_value(value,1)
            else:
                value = 0
                formatted_value= 0
            
            inv_sum = inv_sum + value
            rows.append(value)
            html_table += f"""<td style="text-align: center;>
                            <span style="float: left;">{formatted_value}</span>
                            <span style="text-align: right; margin-left: auto;"> kWh</span>
                            </td>"""

        html_table += "</tr>"
    
    if  SANTRAL == tesis_nam[2]:
        inv_sum = format_value(float(inv_sum),0)
        print(inv_sum)
        html_table += "<tr style='background-color: #e0e0e0;'>"
        html_table += f"<td style='padding: 8px; text-align: left;'><b>TOPLAM</b></td>"
        html_table += f"<td style='padding: 8px; text-align: center;'>{inv_sum} kWh</td>"
        html_table += "</tr>"
    
    # Closing the table
    html_table += "</tbody></table>"
    
    if type == 'default':
        return html_table
    if type == 'dict':
        return filtered_dict
    if type == 'rows':
        return rows

def fetch_alarms():
    engine = get_db_conn(SERVER, DATABASE, type='engine')
    sql_query = 'SELECT * FROM [dbo].[UFUAAuditLogItem]'
    data = pd.read_sql_query(sql_query, engine)

    filter_date = DATE
    filtered_data = data[(data['EventDateTime'].dt.date == filter_date)]
    
    filtered_df = filtered_data[filtered_data['EventType'] == 'i=10751']

    # HTML Table
    html_table = """
    <table border="1" style="border-collapse: collapse; text-align: center; width: 100%;">
        <thead>
            <tr style="background-color: #e0e0e0;">
                <th>TARİH</th>
                <th>ALARM</th>
                <th>MESAJ</th>
            </tr>
        </thead>
        <tbody>
    """

    # Creating table from dataframe
    for _, row in filtered_df.iterrows():
        html_table += "<tr>"
        html_table += f"<td>{row['EventDateTime']}</td>"
        html_table += f"<td>{row['SourceName']}</td>"
        html_table += f"<td>{row['EventMessage']}</td>"
        html_table += "</tr>"

    # Closing table
    html_table += """
        </tbody>
    </table>
    """
    return html_table

# Main function
def job():
    W_table = fetch_energy_data(type='html', table='W_table')
    average_table = fetch_total_kayit_data()
    inv_table = fetch_inv_energy()
    alarm_table = fetch_alarms()

    # Creating e-mail body
    email_body = f"""
<html>
    <body style="font-family: Arial, sans-serif; margin: 20px;">
        <!-- Tarih Bilgisi -->
        <div style="text-align: left; font-size: 12px; margin-bottom: 10px;">
            <strong>Tarih:</strong> {DATE}
        </div>
        
        <!-- Başlık -->
        <h2 style="text-align: center;">{SANTRAL} GES Günlük Rapor</h2>
"""        
    email_body += f"""
        <!-- Aktif_Tablosu -->
        <div style="width: 100%; max-width: 600px; margin: 0 auto; margin-bottom: 20px; text-align: center;">
            <h3>Aktif-Reaktif Enerji Tablosu</h3>
            <div style="overflow-x: auto;">
                {W_table}
            </div>
        </div>"""
    
    if TM_SUM > 1:
        tm_rows = fetch_energy_data(type='rows')
        energy_table = fetch_energy_data(type='html')

        email_body += f"""
            <!-- Enerji_Tablosu -->
            <div style="width: 100%; max-width: 600px; margin: 0 auto; margin-bottom: 20px; text-align: center;">
                <h3>Günlük Enerji Tablosu</h3>
                <div style="overflow-x: auto;">
                    {energy_table}
                </div>
            </div>"""

        # Creating graph
        tm_chart = create_bar_chart(tm_rows, tm_column_names)
        email_body += f"""
        <!-- Enerji_Grafik -->
        <div style="width: 100%; max-width: 600px; margin: 0 auto; text-align: center;">
            <h3>Günlük Enerji Grafiği</h3>
                {tm_chart}
        </div>"""
    
    email_body += f"""
        <!-- Ortalama_Tablosu -->
        <div style="width: 100%; max-width: 600px; margin: 0 auto; margin-bottom: 20px; text-align: center;">
            <h3>Günlük Özet Tablosu</h3>
            <div style="overflow-x: auto;">
                {average_table}
            </div>
        </div>

        <!-- İnverter_Tablosu -->
        <div style="width: 100%; max-width: 600px; margin: 0 auto; margin-bottom: 20px; text-align: center;">
            <h3>Günlük İnverter Üretimi</h3>
            <div style="overflow-x: auto;">
                {inv_table}
            </div>
        </div>

        <!-- Alarm_Tablosu -->
        <div style="width: 100%; max-width: 600px; margin: 0 auto; margin-bottom: 20px; text-align: center;">
            <h3>Bugün Gelen Alarmlar</h3>
            <div style="overflow-x: auto;">
                {alarm_table}
            </div>
        </div>
    </body>
</html>
    """
    for mail_adres in mail_list:
        send_email(mail_adres, f'{SANTRAL} GES Günlük Rapor', email_body)

job() # Calling main function
