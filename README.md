# SPP Automatic Reporting Module (SCADA Integrated)

This Python script is an automatic reporting module developed to enhance the operational efficiency of **Movicon-based Solar Power Plants (SPP)**. At a specified time (usually end-of-day), it retrieves production, reactive energy, field parameters, and alarm data from a Microsoft SQL Server database, generates a professional HTML report containing dynamic charts and tables, and automatically sends it to relevant managers via email.

-----

## ⚙️ Features

  * **Database Integration:** Retrieves daily energy, total records, and alarm logs from the Microsoft SQL Server database (recorded by Movicon SCADA).
  * **Data Analysis:** Calculates daily total energy production on a substation (TM) basis, and determines the daily average, maximum, and minimum values of critical field parameters (Current, Voltage, Irradiation, PR, Temperature).
  * **Visualization:** Creates an embedded dynamic bar chart in BASE64 format comparing the daily production of substations.
  * **HTML Reporting:** Places all analyses and detailed inverter production data into the email body as readable, mobile-friendly HTML tables.
  * **Automation:** Regularly sends the report via SMTP protocol to a specified email list.

-----

## 🛠️ Prerequisites

The script requires the following Python libraries and a connection to a SQL Server.

### Python Libraries

```bash
pip install pyodbc pandas matplotlib sqlalchemy
```

*(Note: **Microsoft ODBC Driver for SQL Server** must be installed on your system for the `pyodbc` library to work.)*

### Database Requirements

  * **Microsoft SQL Server:** Requires a server with access permission using Windows Authentication (Trusted Connection).
  * **Required Database Tables:**
      * `ENERJI` (Daily energy and active/reactive meter records)
      * `TOTAL_KAYIT` (Total field parameters, PR, Irradiation, etc.)
      * `INV_ENERJI` (Detailed inverter production data)
      * `UFUAAuditLogItem` (Alarm and event records)

-----

## 📝 Configuration

You need to adjust the `VARIABLES` section within the script according to your power plant and server settings.

### 1\. SMTP Settings

  * **`EMAILS` / `PASSWORDS`:** The email address (e.g., Gmail) to be used for sending the report and the **App Password** associated with that address. Do not use your regular email password for security reasons.

<!-- end list -->

```python
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL = EMAILS[tesis_no]       # The email address sending the report
PASSWORD = PASSWORDS[tesis_no] # App Password
```

### 2\. Plant Settings

The following variables determine the operational parameters of the script:

| Variable Name | Description | Example Value |
| :--- | :--- | :--- |
| **`tesis_no`** | The index number of the power plant being reported. | `1` (according to the tesis\_nam list) |
| **`mail_list`** | The recipient email list for the report. | `mail_adr` or `test_mail` |
| **`SERVER`** | SQL Server name/address. | `DESKTOP_B` (or server name) |
| **`DATABASE`** | The database name in SQL Server. | `DESKTOP_C_DB` (or database name) |
| **`DATE`** | Which day's data is to be reported. | `current_date` or `test_date` |
| **`filter_hour`** | The timestamp hour when the daily energy is read. | `'18:00:00'` |
| **`TM_SUM`** | Total number of substations (TM). | `6` |
| **`INV_PER_TM`** | Number of inverters per substation. | `19` |

-----

## ▶️ Running the Script

The script primarily runs by calling the `job()` function.

1.  Complete your configuration settings.
2.  Run the script file:
    ```bash
    python your_script_name.py
    ```

**Automatic Execution:** You can use Windows **Task Scheduler** or Linux cron jobs to run the script automatically at a specific time each day (e.g., 18:05). The script's `filter_hour` variable determines which timestamp's data will be fetched.

-----

## 🔑 Key Functions

  * **`get_db_conn(server_name, database_name)`:** Establishes the SQL Server connection.
  * **`fetch_energy_data(type, table)`:** Retrieves daily production and active/reactive meter data for substations.
  * **`fetch_total_kayit_data()`:** Calculates daily average/maximum/minimum values for parameters like current, voltage, PR, and temperature.
  * **`fetch_inv_energy()`:** Retrieves detailed inverter-based daily production data.
  * **`fetch_alarms()`:** Retrieves all alarm and event records from the current day.
  * **`create_bar_chart(rows, column_names)`:** Creates a bar chart from production data and converts it to Base64 format.
  * **`send_email(to_email, subject, body)`:** Sends the prepared HTML report via SMTP.
  * **`format_value(value, aftercom)`:** Ensures that the comma (`,`) is used as the decimal separator according to Turkish locale settings.