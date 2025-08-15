<div align = center>

## :books: Quantum - Automated Email Excel Processor

![License](https://custom-icon-badges.demolab.com/github/license/RogerioLS/extract_email_files?logo=law&color=dark-green)
![Last commit](https://custom-icon-badges.demolab.com/github/last-commit/RogerioLS/extract_email_files?logo=history&color=dark-green)
![Code size in bytes](https://img.shields.io/github/languages/code-size/RogerioLS/extract_email_files?logo=file-code&color=dark-green)
![Repo size](https://img.shields.io/github/repo-size/RogerioLS/extract_email_files?logo=database)
![Top language](https://img.shields.io/github/languages/top/RogerioLS/extract_email_files?color=dark-green)
![Languages](https://custom-icon-badges.demolab.com/github/languages/count/RogerioLS/extract_email_files?logo=command-palette&color=red)
[![Pre-commit](https://img.shields.io/badge/pre--commit-enabled-brightgreen?logo=pre-commit&logoColor=white)](https://github.com/pre-commit/pre-commit)
</div>

### Overview

Quantum is an automated Python script that monitors an Outlook inbox for specific emails, extracts attached Excel files, performs a data quality check, and sends status notifications. It is designed to streamline the process of handling recurring data files received via email.

---

### Features

- **Email Monitoring**: Connects to an Outlook account and searches for emails with a specific subject line.
- **Attachment Extraction**: Downloads and saves `.xlsx` attachments from the target email.
- **Data Validation**: Uses `pandas` to read the Excel file and checks for a predefined number of missing values (`NaN`) in a specific column to ensure data quality.
- **Automated Notifications**:
    - Sends an **alert email** if the data quality check fails.
    - Sends a **success email** upon successful processing.
- **Robust Logging**:
    - Provides real-time, color-coded feedback in the terminal.
    - Saves detailed logs in JSON format, organized by date (`INFO` and `ERROR` logs are saved separately).
- **Environment-based Configuration**: Uses a `.env` file to manage sensitive data and paths, keeping them separate from the code.

---

### How It Works

The main workflow is orchestrated by `main.py`:

1.  **Initialization**: Loads environment variables from the `.env` file.
2.  **Email Extraction**: The `extrair_excel_email` module connects to Outlook, finds the most recent email matching the `HEADLINE_PREFIX` subject, and saves its Excel attachment to the `PASTA_RAIZ_QUANTUM` directory.
3.  **Excel Processing**: The `processar_excel_extraido` module reads the newly downloaded Excel file. It counts the number of `NaN` values in the 'Retorno' column.
4.  **Decision Making**:
    - **On Failure**: If the `NaN` count exceeds a defined limit, the process is halted. The `enviar_email_alerta` module sends an email detailing the error.
    - **On Success**: If the data is valid, the `enviar_email_sucesso` module sends a confirmation email.
5.  **Logging**: Throughout the process, the `logger_quantum` module records all actions, warnings, and errors.


### Project Structure

```bash
.
├── .env                # Environment variables (credentials, paths) - Not versioned
├── .gitignore          # Specifies files to be ignored by Git
├── documentation.txt   # This file
├── LICENSE             # Project license
├── main.py             # Main script, orchestrator of the workflow
├── README.md           # Project summary and setup guide
├── requirements.txt    # List of Python dependencies
└── source/
    ├── email/
    │   ├── extrair_excel_email.py  # Handles Outlook connection and attachment extraction
    │   ├── envia_email_alerta.py   # Sends data quality alert emails
    │   └── envia_email_sucesso.py  # Sends success confirmation emails
    ├── logger/
    │   └── logger_config.py        # Configures console and file logging
    └── manipulação_excel/
        └── manipulação_excel.py    # Handles Excel file reading and data validation
```

---

### Module Descriptions

#### `main.py`

The entry point of the application. It controls the execution flow by calling modules in the correct order:
1.  Loads environment variables from `.env`.
2.  Calls `extrair_excel_email` to get the attachment.
3.  Calls `processar_excel_extraido` to validate the data.
4.  Based on the validation result, calls either `enviar_email_alerta` or `enviar_email_sucesso`.
5.  Includes top-level error handling to catch any unexpected exceptions during the process.

#### `source/email/extrair_excel_email.py`

-   **Purpose**: To connect to Outlook, find a specific email, and download its attachment.
-   `inicializar_outlook()`: Establishes a connection with the Outlook application. It includes a retry mechanism that kills and restarts the Outlook process if the initial connection fails.
-   `extrair_excel_email()`:
    -   Searches the inbox of the specified Outlook account (`@asa.com.br`).
    -   Filters emails by the current date and a subject line containing `HEADLINE_PREFIX`.
    -   Saves the first matching `.xlsx` attachment found to the `PASTA_RAIZ_QUANTUM` directory.

#### `source/manipulacao_excel/manipulacao_excel.py`

-   **Purpose**: To read and validate the data from the extracted Excel file.
-   `ler_excel_mais_recente_da_pasta()`: Finds the most recently modified Excel file (`.xlsx` or `.xls`) in a given directory.
-   `quantidade_nan()`: Counts the number of `NaN` (Not a Number) values in the "Retorno" column of a pandas DataFrame.
-   `processar_excel_extraido()`: Orchestrates the reading and validation process. It returns the DataFrame if the `NaN` count is within the allowed limit (`limites_null`), otherwise it returns the `NaN` count.

#### `source/email/envia_email_alerta.py`

-   **Purpose**: To notify the user of a data quality issue.
-   `enviar_email_alerta()`:
    -   Constructs an HTML-formatted email.
    -   The email body includes the number of `NaN` values found versus the allowed limit.
    -   Connects to an SMTP server (Office365) using credentials from `.env` and sends the alert.

#### `source/email/envia_email_sucesso.py`

-   **Purpose**: To confirm that the process completed successfully.
-   `enviar_email_sucesso()`:
    -   Constructs a simple HTML-formatted success email.
    -   Connects to the SMTP server and sends the confirmation.

#### `source/logger/logger_config.py`

-   **Purpose**: To provide structured and informative logging.
-   `print_log()`: A function to print color-coded and timestamped messages to the console. The color can be based on the log level (`INFO`, `ERROR`, etc.) or a specified theme color.
-   `Logger` class: A silent, file-based logger.
    -   It accumulates log entries in memory (`info` and `error` lists).
    -   Using `atexit`, it automatically saves the logs to JSON files when the script terminates.
    -   Logs are organized into `YYYY/MM/` subdirectories, with filenames containing the date (e.g., `quantum_info_20250814.json`).
    -   Error logs include a full traceback for easier debugging.

---

### Configuration (`.env`)

The `.env` file is crucial for configuring the script without hardcoding sensitive information.

-   `EMAIL_USER`: The sender's email address (must have SMTP access).
-   `EMAIL_PASSWORD`: The password for the sender's email. For accounts with 2FA, an "app password" is usually required.
-   `EMAIL_DESTINATARIO`: The recipient of the alert and success emails.
-   `PASTA_RAIZ_QUANTUM`: The absolute path to the directory where the script will save the extracted Excel files.
-   `HEADLINE_PREFIX`: The text string the script looks for in the email subject to identify the correct email.
-   `PASTA_LOG`: The absolute path to the directory where JSON log files will be stored.

---

### Setup and Configuration

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd quantum
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Create the `.env` file:**
    Create a file named `.env` in the root directory and add the following variables. This file is included in `.gitignore` to prevent committing sensitive information.

    ```properties
    # Email credentials for sending notifications
    EMAIL_USER=your_email@example.com
    EMAIL_PASSWORD=your_app_password # Use an app password if 2FA is enabled
    EMAIL_DESTINATARIO=recipient_email@example.com

    # Path where the extracted Excel file will be saved
    PASTA_RAIZ_QUANTUM=W:\\path\\to\\your\\excel\\folder

    # Subject line prefix to identify the target email
    HEADLINE_PREFIX=Daily Fundos

    # Path to store JSON log files
    PASTA_LOG=W:\\path\\to\\your\\logs\\folder
    ```

---

### Usage

Run the main script from the project's root directory:

```bash
python main.py
```

The script will start, log its progress in the console, and perform the defined workflow.

---

### Dependencies

- `pandas`
- `python-dotenv`
- `colorama`
- `openpyxl`
- `pywin32`
