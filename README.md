# UniUni BOL Automation Agent 🚛

This is an automated tool designed to streamline the creation of **Bill of Lading (BOL)** forms on Smartsheet. It utilizes **Python** and **Selenium** to simulate browser actions, automatically filling out complex web forms based on simple shorthand commands.

## 🚀 Key Features

* **User-Friendly GUI:** Built with `tkinter`, offering a clean interface to input data and monitor progress.
* **Batch Processing:** Supports bulk creation of BOLs. You can paste a list of orders, and the agent will process them sequentially.
* **Smart Address Mapping:** Automatically converts shorthand codes (e.g., `LAX`, `ORD`) into full warehouse addresses required by the form.
* **Intelligent Logic:**
    * **Auto-Carrier Selection:** Automatically selects the correct Carrier/Broker (e.g., *Han Express*, *80s Express*, *NYQZ*) based on the destination.
    * **Pallet Calculation:** Automatically determines standard pallet counts (12 vs. 26) based on the route.
* **Robust Error Handling:** Features advanced element detection (using `aria-label` and label text matching) to handle dynamic Smartsheet forms, including specific fixes for date pickers and multi-line text areas.

## 🛠️ Prerequisites

* **Operating System:** macOS or Windows.
* **Browser:** Google Chrome (latest version recommended).
* **Python:** Python 3.8 or higher.

## 📦 Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/YOUR_USERNAME/uniuni-bol-automation.git](https://github.com/YOUR_USERNAME/uniuni-bol-automation.git)
    cd uniuni-bol-automation
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    *(Note: The main dependencies are `selenium` and `webdriver-manager`.)*

## 🖥️ How to Use

1.  **Run the Application:**
    ```bash
    python bol_agent_final.py
    ```

2.  **Configure the Run:**
    * **Batch Number:** Enter the current batch number (e.g., `BATCH-2025-11-25-01`). Defaults to today's date.
    * **Email:** Verify your email address for the BOL receipt.

3.  **Input Orders:**
    Paste your orders into the text area using the following format:
    `ORIGIN-DESTINATION *QUANTITY`

    **Example:**
    ```text
    EWR936-ORD *1
    EWR936-LAX *2
    EWR936-JFK *1
    ```
    * *This will create 1 BOL for ORD, 2 BOLs for LAX, and 1 BOL for JFK.*

4.  **Start:**
    Click the **"开始自动开单" (Start Automation)** button. The agent will launch a Chrome browser and begin filling out the forms. You can watch the progress in the log window.

## ⚙️ Configuration & Logic

The core business logic is located in `bol_agent_final.py`. You can modify the following dictionaries/functions to update warehouse addresses or rules:

* **`ADDRESS_MAP`**: Maps shorthand keys (e.g., `ORD121`) to full address strings.
* **`get_carrier(destination_key)`**: Defines rules for selecting carriers (Han Express, 80s Express, etc.).
* **`get_pallet_count(destination_key)`**: Defines default pallet counts for short-haul vs. long-haul.

## ⚠️ Important Notes

* **Private Data:** This repository is set to **Private** as it contains internal business links and email addresses. Do not share this code publicly.
* **Network:** Ensure you have a stable internet connection, as the script interacts with a live web form.
* **Interruption:** If you need to stop the script, simply close the GUI window or press `Ctrl+C` in the terminal.

## 📝 License

For internal use only.
