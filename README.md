# UniUni BOL Automation Agent 🚛

This is an automated tool designed to streamline the creation of **Bill of Lading (BOL)** forms on Smartsheet. It utilizes **Python** and **Selenium** to simulate browser actions, automatically filling out complex web forms based on simple shorthand commands.

> **✨ New Feature:** This tool is now powered by `uv`. You **do NOT** need to manually install Python or manage virtual environments. It works instantly on a brand new computer.

## 🚀 Key Features

* **Zero Setup Required:** No complex installation. Just run one command.
* **User-Friendly GUI:** Built with `tkinter`, offering a clean interface to input data and monitor progress.
* **High Concurrency:** Supports multi-window processing (e.g., 3-5 workers) to speed up billing.
* **Smart Address Mapping:** Automatically converts shorthand codes (e.g., `LAX`, `ORD`) into full warehouse addresses required by the form.
* **Intelligent Logic:**
    * **Auto-Carrier Selection:** Automatically selects the correct Carrier/Broker (e.g., *Han Express*, *80s Express*, *NYQZ*) based on the destination.
    * **Pallet Calculation:** Automatically determines standard pallet counts (12 vs. 26) based on the route.
* **Robust Error Handling:** Features advanced element detection to handle dynamic Smartsheet forms.

## 🛠️ Prerequisites

You only need to install **uv** (a modern Python tool manager). This is a one-time step.

* **Mac / Linux:**
    ```bash
    curl -LsSf https://astral.sh/uv/install.sh | sh
    ```
    *(If you see a warning about PATH, restart your terminal or run `source $HOME/.local/bin/env`)*

* **Windows:**
    ```powershell
    powershell -c "irm [https://astral.sh/uv/install.ps1](https://astral.sh/uv/install.ps1) | iex"
    ```

## 📦 Installation

1.  **Download the Script:**
    Save the agent code as `bol_agent_gui.py` on your computer.

2.  **No further installation needed:**
    Unlike traditional Python scripts, you do **not** need to run `pip install`. The script handles its own dependencies automatically.

## 🖥️ How to Use

1.  **Run the Application:**
    Open your terminal in the folder where you saved the file and run:
    ```bash
    uv run bol_agent_gui.py
    ```
    *(The first run will automatically download Python and necessary drivers. Subsequent runs are instant.)*

2.  **Configure the Run:**
    * **Batch Number:** Enter the current batch number (e.g., `BATCH-2025-11-25-01`).
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
    Click the **"🚀 启动多窗口并发开单"** button. The agent will launch Chrome windows and begin processing.

## ⚙️ Configuration & Logic

Since this is a single-file script, the core business logic is located directly in `bol_agent_gui.py`. You can modify the top sections of the file:

* **`ADDRESS_MAP`**: Maps shorthand keys (e.g., `ORD121`) to full address strings.
* **`get_carrier(destination_key)`**: Defines rules for selecting carriers (Han Express, 80s Express, etc.).
* **`get_pallet_count(destination_key)`**: Defines default pallet counts for short-haul vs. long-haul.

## ⚠️ Important Notes

* **Private Data:** This code contains internal business logic. Do not share publicly.
* **Network:** Ensure you have a stable internet connection, as the script interacts with a live web form.
* **Interruption:** If you need to stop the script, simply close the GUI window or press `Ctrl+C` in the terminal.

## 📝 License

For internal use only.
