# ⚡ Renewable-Energy-Dashboard

An advanced, Python-based analytical application built with Streamlit for simulating, optimizing, and evaluating Round-The-Clock (RTC) Hybrid Renewable Energy Systems (Solar + Wind + BESS). 

This dashboard processes high-frequency (15-minute) time-series generation data to model complex grid connectivity constraints, Battery Energy Storage System (BESS) lifecycles, Demand Fulfillment Ratios (DFR), and power exchange economics.

---

## 🚀 Project Overview

Managing hybrid renewable energy plants requires strict adherence to grid evacuation limits and contractual demand targets. This tool acts as a comprehensive computational engine that:
*   Ingests raw per-unit (PU) generation profiles for Wind and Solar assets.
*   Runs chronological storage simulation loops to track Battery State of Charge (SOC), charging, and discharging dynamically.
*   Calculates commercial metrics like Power Injected to SECI, Energy diverted to Power Exchanges, and Unsold/Curtailed power.
*   Generates highly formatted, multi-sheet corporate Excel reports for daily, monthly, and 25-year lifecycle performance.

---

## 🎛️ Unified Architecture (Three Sublayers)

This application unifies three distinct computational models into a single, seamless interface:

### 1. Model 1: PPA RTC (Seasonal Compliance)
*   Integrates dynamic Summer and Winter compliance factors.
*   Adjusts Peak Demand dynamically based on customized time-of-day and seasonal brackets.
*   Calculates Demand Fulfillment Ratio (DFR) shortfall penalties.

### 2. Model 2: PPA RTC (Standard DFR)
*   A streamlined operational model featuring strict 70% (Non-Peak) and 90% (Peak) DFR targets.
*   Allows users to instantly toggle peak-hour logic on or off, actively shifting the entire analytical matrix.

### 3. Model 3: NHPC BESS 25-Year Hybrid
*   Long-term lifecycle analytical framework predicting plant output over 25 years.
*   Accounts for distinct annual degradation rates for Solar and Wind assets.
*   Maps long-term BESS capacity fading and inverter efficiency losses.

---

## 🛠️ Key Features

*   **Dynamic Scenario Engine:** Instantly switch between P50, P75, and P90 probability generation scenarios.
*   **Customizable Time Windows:** Interactive sliders to define precise Morning and Evening Peak hour blocks.
*   **Sequential Data Processing:** Advanced handling of complex chronological loops to calculate BESS increments without look-ahead bias.
*   **Automated Corporate Reporting:** One-click generation of formatted `.xlsx` workbooks containing 15-Minute Analytics, Daily Summaries, Monthly Trends, and Executive KPI Dashboards.

---

## 💻 Tech Stack

*   **Language:** Python 3.x
*   **Frontend Framework:** Streamlit
*   **Data Manipulation:** Pandas, NumPy
*   **Excel I/O Engine:** OpenPyxl, io.BytesIO

---

## ⚙️ Installation & Usage

Follow these steps to run the dashboard locally:

**1. Clone the repository**
```bash
git clone [https://github.com/your-username/Renewable-Energy-Dashboard.git](https://github.com/your-username/Renewable-Energy-Dashboard.git)
cd Renewable-Energy-Dashboard
