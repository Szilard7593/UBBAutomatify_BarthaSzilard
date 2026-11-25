# UBBAutomatify — DemoBlaze Purchase Order Automation

An end-to-end **UiPath Robotic Process Automation (RPA)** solution that automates the processing of Purchase Orders (POs) received via email and executes them on the [DemoBlaze](https://www.demoblaze.com/) e-commerce platform.

Developed by **Bartha Szilard**.

---

## Overview

This project streamlines the manual order-placement lifecycle by monitoring incoming emails for Purchase Order spreadsheets, extracting customer account details, validating requested items against live DemoBlaze catalog data, managing the shopping cart, placing orders, generating Excel status reports, and sending automated confirmation emails with visual proof of purchase.

## Key Features

- **Automated Email Ingestion:** Fetches incoming PO emails, downloads Excel attachments into organized folders (`Data/InputEmail/`), and parses account details directly from the email body.
- **Dynamic Catalog Scraping & Validation:** Scrapes available product data from DemoBlaze and validates PO line items before attempting checkout.
- **Modular Web Automation (Object Repository):** Uses UiPath's Object Repository (`.objects/`) for reliable UI interactions across navigation, authentication, product selection, and cart management.
- **Cart Verification:** Clears pre-existing cart items, adds validated order items, and cross-checks the live cart against an expected cart model prior to placing the order.
- **Reporting & Confirmation:** Captures a screenshot of the order confirmation (`OrderConfirmations/OrderConfirmation.png`), generates a structured status report (`Data/PO_StatusReport.xlsx`), and emails the summary back to the requester.

---

## End-to-End Workflow

1. **Initialization:** `Main.xaml` loads project settings from `Data/Config.xlsx` via `Data/ReadConfig.xaml`.
2. **Email Retrieval:** `Email/GetPOFromEmail.xaml` downloads PO Excel attachments and `Email/GetAccountFromEmailBody.xaml` extracts account information.
3. **Browser & Catalog Setup:** `DemoBlaze/Navigation/OpenDemoBlaze.xaml` launches the store and `DemoBlaze/Products/ScrapeProductsTable.xaml` collects available product details.
4. **PO Processing Loop (`Data/ProcessAllPOEmails.xaml`):**
   - Reads the PO file (`Data/GetPO.xaml`) and validates items (`Data/ValidatePO.xaml`).
   - Builds the expected cart state (`Data/BuildExpectedCart.xaml`).
   - Logs into the user's DemoBlaze account (`DemoBlaze/Login/PerformLogin.xaml`) and clears any existing cart items (`DemoBlaze/Cart/EmptyCart.xaml`).
   - Navigates through categories and products to add each valid item to the cart (`DemoBlaze/Cart/ProcessAllOrderItems.xaml` → `ProcessOrderItem.xaml` → `AddToCart.xaml`).
   - Validates the cart contents (`DemoBlaze/Cart/ValidateCart.xaml`) and completes checkout (`DemoBlaze/Cart/PlaceOrder.xaml`), saving a confirmation screenshot.
   - Logs out (`DemoBlaze/Login/PerformLogout.xaml`).
5. **Reporting & Notification:**
   - Generates the final report (`Data/GenerateStatusReport.xaml` → `Data/PO_StatusReport.xlsx`).
   - Builds and dispatches the confirmation email (`Email/BuildEmailBody.xaml` & `Email/SendConfirmationEmail.xaml`).
6. **Cleanup:** Closes the browser session (`DemoBlaze/Navigation/CloseDemoBlaze.xaml`).

---

## Project Structure

```text
UBBAutomatify_BarthaSzilard-main/
├── Main.xaml                               # Main entry point orchestrating the workflow
├── project.json                            # UiPath project metadata and dependencies
├── .objects/                               # UiPath Object Repository (UI descriptors & targets)
├── Data/
│   ├── Config.xlsx                         # Central configuration file (URLs, paths, email settings)
│   ├── ReadConfig.xaml                     # Loads Config.xlsx into a dictionary
│   ├── ProcessAllPOEmails.xaml             # Iterates through and processes all retrieved PO emails
│   ├── GetPO.xaml                          # Reads Purchase Order data from Excel
│   ├── ValidatePO.xaml                     # Validates PO items against scraped catalog data
│   ├── BuildExpectedCart.xaml              # Computes expected cart items and totals
│   ├── GenerateStatusReport.xaml           # Produces the PO_StatusReport.xlsx output
│   ├── PO.xlsx                             # Working/sample Purchase Order file
│   ├── PO_StatusReport.xlsx                # Generated execution status report
│   └── InputEmail/                         # Downloaded email PO attachments (Email_1, Email_2, ...)
├── DemoBlaze/
│   ├── Cart/
│   │   ├── AddToCart.xaml                  # Adds an item to the cart and handles alerts
│   │   ├── EmptyCart.xaml                  # Clears existing items from the cart
│   │   ├── PlaceOrder.xaml                 # Submits the checkout form and saves confirmation
│   │   ├── ProcessAllOrderItems.xaml       # Loops through all validated items in a PO
│   │   ├── ProcessOrderItem.xaml           # Handles navigation and addition of a single item
│   │   └── ValidateCart.xaml               # Compares actual cart contents against expected cart
│   ├── Login/
│   │   ├── PerformLogin.xaml               # Authenticates on DemoBlaze
│   │   └── PerformLogout.xaml              # Logs out of DemoBlaze
│   ├── Navigation/
│   │   ├── OpenDemoBlaze.xaml              # Opens the browser and navigates to DemoBlaze
│   │   ├── CloseDemoBlaze.xaml             # Closes the DemoBlaze browser instance
│   │   ├── NavigateToHome.xaml             # Navigates to the homepage
│   │   ├── NagivateToCategory.xaml         # Navigates to a specified product category
│   │   ├── NavigateToProduct.xaml          # Opens a specific product page
│   │   └── NavigateToCart.xaml             # Navigates to the shopping cart page
│   └── Products/
│       └── ScrapeProductsTable.xaml        # Scrapes product names and prices from DemoBlaze
├── Email/
│   ├── GetPOFromEmail.xaml                 # Fetches PO emails and saves Excel attachments
│   ├── GetAccountFromEmailBody.xaml        # Extracts account credentials/info from email body
│   ├── BuildEmailBody.xaml                 # Formats the confirmation email body
│   └── SendConfirmationEmail.xaml          # Sends the confirmation email with attachments
└── OrderConfirmations/
    └── OrderConfirmation.png               # Screenshot captured upon successful order placement
```

---

## Prerequisites

- **UiPath Studio** (Windows / Modern Design Experience recommended, matching the versions defined in `project.json`).
- **Web Browser** (e.g., Google Chrome or Microsoft Edge) with the **UiPath Browser Extension** installed and enabled.
- **Microsoft Excel** (or UiPath Excel/Workbook activities support) for reading POs and generating reports.
- **Email Account Access** (SMTP/IMAP, Outlook, or Gmail integration as configured in `Email/` workflows and `Data/Config.xlsx`).

---

## Configuration

Before running the automation, review and update `Data/Config.xlsx` with your environment-specific settings, such as:

- **DemoBlaze URL** (`https://www.demoblaze.com/`)
- **Email settings** (inbox folder, sender/recipient addresses, credentials or UiPath Orchestrator Asset names)
- **File paths** (input email directory, status report output path, and order confirmation screenshot directory)

---

## How to Run

1. Clone or extract the repository:
   ```bash
   git clone https://github.com/Szilard7593/UBBAutomatify_BarthaSzilard.git
   ```
2. Open **UiPath Studio** and select **Open a Local Project**, then choose `project.json` inside `UBBAutomatify_BarthaSzilard-main/`.
3. Wait for UiPath Studio to restore all NuGet package dependencies.
4. Verify your settings in `Data/Config.xlsx` and ensure your email client/credentials are properly configured.
5. Open `Main.xaml` and click **Run File** (or **Debug File**) in UiPath Studio.

---

## Input & Output Artifacts

| Type | Path | Description |
| :--- | :--- | :--- |
| **Input** | `Data/Config.xlsx` | Workflow configuration parameters. |
| **Input** | `Data/InputEmail/Email_*/\*.xlsx` | Incoming Purchase Order Excel files downloaded from emails. |
| **Output** | `Data/PO_StatusReport.xlsx` | Detailed execution report summarizing processed, valid, and invalid PO items. |
| **Output** | `OrderConfirmations/OrderConfirmation.png` | Screenshot of the DemoBlaze order confirmation dialog. |

---

## Author

- **Bartha Szilard** — *UBBAutomatify Project*
