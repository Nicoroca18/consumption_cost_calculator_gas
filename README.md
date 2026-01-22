# Google Sheets Consumption Cost Calculator

A Google Apps Script that calculates monthly consumption costs in Google Sheets using discounted and standard pricing with a global cost threshold.

---

## Overview

This script processes usage data from a Google Spreadsheet and calculates costs based on:

- Country-based pricing
- Discounted and standard rates
- A global cost threshold
- Progressive cost accumulation

Once the threshold is exceeded, all remaining usage is billed at the standard price.

---

## Features

- Dynamic pricing by country
- Optimized reuse of prices for consecutive rows
- Partial discount application when threshold is exceeded
- Bottom-up cumulative cost calculation
- Automatic cleanup of previous results
- Built-in error handling and user alerts

---

## Spreadsheet Structure

### Active Sheet (Consumption Data)

| Column | Description |
|------|------------|
| N | SKU or identifier |
| O | Usage amount |
| P | Country |
| S4 | Global discounted cost threshold |
| W–AD | Script-generated output |

### Pricing Sheet

The pricing sheet name is dynamically read from cell **S1** of the active sheet.

| Column | Description |
|------|------------|
| A | Country |
| B | Discounted price |
| C | Standard price |

---

## Output Columns

| Column | Content |
|------|--------|
| W–Y | Copied SKU, usage, country |
| Z–AA | Applied discounted and standard prices |
| AB | Cost at discounted price |
| AC | Cost at standard price |
| AD | Cumulative total cost |

---

## Cost Calculation Logic

1. Usage is multiplied by discounted and standard prices.
2. A global threshold limits how much cost can be billed at discounted rates.
3. Rows are processed from bottom to top.
4. If a row exceeds the remaining discounted budget:
   - Part of the usage is billed at discounted price.
   - The rest is billed at standard price.
5. Once the threshold is exceeded:
   - All remaining usage is billed at standard price.
6. A cumulative total cost is stored per row.

---

## Usage

1. Open your Google Spreadsheet.
2. Go to **Extensions → Apps Script**.
3. Paste the script into the editor.
4. Save and run `calculateMonthlyConsumption`.

---

## Requirements

- Google Sheets
- Google Apps Script permissions

---

## Disclaimer

This repository contains **generic calculation logic only**.

No real pricing, customer data, or contractual information is included.
