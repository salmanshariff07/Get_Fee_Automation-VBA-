# Dividend Reissue Fee Automation Tool (VBA)

## From 10 Minutes to Under 10 Seconds — Automating a Manual Workflow with VBA


## Project Overview
This VBA Macro project automates the calculation and extraction of dividend reissue fees, a process that previously took ~10 minutes manually. With this tool, the same workflow is now completed in under 10 seconds — drastically improving efficiency and accuracy.

The tool supports call center agents by speeding up responses to fee-related queries and helping maintain optimal average handling time (AHT).

## What the code Does:

-<a href = "https://github.com/salmanshariff07/Get_Fee_Automation-VBA-/blob/main/AutomationMainCode.txt">AutomationMainCode

##Input Abnormal Data Format

A raw data dump from a webpage (unstructured format) is pasted into the first sheet.
Click "Convert Data" → cleans and reshapes the data into a structured format.
Reformatted data is automatically pasted into the Fee Data sheet.

##Filter Unclaimed Payments

On the Fee Data sheet, a button called "Filter Unclaimed" isolates records with unclaimed status.
These are transferred to a dedicated Unclaimed Payments sheet via automated code logic.

##Calculate Reissue Fees

The button "Get Fee Details" fetches individual fee details for each payment.
A total fee summary is automatically calculated and displayed at the bottom.

##Unclaimed Fee Calculation

On the Unclaimed Sheet, clicking "Get Unclaimed Fee" calculates:
Total number of unclaimed payments
Total fee amount for those payments

## Business Impact

⏳ Reduced time from 10 minutes to <10 seconds
📞 Supports call center efficiency & AHT targets
💯 Improved accuracy, no manual errors
🔄 Reusable tool for ongoing fee processing


## How to Use the below project:

- <a href = "https://github.com/salmanshariff07/Get_Fee_Automation-VBA-/blob/main/Dividend_fee_VBAProject.xlsm"> Dividend_Fee_VBAProject

Paste the webpage data into Sheet 1

  Click "Convert Data"

Go to Fee Data Sheet:
Use buttons on Fee Data Sheet:

  Filter Unclaimed
  Get Fee Details

Go to Unclaimed Sheet
  Click "Get Unclaimed Fee"

  
