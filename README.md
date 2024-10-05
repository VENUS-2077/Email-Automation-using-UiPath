# Email Automation using UiPath

This project automates the process of retrieving unread emails from Outlook, preprocessing their content, performing sentiment analysis using a Python script, and sending the analyzed results back via email. The workflow is built using UiPath.

## Overview

This UiPath project is designed to automate email processing tasks, specifically:

1. Retrieve unread emails from a specific folder in Outlook.
2. Extract relevant details such as sender name, email address, subject, date and time, and body content.
3. Preprocess the email text by cleaning and formatting it for sentiment analysis.
4. Run sentiment analysis using an external Python script.
5. Save the processed data and analysis results to a text file.
6. Send the analysis results back to the user via email.

## Workflow Structure

The main workflow is composed of several sequences and activities, structured as follows:

1. **Main Sequence**
2. **Step 1: Retrieving Unread Outlook Mails**
3. **Step 2: For Each current Mail**
   - Retrieve Sender Name
   - Retrieve Mail ID
   - Retrieve Subject
   - Retrieve Date and Time
   - Retrieve Email Body
4. **Step 3: Preprocess Text**
   - Invoke Python Method
5. **Step 4: Write Data to Text File**
6. **Step 5: Sentiment Analysis**
   - Invoke Python Method
7. **Retrieve Report Results and Send Email Back to the User**
   - Send Outlook Mail Message

## Requirements

- UiPath Studio
- Microsoft Outlook(old version,not one from MS Store)
- Python (with necessary libraries for sentiment analysis)

## Usage

- Run the Main.xaml file in UiPath Studio.
- The workflow will automatically retrieve unread emails from Outlook, preprocess their content, perform sentiment analysis, and send the results back via email.
