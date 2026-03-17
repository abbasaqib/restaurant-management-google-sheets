# Culinary Automation System

A Google Apps Script web application for automating culinary service documentation including quotes, contracts, menu schedules, labels, and AI-generated shopping lists and reheating instructions.

## Features

- **Quote Generation**: Create professional catering proposals with menu details and pricing
- **Contract Creation**: Generate comprehensive catering service agreements
- **Menu Labels**: Automatically create dish and ingredient labels (splits into multiple documents if >30 labels)
- **Menu Schedules**: Organize dishes by date and meal type
- **AI Integration**: Uses Google's Gemini API to generate:
  - Shopping lists from menus
  - Reheating instructions for clients
- **Email Automation**: Send documents to clients and team members with appropriate access
- **Google Sheets Integration**: Store all submissions and track document URLs
- **Folder Organization**: Automatically organize files in Drive folders by client/event

## Prerequisites

- Google account with access to:
  - Google Drive
  - Google Docs
  - Google Sheets
  - Google Apps Script
- Gemini API key (for AI features)

## Setup Instructions

### 1. Create a New Google Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click "New Project"
3. Delete any default code
4. Copy the entire `Code.gs` file from this repository
5. Paste into the editor

### 2. Configure the Script

Update the `CONFIG` object at the top of the file with your information:

```javascript
const CONFIG = {
  COMPANY_NAME: 'Your Company Name',
  CHEF_NAME: 'Your Chef Name',
  EMAIL: 'info@yourcompany.com',
  PHONE: '(123) 456-7890',
  WEBSITE: 'www.yourcompany.com',
  TAX_RATE: 0.0895, // Your tax rate (e.g., 8.95%)
  FOLDER_ID: 'YOUR_DRIVE_FOLDER_ID', // Main folder for all documents
  LOGO_FILE_ID: 'YOUR_LOGO_FILE_ID', // Google Drive file ID for your logo
  SIGNATURE_IMAGE_ID: 'YOUR_SIGNATURE_ID', // Google Drive file ID for signature
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID', // Will be created in step 4
  GEMINI_API_KEY: 'YOUR_GEMINI_API_KEY', // Get from Google AI Studio
  EDITORS: [
    'editor1@yourcompany.com',
    'editor2@yourcompany.com'
  ],
  COMPANY_ADDRESS: 'Your Company Address',
  CHEF_TITLE: 'Executive Chef & Owner',
  SIGNATURE_NAME: 'Your Name',
  DEFAULT_DEPOSIT_PERCENT: 50,
  DEFAULT_CANCELLATION_DAYS: 5,
  MENU_TEXT: ""
};
