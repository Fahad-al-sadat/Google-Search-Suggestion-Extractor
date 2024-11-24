# **Google Search Suggestion Extractor**

**Overview**
This project is a Python script designed to extract Google search suggestions for a list of keywords stored in an Excel file. 
It automates the process of fetching search suggestions, identifies the longest and shortest suggestions, and writes the results back to the Excel file.

 **Features**
- **Automated Google Search**: Uses Selenium to fetch Google suggestions for each keyword.
- **Keyword Input**: Reads keywords from an Excel file with daily-named sheets (e.g., `Monday').
- **Result Processing**:
  - Extracts all search suggestions.
  - Identifies the longest and shortest suggestions.
  - Writes the results into the Excel file.
- **Error Handling**: Handles missing files, consent popups, and invalid or empty keywords.
  
 **Technologies Used**
- **Python**: Core programming language.
- **Selenium**: For web automation.
- **OpenPyxl**: For handling Excel files.

## **Setup and Installation**
 **1. Prerequisites**
- Python 3.x installed on your system.
- Chrome browser installed.
- ChromeDriver installed (compatible with your Chrome version). Download it from [ChromeDriver Downloads](https://sites.google.com/chromium.org/driver/).

 **2. Install Dependencies**
Install the required Python libraries:
***bash
pip install selenium openpyxl

 **3. Folder Setup**
- Place your Excel file (e.g., 'data.xlsx') in the `D:/test/' directory.
- Ensure the Excel file has:
  - **Column A**: Keywords (starting from row 2).
  - **Sheet Name**: Matches the current day (e.g., 'Monday', 'Tuesday', etc.).

## **Usage**
**1. **Run the Script**:
   - Execute the script in your terminal or IDE:
   ***bash
     python script_name.py

**2. **Google Suggestions**:
   - The script will open a Chrome browser.
   - For each keyword, it fetches Google search suggestions.

**3. **Results Output**:
   - Suggestions are saved in an updated Excel file at `D:/test/updated_data.xlsx':
     - **Column B**: Longest suggestion.
     - **Column C**: Shortest suggestion.

 **Error Handling**
- **Missing Files**: The script alerts you if the Excel file or a sheet for the current day is missing.
- **Consent Popups**: Automatically handles Google's consent dialog.
- **Empty Keywords**: Skips rows with no keywords.

 **License**
This project is open-source. Feel free to modify and use it for educational or commercial purposes.

---

 **Contributions**
Contributions are welcome! If you have ideas or improvements, feel free to fork the repository and create a pull request.

