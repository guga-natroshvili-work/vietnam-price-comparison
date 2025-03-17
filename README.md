![VPC LOGO](/Logo.ico "VPC Logo")
# VPC (Vietnam Price Comparison)
A Python-based tool for generating price comparison reports between Rustar Online Tourism (DrMsql) and its competitors, Amega Travel, Crystal Tourism, Victoria Tour, and Rustar DMC Vietnam. The app matches hotel and category data, processes pricing information, and generates an Excel report highlighting price differences.

## Functionality  
VPC generates an Excel file containing price comparisons between **Rustar Online Tourism (DrMsql)** and its competitors: **Amega Travel, Crystal Tourism, Victoria Tour, and Rustar DMC Vietnam**.  

### How It Works  
1. **Data Collection:** Competitor prices are gathered from **Onur**, while Rustar prices come from **DrMsql**.  
2. **Hotel & Category Matching:** The system uses **RapidFuzz** to match hotels and categories between datasets.  
3. **Stored Procedure Processing:**  
   - Inside **`reporting.dbo.HotelPrices_compare`**, hotel and category IDs from Onur are mapped to DrMsql.  
   - The stored procedure takes **two parameters**: `date` and `number of adults`.  
   - It retrieves all relevant price data, matches them, and calculates the **price differences**.  
4. **Price Highlighting:**  
   - If **our price is higher** than a competitor's, it is **highlighted in red**.  
   - If **our price is lower**, it is **highlighted in green**.  
5. **Excel Report Generation:** The main executable (`main.exe`) fetches the processed data from the stored procedure, applies Excel formatting, and saves the final report.  



## Installation  

### **Requirements**  
- **Windows OS**  
- **ODBC Driver 17 & 18 for SQL Server** ([Download Here](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server))  

### **Installation Steps**  
- Install **ODBC Driver 17 and 18 for SQL Server**
- **Clone** the repository or Extract the `VNM.rar`
- **Create a shortcut** of `main.exe` on your desktop.



## How to Use
1. **Input Date**  
2. **Input Number of Adults**  
3. **Select Directory** to Save Excel Files  
4. **Press "Download"**  
   - The file will be saved in this format:  
     ```
     VNM dd-mm-yy, [adult].xlsx
     ```



## App Preview
![Preview](/Preview.png "App Preview Image")

---

## **Troubleshooting**
- If the app does **not run**, make sure **ODBC drivers** are installed.  
- If SQL connection fails, check **firewall settings** or **database credentials**.


---

## **Development Files**
![Python](https://img.shields.io/badge/Python-3.13-green.svg)
![Pandas](https://img.shields.io/badge/Pandas-2.2.3-white.svg)
![RapidFuzz](https://img.shields.io/badge/RapidFuzz-3.12.2-blue.svg)
![OpenPyXl](https://img.shields.io/badge/OpenPyXl-3.1.5-darkgreen.svg)


These files were used during the initial stages of development and are still relevant for ongoing processes, such as matching new hotels and categories. While some functionalities have been integrated into the SQL stored procedure, the matching logic remains a crucial part of the workflow for adding new data.

### `data_io.py`
Handles reading and exporting data between Excel files. It includes functions to:
- **Import data** from Excel sheets into pandas DataFrames (`enter_data`).
- **Export the processed data** back to an Excel file (`export_data`).

### `match_hotels.py`
Uses **RapidFuzz** for fuzzy matching to match hotel names between datasets. The script processes hotel names from the Rustar dataset (`ro_df`) and compares them against competitor hotels from the Competators dataset (`summary`). It merges the matched data and exports it to an Excel file.

### `match_categories.py`
Fuzzy matches room categories between the Rustar (`ro_df`) and Competators(`samo_df`) datasets using **RapidFuzz**. It processes room categories, removes text inside parentheses, and matches the stripped categories. It also includes an additional function to match the text inside parentheses for more granular comparison. The results are saved to an Excel file.

### Required Python Packages for Matching
- **Pandas**: 2.2.3 ([More Info](https://pypi.org/project/pandas/))
- **Rapidfuzz**: 3.12.2 ([More Info](https://pypi.org/project/RapidFuzz/))
- **openpyxl**: 3.1.5 ([More Info](https://pypi.org/project/openpyxl/))

```bash
pip install pandas==2.2.3, rapidfuzz==3.12.2, openpyxl==3.1.5
```

These files remain essential for matching new hotels and categories and uploading the results to SQL for further processing.
