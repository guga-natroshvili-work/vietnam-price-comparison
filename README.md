# VPC (Vietnam Price Comparison)
![VPC LOGO](/Logo.ico "VPC Logo")

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
