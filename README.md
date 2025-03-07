# Intune-Tools

## Get-IntuneDevices  

### Overview  

**Get-IntuneDevices** is a PowerShell tool designed to quickly retrieve Intune-managed device details from an **Entra ID group**. Microsoft Intune and Entra ID do not provide a straightforward way to fetch device objects from a given group. While exporting CSV files or running PowerShell queries is possible, this script simplifies the process into **three easy steps**.  

### Features  

- Fetches **Intune-managed devices** from a specified Entra ID group (users & devices).  
- Supports authentication via **Microsoft Graph API** (using Tenant ID or Access Token).  
- Retrieves detailed device attributes, including:  
  - **Device Name, Compliance State, OS, OS Version, Enrollment Date, Last Check-In, Serial Number, Manufacturer, and more.**  
- Provides multiple output formats:  
  - **Table view** for quick visualization.  
  - **List view** for detailed inspection.  
  - **CSV export** for further analysis.  

### Installation  

Install the script directly from **PowerShell Gallery**:  

```powershell
Install-Script Get-IntuneDevices -Scope CurrentUser -Force
```

Ensure you have the required **Microsoft Graph module** installed:  

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

### Usage  
![Get-IntuneDevices](assets/Get-IntuneDevices.gif)

1. **Run the script** in PowerShell:  

   ```powershell
   Get-IntuneDevices
   ```

2. **Authenticate** using either:  
   - **Access Token** (if you already have one).  
   - **Tenant ID** (interactive login with Microsoft Graph).  

3. **Enter an Entra ID group** (paste the group link or GUID).  

4. Choose how to display or export the results:  
   - **View as Table**  
   - **View as List**  
   - **Export to CSV**  

## Example  

```
Select an option:
1. Set Tenant or AccessToken  
2. Disconnect from Tenant  
3. Check group  
4. Show devices as table  
5. Show devices as list  
6. Export devices to CSV  
7. Exit  
Enter choice (1-7):  
```
