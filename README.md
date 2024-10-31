# 📄 **Email Signature Import Script Documentation**

## 🚀 **Overview**

The **Email Signature Import Script** is a PowerShell tool that automates the creation and configuration of email signatures based on user-provided details. It populates placeholders in a template signature, renames files according to the user’s office and email, and saves the customized signature to the Outlook signatures folder.

## ⚠️ **Disclaimer**

**Important:**  
_Ensure you are using this script within an approved and secure environment. Double-check permissions to avoid unexpected modifications to your Outlook signature folder._

## 🌟 **Features**

- **Customized Signatures**: Generate a personalized signature with details like `Name`, `Position`, `Phone`, `Email`, and `Office`.
- **Multiple Office Support**: Automatically formats address and greeting based on the specified office location.
- **File and Folder Renaming**: Saves files and folders with unique names based on the user’s office and email.
- **Placeholders Replacement**: Dynamically replaces placeholders in `.htm`, `.txt`, and `.xml` files with provided user details.

## 🛠 **Installation**

1. **Ensure PowerShell is Installed**

   Verify PowerShell installation:

   ```powershell
   $PSVersionTable.PSVersion
   ```

   This script is compatible with PowerShell 5.1 and above.

2. **Prepare Template Folder**

   The script expects a folder named `NionIT_Template` in the same directory, containing:
   - The template files for signature (`.htm`, `.txt`, `.xml`).
   - An image folder (`NionIT_Template_files`) with any necessary branding assets.

## 📖 **Usage**

Run the script from PowerShell, specifying the required parameters.

### **Syntax**

```powershell
./Generate-SignatureForOutlook.ps1 -Name "<Full Name>" -Position "<Position>" -Phone "<Phone Number>" -Email "<Email Address>" -Office "<Office Name>"
```

### **Positional Arguments**

| **Parameter** | **Description**                                             | **Example**               |
|---------------|-------------------------------------------------------------|---------------------------|
| `Name`        | Full name of the user                                       | `John Doe`                |
| `Position`    | Job title or position                                       | `Senior Developer`        |
| `Phone`       | Phone number                                                | `+46 701 234 567`         |
| `Email`       | Email address for the signature (case-sensitive)            | `John.Doe@nionit.com`     |
| `Office`      | Office location (case-sensitive)                            | `Stockholm`               |

### **Examples**

1. **Create Signature for Stockholm Office**

   ```powershell
   ./Generate-SignatureForOutlook.ps1 -Name "John Doe" -Position "Senior Developer" -Phone "+46 701 234 567" -Email "john.doe@nionit.com" -Office "Stockholm"
   ```

   - **Behavior**: Populates the signature with provided details, using the Stockholm address and greeting.
   - **Output Folder**: `%APPDATA%\Microsoft\Signatures`

2. **Create Signature with a Custom Folder Structure**

   ```powershell
   ./Generate-SignatureForOutlook.ps1 -Name "Jane Doe" -Position "Marketing Manager" -Phone "+46 701 456 789" -Email "jane.doe@nionit.com" -Office "Göteborg"
   ```

   - **Behavior**: Uses Göteborg-specific address and greeting, saving the signature files in the Outlook signatures folder.
   - **Output Folder**: `%APPDATA%\Microsoft\Signatures`

## 🧩 **Template Structure**

The `NionIT_Template` folder should be structured as follows:

```
NionIT_Template/
├── NionIT_Template.htm
├── NionIT_Template.txt
├── NionIT_Template.xml
└── NionIT_Template_files/
    └── colorschememapping.xml
    └── filelist.xml
    └── image001.png
    └── themedata.thmx
```

### Placeholder Variables in Templates

The following placeholders in the template files will be replaced by user-specific values:

- `{{Address}}`: Address based on specified office
- `{{Email}}`: User's email (converted to lowercase)
- `{{Greeting}}`: Custom greeting based on office
- `{{Name}}`: User's name
- `{{Office}}`: Office location name
- `{{Phone}}`: User's phone number
- `{{Position}}`: User's job title

## 🛠 **Troubleshooting**

- **Case Sensitivity**: Ensure the `Office` and `Email` parameters are provided in the correct case. Mismatches may cause incorrect folder creation.
- **Folder Structure**: Verify that all template files are in the `NionIT_Template` folder and that the folder structure matches the script requirements.
- **File Permissions**: Running as an administrator may be necessary if permission errors occur when accessing the Outlook signatures folder.
