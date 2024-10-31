<#
.SYNOPSIS
    This script replaces placeholders in a single email signature template with user-provided information (name, position, phone, email, city),
    renames the signature files based on the selected city, and copies the files and images to the Outlook signature folder.

.PARAMETER Name
    The name of the user to appear in the signature (e.g., John Doe).

.PARAMETER Position
    The job position of the user (e.g., Senior Developer).

.PARAMETER Phone
    The phone number of the user (e.g., +46 701 234 567).

.PARAMETER Email
    The email address of the user (e.g., John.Doe@nionit.com).

.PARAMETER Office
    The office for which the signature is being generated (e.g., Stockholm, Skopje, Bitola, Malmö, Sofia, Göteborg).

.EXAMPLE
    ./import-signature.ps1 -Name "John Doe" -Position "Senior Developer" -Phone "+46 701 234 567" -Email "john.doe@nionit.com" -Office "Stockholm"

    This example creates a signature for "John Doe" for the "Stockholm" office using a single template, renames the files, and copies them to the Outlook folder.

.NOTES
    The script is case sensitive. Make sure that the email and city names are provided in the correct case to match the Outlook signature folder structure.
    Outlook may not recognize the signature properly if the file or folder names have incorrect casing.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [string]$Position,

    [Parameter(Mandatory = $true)]
    [string]$Phone,

    [Parameter(Mandatory = $true)]
    [string]$Email,

    [Parameter(Mandatory = $true)]
    [ValidateSet("Stockholm", "Skopje", "Bitola", "Malmö", "Sofia", "Göteborg")]
    [string]$Office
)

# Define city-specific information
$cities = @{
    'Stockholm' = @{
        'Address' = 'Drottninggatan 95A | 113 60 Stockholm | Sweden';
        'Greeting' = 'Kind Regards | Med vänlig hälsning';
    }
    'Skopje' = @{
        'Address' = 'City Gallery, Floor 2 (11 Oktomvri 1/1) | 1000 Skopje | North Macedonia';
        'Greeting' = 'Kind Regards | Срдечен Поздрав | Med vänlig hälsning';
    }
    'Bitola' = @{
        'Address' = 'Vlatko Milenkoski 1 | 7000 Bitola | North Macedonia';
        'Greeting' = 'Kind Regards | Срдечен Поздрав | Med vänlig hälsning';
    }
    'Malmö' = @{
        'Address' = 'Norra Vallgatan 58 | 211 22 Malmö | Sweden';
        'Greeting' = 'Kind Regards | Med vänlig hälsning';
    }
    'Sofia' = @{
        'Address' = '3 Vitosha Blvd (Floor 6) | 1000 Sofia | Bulgaria';
        'Greeting' = 'Kind Regards | Срдечен Поздрав | Med vänlig hälsning';
    }
    'Göteborg' = @{
        'Address' = 'Vasagatan 58 | 411 37 Göteborg | Sweden';
        'Greeting' = 'Kind Regards | Med vänlig hälsning';
    }
}

# Get office-specific details
$Address = $cities[$Office].Address
$Greeting = $cities[$Office].Greeting

# Paths and file names
$scriptLocation = $PSScriptRoot
$TemplateName = 'NionIT_Template'
$TemplateFolderPath = Join-Path $scriptLocation $TemplateName
$outlookSignatureFolder = "$env:APPDATA\Microsoft\Signatures"
$signatureName = "_$Office ($Email)"

# Construct the destination folder name for the images
$signatureFilesFolder = ($TemplateName -replace '_Template', "_$Office ($Email)_files")

# Construct the image path for the current signature
$signatureImagePath = $signatureFilesFolder + "/logo.png"

# Function to replace placeholders in files
function Replace-Placeholders {
    param (
        [string]$filePath,
        [string]$name,
        [string]$position,
        [string]$phone,
        [string]$email,
        [string]$address,
        [string]$greeting,
        [string]$imagePath,
        [string]$outputFilePath,
        [string]$office,
        [string]$signatureName
    )
    
    # Read the template file content
    $content = Get-Content $filePath -Raw
    
    $signatureImagePath = $signatureImagePath -replace ' ', '%20'

    # Replace placeholders
    $content = $content -replace "{{Name}}", $name
    $content = $content -replace "{{Position}}", $position
    $content = $content -replace "{{Phone}}", $phone
    $content = $content -replace "{{Email}}", $email
    $content = $content -replace "{{Address}}", $address
    $content = $content -replace "{{Greeting}}", $greeting
    $content = $content -replace "{{ImagePath}}", $imagePath
    $content = $content -replace "{{Office}}", $office
    $content = $content -replace "{{SignatureName}}", $signatureName 

    # Write the updated content to the renamed output file
    Set-Content -Path $outputFilePath -Value $content
}

# Search for all files in all subfolders of the template folder
$templateFiles = Get-ChildItem -Path $TemplateFolderPath -Recurse

# Rename and replace placeholders in all template files, copying the folder structure
foreach ($file in $templateFiles) {
    # Calculate the relative path of the file to the template folder and ensure it's a single string
    $relativePath = [System.IO.Path]::GetRelativePath($TemplateFolderPath, $file.FullName)
    $relativePath = ($relativePath -replace '_Template', $signatureName)
    
    # Replace '_Template' in the file name with the office name and add email if applicable
    $newFileName = ($file.Name -replace '_Template', $signatureName)
    
    # Construct the destination path by combining the base folder, relative path, and new file name
    $outputFilePath = Join-Path $outlookSignatureFolder ($relativePath -replace [Regex]::Escape($file.Name), $newFileName)

    # Ensure the directory exists at the destination before copying the file
    $destinationDir = Split-Path -Path $outputFilePath
    if (!(Test-Path -Path $destinationDir)) {
        New-Item -Path $destinationDir -ItemType Directory | Out-Null
    }

    # If the file is an .htm or .txt file, replace placeholders
    if ($file.Extension -eq '.htm' -or $file.Extension -eq '.txt' -or $file.Extension -eq '.xml') {
        Replace-Placeholders -filePath $file.FullName -name $Name -position $Position -phone $Phone -email $Email.ToLower() -address $Address -greeting $Greeting -imagePath $signatureImagePath -outputFilePath $outputFilePath -signatureName $signatureName
    } else {
        # For non-.htm/.txt files, simply copy them without modification
        Copy-Item -Path $file.FullName -Destination $outputFilePath -Force
    }
}

Write-Host "Signature for $Email ($Office) created, renamed, and copied to Outlook signature folder successfully!"
