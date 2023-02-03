<#
.SYNOPSIS
This function downloads templates to the user's Office templates folder and updates them if newer versions are available.

.DESCRIPTION
The function takes an array of template names, the name of a folder to store the templates in, and the URL to download the templates from.
The function will download each template to the specified folder, if it doesn't exist or if the existing template is older.
The folder will be created in the user's templates folder if it doesn't already exist using TLS 1.2 for web requests.

.PARAMETER TemplateNames
An array of template names, including the extension.

.PARAMETER FolderName
The name of the folder to store the templates in.

.PARAMETER DownloadUrl
The URL to download the templates from.

.EXAMPLE
function Download-latest-Templates -TemplateNames @('template1.dotx', 'template2.xltx') -FolderName 'MyTemplates' -DownloadUrl 'https://mytemplates.com/templates'

This example downloads the templates 'template1.dotx' and 'template2.xltx' to the folder 'MyTemplates' in the user's Office templates folder.
The templates will be updated if newer versions are available at the specified download URL.
#>
function Download-latest-Templates {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$TemplateNames,
        [Parameter(Mandatory=$true)]
        [string]$FolderName,
        [Parameter(Mandatory=$true)]
        [string]$DownloadUrl
    )

    # Load the Microsoft.Office.Interop.Word, Microsoft.Office.Interop.Excel, and Microsoft.Office.Interop.PowerPoint assemblies
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel
    Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

    # Get the path to the user's templates folder, based on the extension of the template
    foreach ($templateFullName in $TemplateNames) {
        $extension = [System.IO.Path]::GetExtension($templateFullName).ToLower()
        switch ($extension) {
            '.dotx' {
                $application = 'Word'
                $enum = [Microsoft.Office.Interop.Word.WdDefaultFilePath]::wdUserTemplatesPath
            }
            '.xltx' {
                $application = 'Excel'
                $enum = [Microsoft.Office.Interop.Excel.XlDefaultFilePath]::xlUserTemplatesPath
            }
            '.potx' {
                $application = 'PowerPoint'
                $enum = [Microsoft.Office.Interop.PowerPoint.PpDefaultFilePath]::ppUserTemplatesPath
            }
        }

        $office = New-Object -ComObject "$application.Application"
        $templatesFolder = $office.Options.DefaultFilePath($enum)
        $office.Quit()

        # Create the specified folder in the templates directory, if it doesn't already exist
        $path = "$templatesFolder\$FolderName"
        if (!(Test-Path $path)) {
            New-Item -ItemType Directory -Path $path
        }

        # Force PowerShell to use TLS 1.2 for web requests
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # Download the template to the specified folder, if it doesn't exist or if the existing template is older
        $templateUri = "$DownloadUrl/$templateFullName"
        $templatePath = "$path\$templateFullName"

        if (!(Test-Path $templatePath) -or (Get-ChildItem $templatePath).CreationTime -lt (Invoke-WebRequest $templateUri).CreationTime) {
            Invoke-WebRequest -Uri $templateUri -OutFile $templatePath
        }
    }
}
<#
.SYNOPSIS
This function downloads templates to the user's Office templates folder and updates them if newer versions are available.

.DESCRIPTION
The function takes an array of template names, the name of a folder to store the templates in, and the URL to download the templates from.
The function will download each template to the specified folder, if it doesn't exist or if the existing template is older.
The folder will be created in the user's templates folder if it doesn't already exist using TLS 1.2 for web requests.

.PARAMETER TemplateNames
An array of template names, including the extension.

.PARAMETER FolderName
The name of the folder to store the templates in.

.PARAMETER DownloadUrl
The URL to download the templates from.

.EXAMPLE
function Download-Templates-FileHash -TemplateNames @('template1.dotx', 'template2.xltx') -FolderName 'MyTemplates' -DownloadUrl 'https://mytemplates.com/templates'

This example downloads the templates 'template1.dotx' and 'template2.xltx' to the folder 'MyTemplates' in the user's Office templates folder.
The templates will be updated if FileHash different from the specified download URL.
#>
function Download-Templates-FileHash {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$TemplateNames,
        [Parameter(Mandatory=$true)]
        [string]$FolderName,
        [Parameter(Mandatory=$true)]
        [string]$DownloadUrl
    )

    # Load the Microsoft.Office.Interop.Word, Microsoft.Office.Interop.Excel, and Microsoft.Office.Interop.PowerPoint assemblies
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel
    Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

    # Get the path to the user's templates folder, based on the extension of the template
    foreach ($templateFullName in $TemplateNames) {
        $extension = [System.IO.Path]::GetExtension($templateFullName).ToLower()
        switch ($extension) {
            '.dotx' {
                $application = 'Word'
                $enum = [Microsoft.Office.Interop.Word.WdDefaultFilePath]::wdUserTemplatesPath
            }
            '.xltx' {
                $application = 'Excel'
                $enum = [Microsoft.Office.Interop.Excel.XlDefaultFilePath]::xlUserTemplatesPath
            }
            '.potx' {
                $application = 'PowerPoint'
                $enum = [Microsoft.Office.Interop.PowerPoint.PpDefaultFilePath]::ppUserTemplatesPath
            }
        }

        $office = New-Object -ComObject "$application.Application"
        $templatesFolder = $office.Options.DefaultFilePath($enum)
        $office.Quit()

        # Create the specified folder in the templates directory, if it doesn't already exist
        $path = "$templatesFolder\$FolderName"
        if (!(Test-Path $path)) {
            New-Item -ItemType Directory -Path $path
        }

        # Force PowerShell to use TLS 1.2 for web requests
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # Download the template to the specified folder, if it doesn't exist or if the existing template is older
        $templateUri = "$DownloadUrl/$templateFullName"
        $templatePath = "$path\$templateFullName"

        $webResponse = Invoke-WebRequest $templateUri
        $remoteHash = (Get-FileHash $webResponse.Content).Hash
        $localHash = (Get-FileHash $templatePath -ErrorAction SilentlyContinue).Hash

        if (!(Test-Path $templatePath) -or $remoteHash -ne $localHash) {
            Invoke-WebRequest -Uri $templateUri -OutFile $templatePath
        }

    }
}

$templateNames = @("template1.dotx", "template2.docx")
$folderName = "RH"
$downloadUrl = "https://example.com"
# To Download the latest template
Download-latest-Templates -TemplateNames $templateNames -FolderName $folderName -DownloadUrl $downloadUrl
# To Download the template if different from online source
Download-Templates-FileHash -TemplateNames $templateNames -FolderName $folderName -DownloadUrl $downloadUrl
