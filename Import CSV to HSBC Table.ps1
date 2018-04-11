## Runs an existing SSIS package on the SQL Server, with the addition of selecting the Data Source file to import.

Write-Host "Imports new CSV files into the Bank Statments\HSBC table."  -ForegroundColor Cyan
Write-Host "Make sure the CSV file is correctly sanitised first!" -ForegroundColor Cyan
Write-Host "The script must be ran from the LocalHost!" -ForegroundColor Yellow
Write-Host "Remember to remove any quotations from the CSV, and any 'Amounts' that contain commas" -ForegroundColor Magenta

## Prompt to select the CSV file to import, else Exit script.
$Input = Read-Host "Select CSV file to import? Yes/No"
if($Input -match 'Yes'){
Add-Type -AssemblyName System.Windows.Forms
    
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
     Multiselect = $false ## Multiple files can be chosen
	      Filter = 'CSV (*.csv)|*.csv' # Specified file types
    }
 
[void]$FileBrowser.ShowDialog()

$path = $FileBrowser.FileNames;
}
else
{Exit}

## Variables
$SSISNamespace = "Microsoft.SqlServer.Management.IntegrationServices"
$ssisParameter = "CM.SourceConnectionFlatFile.ConnectionString"
$TargetServerName = "localhost\TheDen"
$TargetFolderName = "Finance Imports"
$ProjectName = "HSBC CSV Import"
$PackageName = "HSBC CSV Import.dtsx"
$ssisParameterValue = "$path" ## Defined by above Windows Form prompt.

## Load the IntegrationServices assembly
$loadStatus = [System.Reflection.Assembly]::Load("Microsoft.SQLServer.Management.IntegrationServices, "+
    "Version=14.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91, processorArchitecture=MSIL")

## Create a connection to the server
$sqlConnectionString = `
    "Data Source=" + $TargetServerName + ";Initial Catalog=Testing;Integrated Security=SSPI;"
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection $sqlConnectionString

## Create the Integration Services object
$integrationServices = New-Object $SSISNamespace".IntegrationServices" $sqlConnection

## Get the Integration Services catalog
$catalog = $integrationServices.Catalogs["SSISDB"]

## Get the folder
$folder = $catalog.Folders[$TargetFolderName]

## Get the project
$project = $folder.Projects[$ProjectName]

## Get the package
$package = $project.Packages[$PackageName]

## Modify the Source Data File to import
$Package.Parameters[$ssisParameter].Set([Microsoft.SqlServer.Management.IntegrationServices.ParameterInfo+ParameterValueType]::Literal,$ssisParameterValue);
$Package.Alter()

## Run the package
Write-Host "Running " $PackageName "..."

$result = $package.Execute("false", $null)

Write-Host "Done."



