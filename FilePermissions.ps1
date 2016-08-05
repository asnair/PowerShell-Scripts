
function Test-IsWritable(){
<#
    .Synopsis
        Command tests if a file is present and writable.
    .Description
        Command to test if a file is writeable. Returns true if file can be opened for write access.
    .Example
        Test-IsWritable -path $foo
		Test if file $foo is accesible for write access.
	.Example
        $bar | Test-IsWriteable
		Test if each file object in $bar is accesible for write access.
	.Parameter Path
        Psobject containing the path or object of the file to test for write access.
#>
	[CmdletBinding()]
	param([Parameter(Mandatory=$true,ValueFromPipeline=$true)][string]$path)
	
	process{
		Write-Host "Test if file $path is writeable"
		if (Test-Path -Path $path -PathType Any){ #Test-Path and Get-Item are restricted to 248 (or 264) characters in full path name, remap drives to get around this limitation
			$target = Get-Item $path -Force
			try{	

                #Because Windows permissions are so bad between NTFS and AD, it's terrible to tell real permissions by reading permissions,
                #better to just try and open the file, and write nothing to it. If opened, you have access.
				$writestream = $target.Openwrite()
				$writestream.Close() | Out-Null			
				Remove-Variable -Name writestream
				Write-Host "File is writable" -ForegroundColor Green
				Write-Output $true
				}
			catch{				
				Write-Host "File is not writable" -ForegroundColor Red
				Write-Output $false
				}
			Remove-Variable -Name target
		}
		else{
			Write-Host "File $path does not exist or is a directory" -ForegroundColor DarkRed
			Write-Output $false
		}
	}
}

write-host "WARNING: If checking deep folders (where the full path is longer than 248 characters) please " -foregroundcolor Yellow -NoNewline
Write-Host "MAP THE DRIVE " -ForegroundColor Red -NoNewline
Write-Host "in order to keep the names as short as possible" -ForegroundColor Yellow
$basefolder = Read-Host -Prompt 'What is the folder or files you want to get permissions of?'

write-host "WARNING: if permissions.csv already exists, it will be overwritten!" -foregroundcolor Yellow
Write-Host 'Export results to CSV? (y/n): ' -ForegroundColor Magenta -NoNewline
$export = Read-Host 

if ($export -like "y")
    {
        Write-Host "Name the file (ex: permissions.csv): " -ForegroundColor Magenta -NoNewline
        $FileName = Read-Host

        $Outfile = “$PSScriptRoot\$FileName”


        write-host "Will write results to $PSScriptRoot\$FileName" -ForegroundColor Green
    }

else
    {
        write-host "User did not type 'y', continuing" -ForegroundColor DarkYellow
    }



$files = get-childitem $basefolder -recurse -File -Force 

Write-Host $files
Write-Host "========================================================================" -ForegroundColor Cyan


$results = foreach($file in $files) {


            New-Object psobject -Property @{
                File = $file;
                Owner = (Get-Acl $file.FullName).Owner;
                Access = $file.FullName | Test-IsWritable
            }

}

Write-Host "Finished combo loop, exporting..." -ForegroundColor Green

$results | Export-Csv $Outfile -NoTypeInformation -Delimiter ";" #Export to CSV file

Write-Host "Converting delimited CSV to Column Excel Spreadsheet"
$outputXLSX = $PSScriptRoot + "\$Filename.xlsx"
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)
$TxtConnector = ("TEXT;" + $Outfile)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = ';'
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1
$query.Refresh()
$query.Delete()
$Workbook.SaveAs($outputXLSX,51)
$excel.Quit()

Remove-Item $Outfile #Delete CSV file, left with spreadsheet
Write-Host "See $PSScriptRoot\$Filename.xlsx for results" -ForegroundColor Green













