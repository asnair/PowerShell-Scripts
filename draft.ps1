#Find words/patterns in files and return which files have that pattern

#The pattern to find, standard regex syntax
$regex = '^draft \= true'

#get the files in the current directory ending with .md
$files = Get-Item *.md

$t = foreach ($file in $files)
{
    $lines = Get-Content $file #get the content in the current file
    $u = foreach ($line in $lines) 
    {

        if ($match = select-string -InputObject $line -Pattern $regex -AllMatches | % { $_.Matches } | % { $_.Value }) #if the line matches the pattern
        {
            Write-Host $file -ForegroundColor Green #outputs the full path + file name
            Write-Host $match -ForegroundColor Cyan #the pattern found (will output $regex)
	    }
    }
}


