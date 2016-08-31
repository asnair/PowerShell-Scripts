#Input file, like a .html, .css, .js, or .txt, extract everything in () that contains an http with other stuff, then output to out.txt

$regex = '\((http.*?)\)'
$file = "file.txt"
$out = "out.txt"

select-string -Path $file -Pattern $regex -AllMatches | % { $_.Matches } | % { $_.Value } > $out