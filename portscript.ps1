$net = “xxx.xxx.xxx.xxx"

$range = 1..65535

Write-Host "testing"


foreach ($r in $range)
{
Write-Host $r
try
{
 
wget -Uri "$net : $r" | Write-Host
}
catch
{
Write-Host "closed"
}

}