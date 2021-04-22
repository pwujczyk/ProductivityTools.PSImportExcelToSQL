clear
Write-Host "pawel"

cd $PSScriptRoot
Import-Module d:\GitHub\ProductivityTools.PSExel2SQL\ProductivityTools.ImportExcelToSQL\\ProductivityTools.ImportExcelToSQL.psm1 -force
Import-Module "d:\GitHub-3.PublishedToLinkedIn\ProductivityTools.PSSQLCommands\ProductivityTools.PSSQLCommands\ProductivityTools.PSSQLCommands.psm1" -Force
cd d:\z1\
Import-ExcelToSql -SqlInstance ".\sql2019" -DatabaseName "EcoVadisDT1" -verbose
#Import-ExcelToSql -SqlInstance ".\sql2019" -DatabaseName "EcoVadisDT" -DatabaseDirectory "d:\bin\DB\" -verbose