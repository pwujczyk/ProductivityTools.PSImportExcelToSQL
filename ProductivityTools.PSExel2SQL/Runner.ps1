clear
Write-Host "pawel"

#cd $PSScriptRoot
Import-Module d:\GitHub\ProductivityTools.PSExel2SQL\ProductivityTools.PSExel2SQL\ProductivityTools.PSExel2SQL.psm1 -force
Import-Module "d:\GitHub\ProductivityTools.PSSQLCommands\ProductivityTools.PSSQLCommands\ProductivityTools.PSSQLCommands.psm1" -Force
cd d:\Employment\EcoVadis\Excels\
Import-ExcelToDB -SqlInstance ".\sql2017" -DatabaseName "EcoVadisDT" -DatabaseDirectory "d:\bin\DB\" -verbose
	 
#ImportExcel -Directory "d:" -SqlInstance ".\sql2014" -DatabaseName "Test123" -DropDatabase $false -verbose
#ImportExcel -SqlInstance ".\sql2014" -DatabaseName "Test123" -DropDatabase $false