function Import-ExcelToSql()
{
	[cmdletbinding()]
	param (
	[Parameter(Mandatory=$false)]
	[string]$Directory,
	
	[Parameter(Mandatory=$true)]
	[string]$SqlInstance,
	
	[Parameter(Mandatory=$true)]
	[string]$DatabaseName,
	
	[Parameter(Mandatory=$false)]
	[string]$SchemaName="xlsx",
	
	[Parameter(Mandatory=$false)]
	[Switch]$DropDatabase=$false,
	
	[Parameter(Mandatory=$false)]
	[string]$DatabaseDirectory)

	Write-Verbose "Import Excel started"
	
	
	if ($Directory -eq "")
	{
		$Directory=$((Resolve-Path .\).Path)
	}
	
	Write-Verbose "Directory $Directory"
	Write-Verbose "SqlInstance $SqlInstance"
	Write-Verbose "DatabaseName $DatabaseName"
	Write-Verbose "SchemaName $SchemaName"
	Write-Verbose "DropDatabase $DropDatabase"
	Write-Verbose "DatabaseDirectory $DatabaseDirectory"
	
	CreateStructure $Directory $SqlInstance $DatabaseName $SchemaName $DropDatabase $DatabaseDirectory
	ImportData  $Directory $SqlInstance $DatabaseName $SchemaName	
	InvokeAdditionalSQL -Directory $Directory -SqlInstance $SqlInstance -DatabaseName $DatabaseName
}

function InvokeAdditionalSql([string]$Directory,[string]$SqlInstance,[string]$DatabaseName)
{
	Ivoke-SQLScripts -SqlInstance $SqlInstance -DatabaseName $DatabaseName -Directory $Directory
}

function ImportData()
{
	[cmdletbinding()]
	param ([string]$directory,[string]$sqlInstance,[string]$databaseName,[string]$schemaName)

	Write-Verbose "$Directory $SqlInstance $DatabaseName $Schema"
	$excelFiles=GetFiles($directory)
	
	foreach($file in $excelFiles)
	{
		$excel=Import-Excel $file		
		$tableName=GetTableName $file
		
		foreach($row in $excel)
		{
			Insert $sqlInstance $databaseName $schemaName $tableName $row
		}
	}
}

function Insert()
{
	[cmdletbinding()]
	param ([string]$sqlInstance, [string]$databaseName, [string]$schemaName,[string]$tableName,$row)

	$headers=""
	$values="'"
	
	$columnNames=Get-Member -InputObject $row |where {$_.MemberType -eq "NoteProperty"} |select Name
	foreach($column in $columnNames)
	{
		$columnName= $column.Name
		$headers+="[$columnName]"+','
		$values+=$row."$columnName"
		$values+="','"
	}
	$headers=$headers.Trim(',')
	$values=$values.TrimEnd("'").TrimEnd(",")
	
	$query="INSERT INTO [$schemaName].[$tableName]($headers) VALUES ($values)"
	
	Invoke-SQLQuery -SqlInstance $SqlInstance -DatabaseName $DatabaseName -Query $Query	-Verbose:$VerbosePreference
}

function GetFiles([string]$directory)
{
	$path="$directory\*.xlsx"
	$excelFiles=Get-ChildItem -Path $path
	return $excelFiles
}

function GetTableName($file)
{
	$tableName=$file.BaseName
	return $tableName
}

function CreateStructure()
{
	[cmdletbinding()]
	param ([string]$directory,[string]$sqlInstance,[string]$databaseName,[string]$schemaName,[bool]$DropDatabase=$false,[string]$DatabaseDirectory)

	Write-Verbose "Create structure started" 
	
	if ($DropDatabase)
	{
		New-SQLDatabase -Path $DatabaseDirectory -SqlInstance $sqlInstance -DatabaseName $databaseName -Force -Verbose:$VerbosePreference
	}
	else
	{
		New-SQLDatabase -Path $DatabaseDirectory -SqlInstance $sqlInstance -DatabaseName $databaseName  -Verbose:$VerbosePreference
	}
	$excelFiles=GetFiles($directory)
	
	foreach($file in $excelFiles)
	{
		$tableName=GetTableName $file
		New-SQLTable -SqlInstance $sqlInstance -DatabaseName $databaseName -SchemaName $schemaName -TableName $tableName -Force -Verbose:$VerbosePreference
		CreateColumns $file $tableName $schemaName
	}
}

function CreateColumns()
{
	[cmdletbinding()]
	param ($file, $TableName, $schemaName)


	$excel=Import-Excel $file
	$properties = Get-Member -InputObject $excel[1] |where {$_.MemberType -eq "NoteProperty"}
	foreach($property in $properties)
	{
		$columnName = $property.Name
		$type = $($property.Definition).split(' ')[0]
		$sqlType="VARCHAR(Max)";
		switch ($type)
		{
			"double" { $sqlType="FLOAT"}
		}
		
		New-SQLColumn -SqlInstance $SqlInstance -DataBaseName $DatabaseName -SchemaName $schemaName -TableName $TableName -ColumnName $columnName -Type $sqlType -Verbose:$VerbosePreference
	}
}

Export-ModuleMember Import-ExcelToSql