<!--Category:Powershell--> 
 <p align="right">
    <a href="https://www.powershellgallery.com/packages/ProductivityTools.PSExcel2SQL/"><img src="Images/Header/Powershell_border_40px.png" /></a>
    <a href="http://productivitytools.tech/excel2sql/"><img src="Images/Header/ProductivityTools_green_40px_2.png" /><a> 
    <a href="https://github.com/pwujczyk/ProductivityTools.PSExcel2SQL"><img src="Images/Header/Github_border_40px.png" /></a>
</p>
<p align="center">
    <a href="http://productivitytools.tech/">
        <img src="Images/Header/LogoTitle_green_500px.png" />
    </a>
</p>


# Excel to SQL

Module takes Excel files from given directory and imports it to SQL Server database.
<!--more-->

I like to have data in the database. I like to make joins an views on it. For example I have table of my team members and different different tables which represents their attendance in conferences. On those data I create views for each year.

Having data in database is great, but inserting data to it is not so simple task. So I am using excel files to it. I store bunch of excel files which contains raw data and I am importing them to database.

This module gets all excel files from given directory. For each excel it creates table and import data with the similar structure as in excel.

It exposes only one command **Import-ExcelToSql**.

Below you can find couple of the diagrams which represents steps performed.

<!--og-image-->
<img src="Images/GeneralDiagram.png" />

## CreateStructure

<img src="Images/CreateStructure.png" />

## ImportData

<img src="Images/ImportData.png" />
