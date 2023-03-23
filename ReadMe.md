<h1 align="center">
  Power Query Consolidator
</h1>


## Overview
Power Query Consolidator aims to ease the process of consolidating local, non-database files such as csv or spreadsheet before importing into Excel or Power BI.<br>
A data schema is utlised for below purposes:
 - Standardise field names across data sources
 - Control the fields selected in each source
 - Specify field types
 - Enable the creation of calculated columns via the data schema
<br>
<br>

## "Installation"
The only 2 power query functions required to be incorporated in custom Power BI or Excel projects are fn_Consolidation and fn_FieldNamesDataVersusSchemaCheck but it is recommended to download the Excel files with example data to view workings:
 - Download the latest zip file here https://github.com/charl-potgieter/PowerQueryConsolidator/releases/latest
 - Unzip in your folder of choice
 - Update Power Query parameter "Example_DataRootFolder"  to contain the the file path of the ExampleData root folder
 - Choose one of the 3 data sets using the Power Query parameter Example_SelectedDataSet
 - Consolidated output should then appear in the Example_Consolidation query
 - Helper queries Example_FieldsInDataNotSchema and Example_FieldsInSchemaNotInData list any mismatched fields betwen the schema file and the data files
<br>
<br>

## Parameters: fn_Consolidation
<br>


| Parameter  | Description |
| ------------- | ------------- |
| DataAccessFunction  | Any function that takes parameters for folder and file name and returns a table of data  |
| SourceFolder  | The path containing the data to be consolidated  |
| SchemaFilePath  | The path containing the data schema file in Excel format as per ExampleDataSchema.xlsx included in the zip file download per latest release in this repository  |
| DataSourceName | The name of the data source listed in the schema file  |
| FilterFromValue (optional) | Utilised to filter files in data source folder  |
| FilterToValue (optional) | Utilised to filter files in data source folder  |
| IsDevMode (optional)  | Boolean value. If set to true only one file of data is returned restricted to 100 rows     |
<br>
<br>

## Parameters: fn_FieldNamesDataVersusSchemaCheck
<br>


| Parameter  | Description |
| ------------- | ------------- |
| DataAccessFunction  | As per fn_Consolidation  |
| SourceFolder  | As per fn_Consolidation |
| SchemaFilePath  | As per fn_Consolidation  |
| DataSourceName | As per fn_Consolidation  |
| DirectionToCheck | Takes one of the following text inputs: "Fields in data not in schema" or "Fields in schema not in data" |
<br>
<br>



## The data schema file
<br>
The data schema file needs to be in the format of ExampleDataSchema.xlsx included in this repository with the below fields
 
 - FieldName representing the column header to be generated in the output table
 - FieldTypeAsText representing the field type (as listed in the DataTypes tab of the schema file)
 - One column for each data source representing a folder path containing source files.  The column header represents the DataSourceName listed in the parameters above.   The data in this column is either the original column name in the data source file or a calculated column formula (refer below)
<br> 
<br>



## Calculated columns
<br>
Calculated columns can be captured in the schema using <> brackets for example entering <[FirstAmount] * 2> in the schema file where [FirstAmount] is an existing column in the source data will evaluate to the following M code:  Table.AddColumn(ExistingTable, "FieldNamePerSchemaFile", each [FirstAmount]* 2, FieldTypePerSchemaFile)
<br>
See example data files and schema for workings
<br> 
<br>



## Referencing source folder and file metadata in calculated columns

The source folder and file metadata can be referenced in calculated columns as per above process using the below fields.  Note that the prefix "PQ.Consol." is added to the standard field names returned by Power Query function Folder.Files() to avoid potential name conflict wth columns in underlying data files.
- [PQ.Consol.Content]
- [PQ.Consol.Name]
- [PQ.Consol.Extension]
- [PQ.Consol.Date accessed]
- [PQ.Consol.Date modified]
- [PQ.Consol.Date created]
- [PQ.Consol.Attributes]
- [PQ.Consol.Folder Path]
