//------------------------------------------------------------------------------
//        Raw data ex metadata which is added at the end
//------------------------------------------------------------------------------    


(
    DataAccessFunction as function, 
    SourceFolder as text,
    SchemaFilePath as text,
    DataSourceName as text,
    DirectionToCheck as text
)=>


let


    //------------------------------------------------------------------------------
    //        Uncomment for debugging
    //------------------------------------------------------------------------------

    /*
    DataAccessFunction = fn_ExcelFirstSheet,
    SourceFolder = "C:\etc\etc",
    SchemaFilePath = "C:\etc\DataSchema.xlsx",
    DataSourceName = "StandardTest",

    //choices for below
    //(1) "Fields in data not in schema",
    //(2) "Fields in schema not in data",
    DirectionToCheck = "Fields in data not in schema",
    */

    //------------------------------------------------------------------------------
    //        DataSchema
    //------------------------------------------------------------------------------

    SchemaSource = Excel.Workbook(File.Contents(SchemaFilePath), null, true),
    SchemaTable  = SchemaSource{[Item = "tbl_DataSchema", Kind="Table"]}[Data],
    SchemaUnpivot = Table.UnpivotOtherColumns(SchemaTable, {"FieldName", "FieldTypeAsText"}, "DataSource", "OriginalFieldName"),
    SchemaChangeType = Table.TransformColumnTypes(SchemaUnpivot,{{"FieldName", type text}, {"FieldTypeAsText", type text}, {"DataSource", type text}, {"OriginalFieldName", type text}}),
    SchemaFilterOnDataSourceName = Table.SelectRows(SchemaChangeType, each [DataSource] = DataSourceName),
    SchemaSelectNonCalcCols = Table.SelectRows(SchemaFilterOnDataSourceName, each Text.Start([OriginalFieldName], 1) <> "<"),
    SchemaOriginalFieldNames = SchemaSelectNonCalcCols[OriginalFieldName],



    //------------------------------------------------------------------------------
    //        Get folder contents and apply filters
    //------------------------------------------------------------------------------

    FolderContents = Folder.Files(SourceFolder),
    FilterOutNonData = Table.SelectRows(FolderContents, each
        Text.Upper([Name]) <> "README.TXT" and
        Text.Upper([Name]) <> "THUMBS.DB" and
        Text.Upper([Extension]) <> ".SQL" and
        Text.Start([Name], 1) <> "~"
        ),


    //------------------------------------------------------------------------------
    //        
    //------------------------------------------------------------------------------

    AddTableColumn = Table.AddColumn(FilterOutNonData, "tbl", each DataAccessFunction([Folder Path], [Name]), type table),
    AddColFieldNamesPerData = Table.AddColumn(AddTableColumn, "FieldNamesPerData", each Table.ColumnNames([tbl]), type list),

    AddColFieldNameDifferences = if DirectionToCheck = "Fields in data not in schema" then
            Table.AddColumn(AddColFieldNamesPerData, "FieldNameDifferences", each List.Difference([FieldNamesPerData], SchemaOriginalFieldNames), type list)
        else if DirectionToCheck = "Fields in schema not in data" then
            Table.AddColumn(AddColFieldNamesPerData, "FieldNameDifferences", each List.Difference(SchemaOriginalFieldNames, [FieldNamesPerData]), type list)
        else
            error "Incorrect parameter DirectionToCheck",
    
    
    SelectCols = Table.SelectColumns(AddColFieldNameDifferences,{"Name", "Folder Path", "FieldNameDifferences"}),
    ExpandedList = Table.ExpandListColumn(SelectCols, "FieldNameDifferences"),
    FilteredOutNullExceptions = Table.SelectRows(ExpandedList, each ([FieldNameDifferences] <> null)),
    ChangeType = Table.TransformColumnTypes(FilteredOutNullExceptions, {{"FieldNameDifferences", type text}})



in
    ChangeType