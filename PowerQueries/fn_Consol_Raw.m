// Raw function excluding the documentation metadata

/*
(
    DataAccessFunction as function, 
    SourceFolder as text, 
    DataSchema as table,
    optional FilterFromValue as any, 
    optional FilterToValue as any, 
    optional IsDevMode as logical
)=>    
*/
let


    //------------------------------------------------------------------------------
    //        Uncomment for debugging
    //------------------------------------------------------------------------------
    
    DataAccessFunction = fn_ExcelFirstSheet,
    SourceFolder = "D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\PowerQueryConsolidator\Testing\Test_Consolidation\Test_Consolidation_Years\",
    DataSchema = Table.SelectRows(ConsolidationDataSchema, each [DataSource] = "StandardTest"),
    FilterFromValue = 2017,
    FilterToValue = 2018,
    IsDevMode = true,



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

    //Filter files
    FilterFromText = Text.From(FilterFromValue),
    FilterToText = Text.From(FilterToValue),

    ApplyFilterFrom = if FilterFromValue <> null then
            Table.SelectRows(FilterOutNonData, each Text.Start([Name], Text.Length(FilterFromText))>=FilterFromText)    
        else
            FilterOutNonData,

    ApplyFilterTo = if FilterToValue <> null then
            Table.SelectRows(ApplyFilterFrom, each Text.Start([Name], Text.Length(FilterToText))<=FilterToText)    
        else
            ApplyFilterFrom,



    //------------------------------------------------------------------------------
    //        Restrict to one file in Dev mode
    //------------------------------------------------------------------------------

    DevMode_FilterOneFile = if IsDevMode is null then
            ApplyFilterTo
        else if IsDevMode then 
            Table.FirstN(ApplyFilterTo, 1) 
        else 
            ApplyFilterTo,

    AddTableColumn = Table.AddColumn(DevMode_FilterOneFile, "tbl", each DataAccessFunction([Folder Path], [Name]), type table),
    ColumnNames = Table.ColumnNames(AddTableColumn[tbl]{0}),




    //------------------------------------------------------------------------------
    //        Expand and restrict to 100 entries if in Dev mode
    //------------------------------------------------------------------------------

    ExpandedRawData = Table.ExpandTableColumn(AddTableColumn, "tbl", ColumnNames, ColumnNames),
    DevMode_RestrictReturnedRecords =  if IsDevMode is null then
            ExpandedRawData
        else if IsDevMode then 
            Table.FirstN(ExpandedRawData, 100) 
        else 
            ExpandedRawData,



    //------------------------------------------------------------------------------
    //        Add calculated Columns
    //------------------------------------------------------------------------------

    SelectCalcColumns = Table.SelectRows(DataSchema, each Text.Start([OriginalFieldName], 1) = "<"),
    CalcColRemovePrefix = Table.TransformColumns(SelectCalcColumns, {"OriginalFieldName", each Text.End(_, Text.Length(_)-1), type text}),
    CalcColRemovePostFix = Table.TransformColumns(CalcColRemovePrefix, {"OriginalFieldName", each Text.Start(_, Text.Length(_)-1), type text}),
    CalcColZipList = List.Zip({CalcColRemovePostFix[FieldName], CalcColRemovePostFix[OriginalFieldName]}),

    AccumulatorFunction = (TableState, CurrentListItem)=>
    let
        FieldName = CurrentListItem{0},
        Formula = CurrentListItem{1},
        ReturnValue = try
                Table.AddColumn(TableState, FieldName, Expression.Evaluate("each" & Formula))
            otherwise
                Table.AddColumn(TableState, FieldName, each error "Incorrect formula error")
    in

        ReturnValue,
    AddCalcCols = List.Accumulate(CalcColZipList, DevMode_RestrictReturnedRecords, AccumulatorFunction),



    //------------------------------------------------------------------------------
    //        Change field names
    //------------------------------------------------------------------------------

    FilterOutCalculatedFields = Table.SelectRows(DataSchema, each Text.Start([OriginalFieldName],1) <> "<"),
    FilterFieldNamesToChange = Table.SelectRows(FilterOutCalculatedFields, each [FieldName] <> [OriginalFieldName]),
    FieldNameChangePairs = List.Zip({FilterFieldNamesToChange[OriginalFieldName], FilterFieldNamesToChange[FieldName]}),
    ChangeFieldNames = Table.RenameColumns(AddCalcCols, FieldNameChangePairs),

    
    //------------------------------------------------------------------------------
    //        Select Cols
    //------------------------------------------------------------------------------

    ColsToSelect = DataSchema[FieldName],
    SelectCols = Table.SelectColumns(ChangeFieldNames, ColsToSelect),


    //------------------------------------------------------------------------------
    //        Change Types
    //------------------------------------------------------------------------------

    AddTypeColToDataSchema = Table.AddColumn(DataSchema, "FieldType", each Expression.Evaluate([FieldTypeAsText]), type type),
    FieldTypePairs = List.Zip({AddTypeColToDataSchema[FieldName], AddTypeColToDataSchema[FieldType]}),
    TransformColTypes = Table.TransformColumnTypes(SelectCols, FieldTypePairs)




in
    TransformColTypes