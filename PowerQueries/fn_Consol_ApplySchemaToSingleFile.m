(tbl_RawData as table, DataSchema as table)=>
let
    

    //------------------------------------------------------------------------------
    //        Uncomment for debugging
    //------------------------------------------------------------------------------

    /*
    tbl_RawData = ExampleSingleRawData,
    DataSchema = Table.SelectRows(ConsolidationDataSchema, each [DataSource] = "StandardTest"),
    */

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
    AddCalcCols = List.Accumulate(CalcColZipList, tbl_RawData, AccumulatorFunction),



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