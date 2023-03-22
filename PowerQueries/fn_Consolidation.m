let


    //------------------------------------------------------------------------------
    //        Raw data ex metadata which is added at the end
    //------------------------------------------------------------------------------    
    
    fn_Consol_Raw = 
    (
        DataAccessFunction as function, 
        SourceFolder as text,
        SchemaFilePath as text,
        DataSourceName as text,
        optional FilterFromValue as any, 
        optional FilterToValue as any, 
        optional IsDevMode as logical
    )=>    
    let


        //------------------------------------------------------------------------------
        //        Uncomment for debugging
        //------------------------------------------------------------------------------
    /*    
        DataAccessFunction = fn_ExcelFirstSheet,
        SourceFolder = "",
        SchemaFilePath = "",
        DataSourceName = "Standard data",
        FilterFromValue = 2017,
        FilterToValue = 2018,
        IsDevMode = true,
    */


        //------------------------------------------------------------------------------
        //        DataSchema
        //------------------------------------------------------------------------------

        SchemaSource = Excel.Workbook(File.Contents(SchemaFilePath), null, true),
        SchemaTable  = SchemaSource{[Item = "tbl_DataSchema", Kind="Table"]}[Data],
        SchemaUnpivot = Table.UnpivotOtherColumns(SchemaTable, {"FieldName", "FieldTypeAsText"}, "DataSource", "OriginalFieldName"),
        SchemaChangeType = Table.TransformColumnTypes(SchemaUnpivot,{{"FieldName", type text}, {"FieldTypeAsText", type text}, {"DataSource", type text}, {"OriginalFieldName", type text}}),
        SchemaFilteredOnDataSource = Table.SelectRows(SchemaChangeType, each [DataSource] = DataSourceName),
        SchemaFilterOutNonUtilisedFields = Table.SelectRows(SchemaFilteredOnDataSource, each ([FieldName] <> null)),
        DataSchema = SchemaFilterOutNonUtilisedFields,

        //Below is utilised to delete unused fields later in the process
        DataSchemaUnusedFields = Table.SelectRows(SchemaFilteredOnDataSource, each [FieldName] = null)[OriginalFieldName],

        //Below is used for error checking before returning final output
        DataSchemaOriginalFieldName = Table.SelectRows(SchemaFilteredOnDataSource, each Text.Start([OriginalFieldName], 1) <> "<")[OriginalFieldName],
        DataSchemaFieldNames = SchemaFilterOutNonUtilisedFields[FieldName],
        DataSchemaOriginalFieldNamesAreUnique = List.Count(DataSchemaOriginalFieldName) = List.Count(List.Distinct(DataSchemaOriginalFieldName)),
        DataSchemaFieldNamesAreUnique = List.Count(DataSchemaFieldNames) = List.Count(List.Distinct(DataSchemaFieldNames)),


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
                
        //Prefix folder contents columns to avoid potential name conflicts with underlying data
        PrefixCols = Table.PrefixColumns(ApplyFilterTo, "PQ.Consol"),



        //------------------------------------------------------------------------------
        //        Restrict to one file in Dev mode
        //------------------------------------------------------------------------------

        DevMode_FilterOneFile = if IsDevMode is null then
                PrefixCols
            else if IsDevMode then 
                Table.FirstN(PrefixCols, 1) 
            else 
                PrefixCols,



        //------------------------------------------------------------------------------
        //        Get data tables, expand and restrict to 100 entries if in Dev mode
        //------------------------------------------------------------------------------


        AddTableColumn = Table.AddColumn(DevMode_FilterOneFile, "tbl", each DataAccessFunction([PQ.Consol.Folder Path], [PQ.Consol.Name]), type table),
        ColumnNames = Table.ColumnNames(AddTableColumn[tbl]{0}),
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
                    Table.AddColumn(TableState, FieldName, Expression.Evaluate("each " & Formula, #shared))
                otherwise
                    Table.AddColumn(TableState, FieldName, each error "Incorrect formula error")
        in

            ReturnValue,

        AddCalcCols = List.Accumulate(CalcColZipList, DevMode_RestrictReturnedRecords, AccumulatorFunction),


        //------------------------------------------------------------------------------
        //        Delete unused fields in data
        //------------------------------------------------------------------------------

        // Need to do this before rename in the evenyt a rename causes a name conflict with an unused field
        RemoveUnusedDataFields = Table.RemoveColumns(AddCalcCols, DataSchemaUnusedFields),
        

        //------------------------------------------------------------------------------
        //        Change field names
        //------------------------------------------------------------------------------

        FilterOutCalculatedFields = Table.SelectRows(DataSchema, each Text.Start([OriginalFieldName],1) <> "<"),
        FilterFieldNamesToChange = Table.SelectRows(FilterOutCalculatedFields, each [FieldName] <> [OriginalFieldName]),
        FieldNameChangePairs = List.Zip({FilterFieldNamesToChange[OriginalFieldName], FilterFieldNamesToChange[FieldName]}),
        ChangeFieldNames = Table.RenameColumns(RemoveUnusedDataFields, FieldNameChangePairs),

        
        //------------------------------------------------------------------------------
        //        Select and reorder Cols
        //------------------------------------------------------------------------------

        //Even though unused data columns are deleted above the columns returned by Folder.Files(...) also
        //need to be deleted.   Below action also re-orders
        ColsToSelectAndReorder = DataSchema[FieldName],
        SelectAndReorderCols = Table.SelectColumns(ChangeFieldNames, ColsToSelectAndReorder),


        //------------------------------------------------------------------------------
        //        Change Types
        //------------------------------------------------------------------------------

        AddTypeColToDataSchema = Table.AddColumn(DataSchema, "FieldType", each Expression.Evaluate([FieldTypeAsText], #shared), type type),
        FieldTypePairs = List.Zip({AddTypeColToDataSchema[FieldName], AddTypeColToDataSchema[FieldType]}),
        TransformColTypes = Table.TransformColumnTypes(SelectAndReorderCols, FieldTypePairs),



        //------------------------------------------------------------------------------
        //        Check for missing and extra fields in data compared to schema
        //------------------------------------------------------------------------------
       
        FieldsInDataNotSchema = Table.RowCount(
            fn_FieldNamesDataVersusSchemaCheck(
                DataAccessFunction, 
                SourceFolder, 
                SchemaFilePath, 
                DataSourceName, 
                "Fields in data not in schema")
            ) <> 0,

        FieldsInSchemaNotData = Table.RowCount(
            fn_FieldNamesDataVersusSchemaCheck(
                DataAccessFunction, 
                SourceFolder, 
                SchemaFilePath, 
                DataSourceName, 
                "Fields in schema not in data")
            ) <> 0,        


        ErrorChecks = if not DataSchemaOriginalFieldNamesAreUnique then
                error "Original field names in the schema are not unique"
            else if not DataSchemaFieldNamesAreUnique then
                error "Data field names in schema are not unique"
           else if FieldsInDataNotSchema then
                error "Error, fields exist in data but not schema, run fn_FieldNamesDataVersusSchemaCheck to identify items."
            else if FieldsInSchemaNotData then
                error "Error, fields exist in schema but not data, run fn_FieldNamesDataVersusSchemaCheck to identify items."
            else
                TransformColTypes         


    in
        ErrorChecks,   //End of Raw function ex-metadata



        //------------------------------------------------------------------------------
        //        Add function metadata
        //------------------------------------------------------------------------------

    CustomFunctionType = type function (
        DataAccessFunction as (type function), 
        SourceFolder as (type text),
        SchemaFilePath as (type text),
        DataSourceName as (type text),
        optional FilterFromValue as (type any), 
        optional FilterToValue as (type any), 
        optional IsDevMode as (type logical) 
        )
        as table meta [
            Documentation.Name =  "fn_Consolidation", 
            Documentation.LongDescription = "Consolidates file in a folder, selects columns, sets types.", 
            Documentation.Examples = {[
                Description =  "Please visit https://github.com/charl-potgieter/PowerQueryConsolidator for documentation" , 
                Code = "", 
                Result = ""]}
        ],


    ReturnValue = Value.ReplaceType(fn_Consol_Raw, CustomFunctionType)




in
    ReturnValue