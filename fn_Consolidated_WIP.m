let


// ------------------------------------------------------------------------------------------------------------------
//                      Main function
// ------------------------------------------------------------------------------------------------------------------ 
    fn_Main = (
        DataAccessMethod as text, 
        SourceFolder as text, 
        optional fn_DataAccessCustom, 
        optional FilterFileNameFrom, 
        optional FilterFileNameTo, 
        optional IsDevMode)=>
        
    let

        ParameterChecksAreOk = fn_ParameterChecksAreOk(DataAccessMethod, SourceFolder, fn_DataAccessCustom, FilterFileNameFrom, FilterFileNameTo, IsDevMode),

        ReturnValue = 
            if Record.Field(try ParameterChecksAreOk, "HasError") then
                ParameterChecksAreOk
            else
                try "Success"
            otherwise
                error "Error in main"
                
    in
        ReturnValue,

    // Replace the standard function type with the custom function type
    ReturnValue = Value.ReplaceType(fn_Main, CustomFunctionType),



// ------------------------------------------------------------------------------------------------------------------
//                      Custom Function Type
// ------------------------------------------------------------------------------------------------------------------ 
    AllowedDataAccessMethods = {"CSV", "First sheet", "Custom"},

    CustomFunctionType = type function (
        DataAccessMethod as (type text meta[
            Documentation.AllowedValues = AllowedDataAccessMethods
            ]), 
        SourceFolder as (type text), 
        optional fn_DataAccessCustom as (type function), 
        optional FilterFileNameFrom as (type text),
        optional FilterFileNameTo as (type text), 
        optional IsDevMode as (type logical)  
        )
        as table meta [
            Documentation.Name =  "fn_Consolidates", 
            Documentation.LongDescription = "Consolidates file in a folder, selects columns, sets types.", 
            Documentation.Examples = {[
                Description =  "www.powernumerics.com/tba" , 
                Code = "", 
                Result = ""]}
        ],



// ------------------------------------------------------------------------------------------------------------------
//                      fn_ParameterChecksAreOk
// ------------------------------------------------------------------------------------------------------------------ 
    fn_ParameterChecksAreOk= (
        DataAccessMethod as text, 
        SourceFolder as text, 
        optional fn_DataAccessCustom, 
        optional FilterFileNameFrom, 
        optional FilterFileNameTo, 
        optional IsDevMode)
    as logical=>
    let
        FilterParametersAreSameLength = fn_FilterParametersAreSameLength (FilterFileNameFrom, FilterFileNameTo),
        
        
        Result = FilterParametersAreSameLength,
        ReturnValue = try 
                Result 
            otherwise if fn_CustomErrorExists(Result) then 
                Result
            else
                fn_RaiseCustomError("Unhanded exception in fn_ParameterChecksAreOk")
    in
        ReturnValue,



// ------------------------------------------------------------------------------------------------------------------
//                      fn_FilterParametersAreSameLength
// ------------------------------------------------------------------------------------------------------------------ 
    fn_FilterParametersAreSameLength = 
(optional FilterFileNameFrom, optional FilterFileNameTo)
as logical=>
let
    Result =if FilterFileNameFrom is null and FilterFileNameTo is null then
            true
        else if FilterFileNameFrom is null and not(FilterFileNameTo is null) then
            fn_RaiseCustomError("Filter parameters are different lengths")
        else if not(FilterFileNameFrom is null) and FilterFileNameTo is null then
            fn_RaiseCustomError("Filter parameters are different lengths")
        else if Text.Length(Text.From(FilterFileNameFrom)) <> Text.Length(Text.From(FilterFileNameTo)) then 
            fn_RaiseCustomError("Filter parameters are different lengths")
        else
            true,
            
    ReturnValue = try 
            Result 
        otherwise if fn_CustomErrorExists(Result) then 
            Result
        else
            fn_RaiseCustomError("Unhanded exception in fn_FilterParametersAreSameLength")

in
    ReturnValue,



// ------------------------------------------------------------------------------------------------------------------
//                      fn_DataAccessMethodIsOk
// ------------------------------------------------------------------------------------------------------------------ 
    fn_DataAccessMethodIsOk = 
    (DataAccessMethod as text)=>
    let
        Result = if List.Contains(AllowedDataAccessMethods, DataAccessMethod) then
                true
            else
                error "Unkown data access method parameter",
        ReturnValue = try Result otherwise "Error in procedure fn_DataAccessMethodIsOk"
    in
        ReturnValue,
    

// ------------------------------------------------------------------------------------------------------------------
//                      fn_DataAccessFirstSheet
// ------------------------------------------------------------------------------------------------------------------
    fn_DataAccessFirstSheet = 
    (Folder as text, FName as text)=>
    let
        Source = Excel.Workbook(File.Contents(Folder & FName), true, true),
        Navigation = Table.SelectRows(Source, each [Kind] = "Sheet")[Data]{0},
        ReturnValue = try Navigation otherwise error "Error in procedure fn_DataAccessFirstSheet"
    in
        ReturnValue,



// ------------------------------------------------------------------------------------------------------------------
//                      DataAccessFunctionSelected
// ------------------------------------------------------------------------------------------------------------------    
    DataAccessFunctionSelected = 
    let
        Result = if DataAccessMethod = "Custom" then
                fn_DataAccessCustom
            else if DataAccessMethod = "First sheet" then
                fn_DataAccessFirstSheet
            else
                error [],
        ReturnValue = try Result otherwise error "Error in procedure DataAccessFunctionSelected"
    in
        ReturnValue,



// ------------------------------------------------------------------------------------------------------------------
//                      fn_TypeConverter
// ------------------------------------------------------------------------------------------------------------------            
fn_TypeConverter = 
(TypeAsText as text)=>
let
    
    ConverterRecord = [
        type null = type null,
        type logical = type logical,
        type number = type number,
        type time = type time,
        type date = type date,
        type datetime = type datetime,
        type datetimezone = type datetimezone,
        type duration = type duration,
        type text = type text,
        type binary = type binary,
        type type = type type,
        type list = type list,
        type record = type record,
        type table = type table,
        type function = type function,
        type anynonnull = type anynonnull,
        type none = type none,
        Int64.Type = Int64.Type,
        Currency.Type = Currency.Type,
        Percentage.Type = Percentage.Type
    ],
    
    Result = Record.Field(ConverterRecord, TypeAsText),
    ReturnValue = try Result otherwise error "Error in fn_TypeConverter"
in
    ReturnValue,



// ------------------------------------------------------------------------------------------------------------------
//                      fn_GetRawData
// ------------------------------------------------------------------------------------------------------------------ 
    fn_GetRawData = 
(
    DataAccessMethod as text, 
    SourceFolder as text, 
    optional fn_DataAccessCustom, 
    optional FilterFileNameFrom, 
    optional FilterFileNameTo, 
    optional IsDevMode)
as table =>
let
    // Get folder contents and filter out non-data files
    FolderContents = Folder.Files(SourceFolder),
    FilterOutNonData = Table.SelectRows(FolderContents, each
        Text.Upper([Name]) <> "README.TXT" and
        Text.Upper([Name]) <> "THUMBS.DB" and
        Text.Upper([Extension]) <> ".SQL" and
        Text.Start([Name], 1) <> "~"
        ),
    
    // Custom table type avoids types being lost on table expansion
    FirstTable = DataAccessFunction(FilterOutNonData[Folder Path]{0}, FilterOutNonData[Name]{0}),
    CustomTableType = Value.Type(FirstTable),
    AddTableCol = Table.AddColumn(FilterOutNonData, "tbl", each DataAccessFunction([Folder Path], [Name]), CustomTableType),

    // Filter data per parameters (using same number of characters)
    FilterFileNameFromText = Text.From(FilterFileNameFrom),
    FilterFileNameToText = Text.From(FilterFileNameTo),
    FilterCharacterLength = Text.Length(FilterFileNameFromText),
    AddFilterCol = Table.AddColumn(AddTableCol, "FilterCol", each Text.Start([Name], FilterCharacterLength), type text),
    FilterFiles = Table.SelectRows(AddFilterCol, each ([FilterCol] >= FilterFileNameFromText) and ([FilterCol] <= FilterFileNameToText)), 
    DevMode_FilterOneFile = if IsDevMode is null then
            FilterFiles
        else if IsDevMode then 
            Table.FirstN(FilterFiles, 1) 
        else 
            FilterFiles,
    
    SelectTableCol = Table.SelectColumns(DevMode_FilterOneFile, {"tbl"}),
    
    // If no file exists return an empty table to prevent an expand error
    Expand = if Table.RowCount(SelectTableCol) = 0 then
            #table({},{})
        else
            Table.ExpandTableColumn(
                SelectTableCol, 
                "tbl", 
                Table.ColumnNames(SelectTableCol[tbl]{0}),
                Table.ColumnNames(SelectTableCol[tbl]{0}))
                
in
    Expand
                
// ------------------------------------------------------------------------------------------------------------------
//                      fn_RaiseCustomError
// ------------------------------------------------------------------------------------------------------------------ 
    fn_RaiseCustomError = (ErrorMessage as text) =>error [ 
                    Reason = "fn_Consolidate - raised error", 
                    Message = "", 
                    Detail = ErrorMessage 
                ],



// ------------------------------------------------------------------------------------------------------------------
//                      fn_CustomErrorExists
// ------------------------------------------------------------------------------------------------------------------ 
    fn_CustomErrorExists =(CurrentItem as any) =>
    let
        TryRecord = try CurrentItem,
        ReturnValue = if not TryRecord[HasError] then
                false
            else
                TryRecord[Error][Reason] = "fn_Consolidate - raised error"
    in
        ReturnValue


// ------------------------------------------------------------------------------------------------------------------
//                      fn_ReturnValueOrError
// ------------------------------------------------------------------------------------------------------------------ 

    // fn_ReturnValueOrError = (CurrentItem as any, CallingFunctionName as text)=>
    // try CurrentItem
    // otherwise if (try CurrentItem)[Error][Reason] = "fn_Consolidate - raised error" then
            // CurrentItem
        // else
           // fn_RaiseCustomError("Error in function " & CallingFunctionName)
    
    
    // let
        // TryRecord = try CurrentItem,
        // ReturnValue =  if not TryRecord[HasError] then
                //No errror exists, return CurrentItem
                // CurrentItem
            // else if TryRecord[Error][Reason] = "fn_Consolidate - raised error" then
                //Custom error has already been raised - return this
                // CurrentItem
            // else
                //Raise a new custom error
                // fn_RaiseCustomError("Error in function " & CallingFunctionName)
                
    // in
        // ReturnValue
        
in
    ReturnValue