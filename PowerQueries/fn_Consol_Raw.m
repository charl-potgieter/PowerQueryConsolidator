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

    //Uncomment for debugging
    DataAccessFunction = fn_ExcelFirstSheet,
    SourceFolder = "D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\PowerQueryConsolidator\Testing\",
    DataSchema = Table.SelectRows(ConsolidationDataSchema, each [DataSource] = "StandardTest"),
    FilterFromValue = 2017,
    FilterToValue = 2018,
    IsDevMode = false,



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
    Custom1 = Table.AddColumn(ApplyFilterTo, "tbl_raw", each fn_ExcelFirstSheet([Folder Path], [Name]), type table)
in
    Custom1