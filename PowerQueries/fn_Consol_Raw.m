// Raw function excluding the documentation metadata

/*
(
    DataAccessFunction as text, 
    optional SourceFolder as text, 
    optional FilterFileNameFrom as any, 
    optional FilterFileNameTo as any, 
    optional IsDevMode as logical
)=>    
*/
let

    FilterFileNameFrom = 2017,
    FilterFileNameTo = 2018,
    ConsolidateAcross = "Folder",
    SourceFolder = "D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\PowerQueryConsolidator\Testing\",

    Source = Folder.Files(SourceFolder),

    //Filter files
    FilterFrom = if FilterFileNameFrom <> null then
            fn__Consol_ApplyFilterFrom(Source, FilterFileNameFrom)
        else
            Source,
     FilterTo = if FilterFileNameTo <> null then
            fn__Consol_ApplyFilterTo(FilterFrom, FilterFileNameTo)
        else
            FilterFrom
    

in
    FilterTo