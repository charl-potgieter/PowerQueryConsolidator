let

    AllowedConsolidateAcross = {"Current workbook", "Another workbook", "Folder"},
    

    CustomFunctionType = type function (
        ConsolidateAcross  as (type text meta[
            Documentation.AllowedValues = AllowedConsolidateAcross
            ]), 
        DataAccessFunction as (type text meta[
            Documentation.AllowedValues = ConsolidationDataAccessMethods
            ]), 
        optional SourceFolder as (type text), 
        optional FilterFileNameFrom as (type text),
        optional FilterFileNameTo as (type text), 
        optional IsDevMode as (type logical)  
        )
        as table meta [
            Documentation.Name =  "fn_Consolidation", 
            Documentation.LongDescription = "Consolidates file in a folder, selects columns, sets types.", 
            Documentation.Examples = {[
                Description =  "www.powernumerics.com/tba" , 
                Code = "", 
                Result = ""]}
        ],


    ReturnValue = Value.ReplaceType(fn_ConsolidationRaw, CustomFunctionType)


in
    ReturnValue