let

    CustomFunctionType = type function (
        DataAccessFunction as (type text meta[
            Documentation.AllowedValues = Consol_DataAccessMethods
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


    ReturnValue = Value.ReplaceType(fn_Consol_Raw, CustomFunctionType)


in
    ReturnValue