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
        ]


in
    ReturnValue