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
    
    ReturnValue = if Record.HasFields(ConverterRecord, TypeAsText) then
            Record.Field(ConverterRecord, TypeAsText)
        else
            fn_RaiseConsolidationError("Incorrect input type " & TypeAsText)
in
    ReturnValue