(optional FilterFileNameFrom, optional FilterFileNameTo)
as logical=>
let
    ReturnValue =if FilterFileNameFrom is null and FilterFileNameTo is null then
            true
        else if FilterFileNameFrom is null and not(FilterFileNameTo is null) then
            fn_RaiseConsolidationError("Filter parameters are different lengths")
        else if not(FilterFileNameFrom is null) and FilterFileNameTo is null then
            fn_RaiseConsolidationError("Filter parameters are different lengths")
        else if Text.Length(Text.From(FilterFileNameFrom)) <> Text.Length(Text.From(FilterFileNameTo)) then 
            fn_RaiseConsolidationError("Filter parameters are different lengths")
        else
            true
in
    ReturnValue