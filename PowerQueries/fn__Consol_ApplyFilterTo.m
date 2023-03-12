(Source as table, FilterTo as any)=>
let
    FilterToText = Text.From(FilterTo),
    FilterLength = Text.Length(FilterToText),
    ReturnValue = Table.SelectRows(Source, each Text.Start([Name], FilterLength)<=FilterToText)    
in
    ReturnValue