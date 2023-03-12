(Source as table, FilterFrom as any)=>
let
    FilterFromText = Text.From(FilterFrom),
    FilterLength = Text.Length(FilterFromText),
    ReturnValue = Table.SelectRows(Source, each Text.Start([Name], FilterLength)>=FilterFromText)    
in
    ReturnValue