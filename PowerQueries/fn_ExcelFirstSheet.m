(Folder as text, FName as text)=>
let
    Source = Excel.Workbook(File.Contents(Folder & FName), null, true),
    Navigation = Table.SelectRows(Source, each [Kind] = "Sheet")[Data]{0},
    PromoteHeaders = Table.PromoteHeaders(Navigation, [PromoteAllScalars=true])
in
    PromoteHeaders