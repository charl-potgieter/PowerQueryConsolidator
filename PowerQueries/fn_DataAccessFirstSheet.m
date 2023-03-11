(Folder as text, FName as text)=>
let
    Source = Excel.Workbook(File.Contents(Folder & FName), true, true),
    Navigation = Table.SelectRows(Source, each [Kind] = "Sheet")[Data]{0}
in
    Navigation