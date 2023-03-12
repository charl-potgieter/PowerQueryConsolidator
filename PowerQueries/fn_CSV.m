(Folder as text, FName as text)=>
let
    Source = Csv.Document(File.Contents(Folder & FName),[Delimiter=",", Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromoteHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromoteHeaders