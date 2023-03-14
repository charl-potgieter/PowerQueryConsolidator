let
    Source = Excel.CurrentWorkbook(){[Name="tbl_DataSources"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"SourceName", type text}, {"Source", type text}, {"DataAccessFunction", type text}})
in
    ChangedType