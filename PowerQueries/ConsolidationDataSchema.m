let
    Source = Excel.CurrentWorkbook(){[Name="tbl_DataSchema"]}[Content],
    Unpivot = Table.UnpivotOtherColumns(Source, {"FieldName", "FieldTypeAsText"}, "DataSource", "OriginalFieldName"),
    ChangedType = Table.TransformColumnTypes(Unpivot,{{"FieldName", type text}, {"FieldTypeAsText", type text}, {"DataSource", type text}, {"OriginalFieldName", type text}}),
    FilterOutNull = Table.SelectRows(ChangedType, each ([FieldName] <> null))
in
    FilterOutNull