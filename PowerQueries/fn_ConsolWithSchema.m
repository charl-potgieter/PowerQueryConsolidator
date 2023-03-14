let
    Source = Table.AddColumn(fn_Consol_Raw, "tbl", each fn_Consol_ApplySchemaToSingleFile([tbl_raw], Table.SelectRows(ConsolidationDataSchema, each [DataSource] = "StandardTest")))
in
    Source