let

    Example_DataRootFolderAdjusted = if Text.End(Example_DataRootFolder, 1) <> "/" then
            Example_DataRootFolder & "/"
        else
            Example_DataRootFolder,

    DataSourceName = "Standard data",

    ReturnTable = fn_FieldNamesDataVersusSchemaCheck(
        fn_ExcelFirstSheet,
        Example_DataRootFolderAdjusted & Example_SelectedDataSet,
        Example_SchemaFilePath,
        DataSourceName, 
        "Fields in data not in schema"
    )
in
    ReturnTable