let

    Example_DataRootFolderAdjusted = if Text.End(Example_DataRootFolder, 1) <> "/" then
            Example_DataRootFolder & "/"
        else
            Example_DataRootFolder,

    DataSourceName = if Example_SelectedDataSet = "04_NameConflict" then
            "Name conflict"
        else
            "Standard data",


    ReturnTable = fn_FieldNamesDataVersusSchemaCheck(
        fn_ExcelFirstSheet,
        Example_DataRootFolderAdjusted & Example_SelectedDataSet,
        Example_SchemaFilePath,
        DataSourceName, 
        "Fields in schema not in data"
    )
in
    ReturnTable