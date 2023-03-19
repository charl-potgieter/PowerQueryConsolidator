let

    Example_DataRootFolderAdjusted = if Text.End(Example_DataRootFolder, 1) <> "/" then
            Example_DataRootFolder & "/"
        else
            Example_DataRootFolder,


    ReturnTable = fn_FieldNamesDataVersusSchemaCheck(
        fn_ExcelFirstSheet,
        Example_DataRootFolderAdjusted & Example_SelectedDataSet,
        Example_SchemaFilePath,
        "StandardTest", 
        "Fields in data not in schema"
    )
in
    ReturnTable