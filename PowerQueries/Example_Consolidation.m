let

    Example_DataRootFolderAdjusted = if Text.End(Example_DataRootFolder, 1) <> "/" then
            Example_DataRootFolder & "/"
        else
            Example_DataRootFolder,

    Source = fn_Consolidation(
        fn_ExcelFirstSheet,
        Example_DataRootFolderAdjusted & Example_SelectedDataSet,
        Example_SchemaFilePath,
        "StandardTest",
        2017,
        2018,
        false)
        
in
    Source