let

    Example_DataRootFolderAdjusted = if Text.End(Example_DataRootFolder, 1) <> "/" then
            Example_DataRootFolder & "/"
        else
            Example_DataRootFolder,

    DataSourceName = if Example_SelectedDataSet = "04_NameConflict" then
            "Name conflict"
        else
            "Standard data",

    Source = fn_Consolidation(
        fn_ExcelFirstSheet,
        Example_DataRootFolderAdjusted & Example_SelectedDataSet,
        Example_SchemaFilePath,
        DataSourceName,
        2017,
        2018,
        false)
        
in
    Source