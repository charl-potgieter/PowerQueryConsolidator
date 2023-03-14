let
        Folder = "D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\PowerQueryConsolidator\Testing\Test_Consolidation\Test_Consolidation_Years\",
    FName = "2017_Test_File.xlsx",
    Source = fn_ExcelFirstSheet(Folder, FName)
in
    Source