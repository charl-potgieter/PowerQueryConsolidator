(DataAccessMethod as text, optional fn_DataAccessCustom)=> if DataAccessMethod = "Custom" then
       fn_DataAccessCustom
    else if DataAccessMethod = "First sheet" then
        fn_Consol_DataAccessFirstSheet
    else
        fn_Consol_RaiseError("Invalid data access method selected")