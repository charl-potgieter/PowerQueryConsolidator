(DataAccessMethod as text, optional fn_DataAccessCustom)=> if DataAccessMethod = "Custom" then
       fn_DataAccessCustom
    else if DataAccessMethod = "First sheet" then
        fn_DataAccessFirstSheet
    else
        fn_RaiseConsolidationError("Invalid data access method selected")