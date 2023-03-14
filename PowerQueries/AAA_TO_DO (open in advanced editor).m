/*

Using an extra query fn_ConsolWithSchema + having 2 table columns to solve Formula.Firewall issue is messy and may use excessive memory?

Current code also wont work that well if a calcualted column needs to refer to a file name or folder

Is it worthwhile revisiting expanding first before applying schema?  Just need to be careful of name conflicts in Folder.Files view and the underlying file.



*/

""