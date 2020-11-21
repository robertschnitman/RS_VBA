Attribute VB_Name = "subs_registerUDF"
' This module is for registering the functions modules so that we can get a tooltip for them.
' Update this module when you add new functions.
' https://stackoverflow.com/questions/4262421/how-to-put-a-tooltip-on-a-user-defined-function


' Master subroutine to loop through all functions, their descriptions, and categories.
Sub RegisterUDF_all()

    ' GENERAL STEPS
    ' 1. Create arrays of function names and descriptions for each module.
    ' 2. Combine each array of function names/descriptions into a master array.
    ' 3. Apply a loop to iterate through each function category.


    ' 1. Create arrays of function names and descriptions for each module.
    ' funs_text
    m_names_text = Array("FINDREPLACE", "FINDREMOVE", "FINDBEFORE", "FINDAFTER", "FINDBETWEEN", "FIRSTNAME", "LASTNAME", "TEXTLIKE", "TEXTSTRIPWS", "TEXTINSERT", "TEXTREVERSE", "TEXTCOMPARE", "TEXTJOINR", "TRIML", "TRIMR", "TRIMLR")
    m_descs_text = Array("Find and replace a character(s).", _
                         "Remove a character(s).", _
                         "Find the substring before a specified character(s)", _
                         "Find the substring after a specified character(s)", _
                         "Find the substring between two characters", _
                         "Find the first name of a name string.", _
                         "Find the last name of a name string.", _
                         "Detect a pattern-match for a string.", _
                         "Remove all spaces in a string.", _
                         "Insert a character at a specified position.", _
                         "Reverse the order of a string.", _
                         "Compare two strings. Based on VBA's StrComp()", _
                         "Join a range of strings into a single string.", _
                         "Trim leading spaces.", _
                         "Trim trailing spaces.", _
                         "Trim leading and trailing spaces." _
                        )
                        
    ' funs_lookup
    m_names_lookup = Array("SLOOKUP", "SMATCH")
    m_descs_lookup = Array("Lookup a value by row value and column name.", "Determine which cell in a range matches a given pattern. The equivalent of XMATCH(pattern, range, 2) or MATCH(pattern, range, 0).")
    
    
    ' funs_logic
    m_names_logic = Array("ISLEN0", "IFBLANK", "SKIPBLANK", "DOIF")
    m_descs_logic = Array("Test whether a cell has no characters. Similar to ISBLANK().", _
                          "Similar to IF(), but performs an action depending on whether a cell is blank or not.", _
                          "Perform an action if a cell is non-blank; otherwise, output blank.", _
                          "Perform an action only if an action is met; otherwise, output blank.")
    
    ' funs_dates
    m_names_dates = Array("WEEKDAYNAME", "YMD", "MDY", "DMY")
    m_descs_dates = Array("Outputs the name of the weekday for a given date.", _
                          "Formats a date value into the ISO standard format (yyyy-mm-dd).", _
                          "Formats a date value into the month-day-year order (mm/dd/yyyy).", _
                          "Formats a date value into the day-month-year order (dd/mm/yyyy).")
                          
    ' 2. Combine each array of function names/descriptions into a master array.
    ' Combine the previous arrays together.
    m_names_all = Array(m_names_text, m_names_lookup, m_names_logic, m_names_dates)
    m_descs_all = Array(m_descs_text, m_descs_lookup, m_descs_logic, m_descs_dates)
                          
    m_categories = Array("RGS_Text", "RGS_Lookup", "RGS_Logic", "RGS_Dates")
    
    ' 3. Apply a loop to iterate through each function category.
    ' For each subarray of function lists and each function in each list, register each function's name, description, and category.
    For i = LBound(m_names_all) To UBound(m_names_all)
    
        ' We have to do a nested loop because m_names_all is an array of arrays (i.e. nested array).
        For j = LBound(m_names_all(i)) To UBound(m_names_all(i))
        
            Dim m_name, m_desc, m_cate As String
            m_name = m_names_all(i)(j) ' e.g. m_names_all(1)(1) = "FINDREPLACE"
            m_desc = m_descs_all(i)(j) ' e.g. m_descs_all(1)(1) = "Find and replace a character(s)."
            m_cate = m_categories(i)   ' e.g. m_categories(1)   = "RGS_Text"
                    
            Application.MacroOptions Macro:=m_name, Description:=m_desc, Category:=m_cate
        
        
        Next
    
    Next
    
End Sub
