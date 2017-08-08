Attribute VB_Name = "RegularExpressions"
Function RegexCountMatches(str As String, Reg As String) As String
'Returns the number of matches found for a given regex
'str - string to test the regex on
'reg - the regular expression
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp"): regex.pattern = Reg: regex.Global = True
    If regex.test(str) Then
        Set Matches = regex.Execute(str)
        RegexCountMatches = Matches.Count
        Exit Function
    End If
End Function
Function RegexExecute(str As String, Reg As String, Optional findOnlyFirstMatch As Boolean = False) As Object
'Executes a Regular Expression on a provided string and returns all matches
'str - string to execute the regex on
'reg - the regular expression
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp"): regex.pattern = Reg
    regex.Global = Not (findOnlyFirstMatch)
    If regex.test(str) Then
        Set RegexExecute = regex.Execute(str)
        Exit Function
    End If
End Function
Function RegexExecuteGet(str As String, Reg As String, Optional matchIndex As Long = 0, Optional subMatchIndex As Long = 0) As String
'Executes a Regular Expression on a provided string and returns a selected submatch
'str - string to execute the regex on
'reg - the regular expression with at least 1 capture '()'
'matchIndex - the index of the match you want to return (default: 0)
'subMatchIndex - the index of the submatch you want to return (default: 0)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp"): regex.pattern = Reg
    regex.Global = Not (matchIndex = 0 And subMatchIndex = 0) 'For efficiency
    If regex.test(str) Then
        Set Matches = regex.Execute(str)
        RegexExecuteGet = Matches(matchIndex).SubMatches(subMatchIndex)
        Exit Function
    End If
End Function
Function RegexTest(str As String, Reg As String) As Boolean
'Executes a Regular Expression on a provided string and returns a selected submatch
'str - string to test the regex on
'reg - the regular expression
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp"): regex.pattern = Reg:    regex.Global = False
    If regex.test(str) Then
        RegexTest = True
        Exit Function
    End If
    RegexTest = False
End Function
Function RegexReplace(str As String, Reg As String, replaceStr As String, Optional replaceLimit As Long = -1) As String
Attribute RegexReplace.VB_Description = "Replace a pattern within a string with the provided replacement string based on all captures of the specified regular expression"
Attribute RegexReplace.VB_ProcData.VB_Invoke_Func = " \n9"
'Replaces a string using Regular Expressions
'str - string within which reg pattern will be replaced with replaceStr
'reg - the regular expression matching substrings to replace
'replaceStr - the string with which the reg pattern substrings are to be replaced with
'replaceLimit - by default unlimited (-1). Providing value will limit the number of performed replacements
    If replaceLimit = 0 Then
        RegexReplace = str: Exit Function
    End If
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp"): regex.pattern = Reg: regex.Global = IIf(replaceLimit = -1, True, False)
    If replaceLimit <> -1 And replaceLimit <> 1 Then
        RegexReplace = str
        Dim i As Long
        For i = 1 To replaceLimit
            RegexReplace = RegexReplace(RegexReplace, Reg, replaceStr, 1)
        Next i
        Exit Function
    End If
    RegexReplace = regex.Replace(str, replaceStr)
End Function
