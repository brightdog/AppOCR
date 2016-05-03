Attribute VB_Name = "modRegExp"
Option Explicit

Public Function RegExecute(ByRef strSource As String, ByVal strPattern As String, Optional ByVal bolGlobal As Boolean = True, Optional ByVal bolIgnorCase As Boolean = True, Optional ByVal bolMutilLine As Boolean = False) As VBScript_RegExp_55.MatchCollection

    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    Reg.Global = bolGlobal
    Reg.MultiLine = bolMutilLine
    Reg.IgnoreCase = bolIgnorCase
    
    Reg.Pattern = strPattern
    
    Dim Mc As VBScript_RegExp_55.MatchCollection
    
    Set Mc = Reg.Execute(strSource)
    
    
    Set RegExecute = Mc
    
    
    
    
    
    Set Reg = Nothing
    


End Function
