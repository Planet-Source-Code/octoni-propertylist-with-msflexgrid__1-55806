VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function GetMatchCollection(strPattern As String, strText As String) As MatchCollection
    Dim objRE As New RegExp
    Dim objMatches As MatchCollection
    Dim objMatch As Match



    objRE.Pattern = strPattern
    objRE.Global = True
    Set objMatches = objRE.Execute(strText)
    Set GetMatchCollection = objMatches

End Function




Function MatchRegex(strPattern As String, strTest As String) As Boolean
    Dim objRE As New RegExp
    Dim objMatches As MatchCollection
    Dim objMatch As Match

    objRE.Pattern = strPattern
    MatchRegex = objRE.Test(strTest)
End Function


Function Replace(strPattern As String, strTest As String, strReplace As String) As String
    Dim objRE As New RegExp
    Dim strExample As String


    objRE.Pattern = strPattern
    objRE.Global = True

    ' Replace occurences of "flat" with "juicy" '
    Replace = objRE.Replace(strTest, strReplace)
    Set objRE = Nothing
End Function

