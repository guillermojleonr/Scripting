Attribute VB_Name = "IsDBSecured"
Option Compare Database

'Find if a DDBB is passwordprotected

'---------------------------------------------------------------------------------------
' Procedure : IsDbSecured
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Determine if a database is password protected or not
'             Returns True = Password Protected, False = Not Password Protected
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sDb       : Fully qualified path and filename w/ extension of the database to check
'
' Usage:
' ~~~~~~
' Call IsDbSecured("C:\Database\Test.accdb")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2014-Jan-02                 Initial Release
'---------------------------------------------------------------------------------------
Public Function IsDBSecured(ByVal sDb As String) As Boolean
    On Error GoTo Error_Handler
    Dim oAccess         As Access.Application
 
    Set oAccess = CreateObject("Access.Application")
    oAccess.Visible = True 'False
 
    'If an error occurs below, the db is password protected
    oAccess.DBEngine.OpenDatabase sDb, False
 
Error_Handler_Exit:
    On Error Resume Next
    oAccess.Quit acQuitSaveNone
    Set oAccess = Nothing
    Exit Function
 
Error_Handler:
    If Err.Number = 3031 Then
        IsDBSecured = True
    Else
        MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Source: IsDbSecured" & vbCrLf & _
               "Error Description: " & Err.Description, _
               vbCritical, "An Error has Occurred!"
    End If
    Resume Error_Handler_Exit
End Function
