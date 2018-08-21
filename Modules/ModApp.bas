Attribute VB_Name = "ModApp"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : MGetVersion
' Author    : YPN
' Date      : 2018/08/21 17:13
' Purpose   : 获取App的版本号
' Param     :
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetVersion(Optional ByVal i_App As Object) As String
    
    On Error GoTo MGetVersion_Error
    
    If i_App Is Nothing Then
        MGetVersion = App.Major & "." & App.Minor & "." & App.Revision
    Else
        If Not (TypeOf i_App Is App) Then Err.Raise 5
        MGetVersion = i_App.Major & "." & i_App.Minor & "." & i_App.Revision
    End If
    
    On Error GoTo 0
    Exit Function
    
MGetVersion_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetVersion of Module ModApp"
    MGetVersion = ""
    
End Function
