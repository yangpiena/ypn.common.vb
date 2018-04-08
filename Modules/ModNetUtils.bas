Attribute VB_Name = "ModNetUtils"
'---------------------------------------------------------------------------------------
' Module    : ModNetUtils
' Author    : Administrator
' Date      : 2018-4-5
' Purpose   : �����๤��
'---------------------------------------------------------------------------------------

Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : MRequestREST
' Author    : YPN
' Date      : 2018-4-5
' Purpose   : ����/����REST�ӿ�
' Param     : i_RequstURL        �����ַ
'           : i_RequestParameter �������
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MRequestREST(ByVal i_RequstURL As String, ByVal i_RequestParameter As String) As String

    Dim v_XmlHttp
    
    On Error GoTo MRequestREST_Error
    
    Set v_XmlHttp = CreateObject("msxml2.xmlhttp")
    
    v_XmlHttp.Open "POST", i_RequstURL, False
    v_XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    v_XmlHttp.send (i_RequestParameter)
    
    MRequestREST = v_XmlHttp.responseText
    
    Set v_XmlHttp = Nothing
    
    On Error GoTo 0
    Exit Function
    
MRequestREST_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MRequestREST of Module ModNetUtils"
    
End Function
