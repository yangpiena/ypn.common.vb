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
    v_XmlHttp.Send (i_RequestParameter)
    MRequestREST = v_XmlHttp.responseText
    Set v_XmlHttp = Nothing
    
    On Error GoTo 0
    Exit Function
    
MRequestREST_Error:
    MsgBox "Error " & Err.Number & " (�������ʧ�ܣ�" & Err.Description & ") in procedure MRequestREST of Module ModNetUtils"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MSendEmail
' Author    : YPN
' Date      : 2018-04-25 17:29
' Purpose   : ���͵����ʼ������ı���ʽ����
' Param     : i_smtpServer SMTP��������ַ
'             i_from       �����������ַ
'             i_password   ��������������
'             i_to         �����������ַ�������ַ����Ӣ�ķֺ�;����
'             i_subject    �ʼ�����
'             i_body       �ʼ����ģ��ı���ʽ
'             i_attachment����ѡ��������ַ�����磺"D:\1.txt"
' Return    : String       ���ͳɹ��򷵻�success��ʧ���򷵻�failure��ʧ��ԭ��
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSendEmail(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String) As String
    Dim v_nameSpace As String
    Dim v_email     As New CDO.Message     'CDO.message��һ�������ʼ��Ķ�������·����C:\Windows\system32\cdosys.dll
    ' Dim v_email As Object
    ' Set v_email = CreateObject("CDO.Message") '����������������ã������ô�
    
    On Error GoTo MSendEmail_Error
    
    v_nameSpace = "http://schemas.microsoft.com/cdo/configuration/" '΢���������ַ���̶����ø�
    v_email.From = i_from                  '�����������ַ
    v_email.To = i_to                      '�����������ַ
    v_email.Subject = i_subject            '�ʼ�����
    v_email.TextBody = i_body              '�ʼ����ģ�ʹ���ı���ʽ�����ʼ�
    If Not ModStringUtils.MIsNull(i_attachment) Then
        v_email.AddAttachment i_attachment '������ַ�����磺"D:\1.txt"
    End If
    With v_email.Configuration.Fields
        .Item(v_nameSpace & "smtpserver") = i_smtpServer  'SMTP��������ַ
        .Item(v_nameSpace & "smtpserverport") = 25        'SMTP�������˿�
        .Item(v_nameSpace & "sendusing") = 2              '���Ͷ˿�
        .Item(v_nameSpace & "smtpauthenticate") = 1       '��Ҫ�ṩ�û��������룬0�ǲ��ṩ
        .Item(v_nameSpace & "sendusername") = i_from      '�����������ַ
        .Item(v_nameSpace & "sendpassword") = i_password  '��������������
        .Update
    End With
    v_email.Send                           'ִ�з���
    Set v_email = Nothing                  '���ͳɹ���ʱ�ͷŶ���
    MSendEmail = "success"
    
    On Error GoTo 0
    Exit Function
    
MSendEmail_Error:
    MSendEmail = "failure:" & Err.Number & " (" & Err.Description & ") in procedure MSendEmail of Module ModNetUtils"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MSendEmail2
' Author    : YPN
' Date      : 2019/04/10 14:35
' Purpose   : ���͵����ʼ������ı���ʽ���ͣ�֧�ֳ��ͺͼ��ܳ���
' Param     : i_smtpServer SMTP��������ַ
'             i_from       �����������ַ
'             i_password   ��������������
'             i_to         �����������ַ�������ַ����Ӣ�ķֺ�;����
'             i_subject    �ʼ�����
'             i_body       �ʼ����ģ��ı���ʽ
'             i_attachment����ѡ��������ַ�����磺"D:\1.txt"
'             i_cc        ����ѡ�������������ַ
'             i_bcc       ����ѡ�����ܳ����������ַ
' Return    : String       ���ͳɹ��򷵻�success��ʧ���򷵻�failure��ʧ��ԭ��
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSendEmail2(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String, Optional ByVal i_cc As String, Optional ByVal i_bcc As String) As String
    Dim v_nameSpace As String
    Dim v_email     As New CDO.Message     'CDO.message��һ�������ʼ��Ķ�������·����C:\Windows\system32\cdosys.dll
    
    On Error GoTo MSendEmail2_Error
    
    v_nameSpace = "http://schemas.microsoft.com/cdo/configuration/" '΢���������ַ���̶����ø�
    v_email.From = i_from                  '�����������ַ
    v_email.To = i_to                      '�����������ַ
    v_email.CC = i_cc                      '�����������ַ
    v_email.BCC = i_bcc                    '���ܳ����������ַ
    v_email.Subject = i_subject            '�ʼ�����
    v_email.TextBody = i_body              '�ʼ����ģ�ʹ���ı���ʽ�����ʼ�
    If Not ModStringUtils.MIsNull(i_attachment) Then
        v_email.AddAttachment i_attachment '������ַ�����磺"D:\1.txt"
    End If
    With v_email.Configuration.Fields
        .Item(v_nameSpace & "smtpserver") = i_smtpServer  'SMTP��������ַ
        .Item(v_nameSpace & "smtpserverport") = 25        'SMTP�������˿�
        .Item(v_nameSpace & "sendusing") = 2              '���Ͷ˿�
        .Item(v_nameSpace & "smtpauthenticate") = 1       '��Ҫ�ṩ�û��������룬0�ǲ��ṩ
        .Item(v_nameSpace & "sendusername") = i_from      '�����������ַ
        .Item(v_nameSpace & "sendpassword") = i_password  '��������������
        .Update
    End With
    v_email.Send                           'ִ�з���
    Set v_email = Nothing                  '���ͳɹ���ʱ�ͷŶ���
    MSendEmail2 = "success"
    
    On Error GoTo 0
    Exit Function
    
MSendEmail2_Error:
    MSendEmail2 = "failure:" & Err.Number & " (" & Err.Description & ") in procedure MSendEmail of Module ModNetUtils"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MSendHTMLEmail
' Author    : YPN
' Date      : 2018-04-25 17:29
' Purpose   : ���͵����ʼ�����HTML��ʽ����
' Param     : i_smtpServer SMTP��������ַ
'             i_from       �����������ַ
'             i_password   ��������������
'             i_to         �����������ַ�������ַ����Ӣ�ķֺ�;����
'             i_subject    �ʼ�����
'             i_body       �ʼ����ģ�HTML��ʽ
'             i_attachment����ѡ��������ַ�����磺"D:\1.txt"
' Return    : String       ���ͳɹ��򷵻�success��ʧ���򷵻�failure��ʧ��ԭ��
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSendHTMLEmail(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String) As String
    Dim v_nameSpace As String
    Dim v_email     As New CDO.Message 'CDO.message��һ�������ʼ��Ķ�������·����C:\Windows\system32\cdosys.dll
    
    On Error GoTo MSendHTMLEmail_Error
    
    v_nameSpace = "http://schemas.microsoft.com/cdo/configuration/" '΢���������ַ���̶����ø�
    v_email.From = i_from                  '�����������ַ
    v_email.To = i_to                      '�����������ַ
    v_email.Subject = i_subject            '�ʼ�����
    v_email.HTMLBody = i_body              '�ʼ����ģ�ʹ��html��ʽ�����ʼ�
    If Not ModStringUtils.MIsNull(i_attachment) Then
        v_email.AddAttachment i_attachment '������ַ�����磺"D:\1.txt"
    End If
    With v_email.Configuration.Fields
        .Item(v_nameSpace & "smtpserver") = i_smtpServer  'SMTP��������ַ
        .Item(v_nameSpace & "smtpserverport") = 25        'SMTP�������˿�
        .Item(v_nameSpace & "sendusing") = 2              '���Ͷ˿�
        .Item(v_nameSpace & "smtpauthenticate") = 1       '��Ҫ�ṩ�û��������룬0�ǲ��ṩ
        .Item(v_nameSpace & "sendusername") = i_from      '�����������ַ
        .Item(v_nameSpace & "sendpassword") = i_password  '��������������
        .Update
    End With
    v_email.Send                           'ִ�з���
    Set v_email = Nothing                  '���ͳɹ���ʱ�ͷŶ���
    MSendHTMLEmail = "success"
    
    On Error GoTo 0
    Exit Function
    
MSendHTMLEmail_Error:
    MSendHTMLEmail = "failure:" & Err.Number & " (" & Err.Description & ") in procedure MSendHTMLEmail_Error of Module ModNetUtils"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MSendHTMLEmail2
' Author    : YPN
' Date      : 2019/04/10 14:29
' Purpose   : ���͵����ʼ�����HTML��ʽ���ͣ�֧�ֳ��ͺͼ��ܳ���
' Param     : i_smtpServer SMTP��������ַ
'             i_from       �����������ַ
'             i_password   ��������������
'             i_to         �����������ַ�������ַ����Ӣ�ķֺ�;����
'             i_subject    �ʼ�����
'             i_body       �ʼ����ģ�HTML��ʽ
'             i_attachment����ѡ��������ַ�����磺"D:\1.txt"
'             i_cc        ����ѡ�������������ַ
'             i_bcc       ����ѡ�����ܳ����������ַ
' Return    : String       ���ͳɹ��򷵻�success��ʧ���򷵻�failure��ʧ��ԭ��
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSendHTMLEmail2(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String, Optional ByVal i_cc As String, Optional ByVal i_bcc As String) As String
    Dim v_nameSpace As String
    Dim v_email     As New CDO.Message 'CDO.message��һ�������ʼ��Ķ�������·����C:\Windows\system32\cdosys.dll
    
    On Error GoTo MSendHTMLEmail2_Error
    
    v_nameSpace = "http://schemas.microsoft.com/cdo/configuration/" '΢���������ַ���̶����ø�
    v_email.From = i_from                  '�����������ַ
    v_email.To = i_to                      '�����������ַ
    v_email.CC = i_cc                      '�����������ַ
    v_email.BCC = i_bcc                    '���ܳ����������ַ
    v_email.Subject = i_subject            '�ʼ�����
    v_email.HTMLBody = i_body              '�ʼ����ģ�ʹ��html��ʽ�����ʼ�
    If Not ModStringUtils.MIsNull(i_attachment) Then
        v_email.AddAttachment i_attachment '������ַ�����磺"D:\1.txt"
    End If
    With v_email.Configuration.Fields
        .Item(v_nameSpace & "smtpserver") = i_smtpServer  'SMTP��������ַ
        .Item(v_nameSpace & "smtpserverport") = 25        'SMTP�������˿�
        .Item(v_nameSpace & "sendusing") = 2              '���Ͷ˿�
        .Item(v_nameSpace & "smtpauthenticate") = 1       '��Ҫ�ṩ�û��������룬0�ǲ��ṩ
        .Item(v_nameSpace & "sendusername") = i_from      '�����������ַ
        .Item(v_nameSpace & "sendpassword") = i_password  '��������������
        .Update
    End With
    v_email.Send                           'ִ�з���
    Set v_email = Nothing                  '���ͳɹ���ʱ�ͷŶ���
    MSendHTMLEmail2 = "success"
    
    On Error GoTo 0
    Exit Function
    
MSendHTMLEmail2_Error:
    MSendHTMLEmail2 = "failure:" & Err.Number & " (" & Err.Description & ") in procedure MSendHTMLEmail_Error of Module ModNetUtils"
End Function
