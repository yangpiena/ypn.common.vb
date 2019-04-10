Attribute VB_Name = "ModNetUtils"
'---------------------------------------------------------------------------------------
' Module    : ModNetUtils
' Author    : Administrator
' Date      : 2018-4-5
' Purpose   : 网络类工具
'---------------------------------------------------------------------------------------
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : MRequestREST
' Author    : YPN
' Date      : 2018-4-5
' Purpose   : 请求/调用REST接口
' Param     : i_RequstURL        请求地址
'           : i_RequestParameter 请求参数
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
    MsgBox "Error " & Err.Number & " (请求服务失败！" & Err.Description & ") in procedure MRequestREST of Module ModNetUtils"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MSendEmail
' Author    : YPN
' Date      : 2018-04-25 17:29
' Purpose   : 发送电子邮件，按文本格式发送
' Param     : i_smtpServer SMTP服务器地址
'             i_from       发信人邮箱地址
'             i_password   发信人邮箱密码
'             i_to         收信人邮箱地址，多个地址间用英文分号;隔开
'             i_subject    邮件主题
'             i_body       邮件正文，文本格式
'             i_attachment（可选）附件地址，例如："D:\1.txt"
' Return    : String       发送成功则返回success，失败则返回failure加失败原因
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSendEmail(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String) As String
    Dim v_nameSpace As String
    Dim v_email     As New CDO.Message     'CDO.message是一个发送邮件的对象。引用路径：C:\Windows\system32\cdosys.dll
    ' Dim v_email As Object
    ' Set v_email = CreateObject("CDO.Message") '创建对象，如果不引用，可以用此
    
    On Error GoTo MSendEmail_Error
    
    v_nameSpace = "http://schemas.microsoft.com/cdo/configuration/" '微软服务器网址，固定不用改
    v_email.From = i_from                  '发信人邮箱地址
    v_email.To = i_to                      '收信人邮箱地址
    v_email.Subject = i_subject            '邮件主题
    v_email.TextBody = i_body              '邮件正文，使用文本格式发送邮件
    If Not ModStringUtils.MIsNull(i_attachment) Then
        v_email.AddAttachment i_attachment '附件地址，例如："D:\1.txt"
    End If
    With v_email.Configuration.Fields
        .Item(v_nameSpace & "smtpserver") = i_smtpServer  'SMTP服务器地址
        .Item(v_nameSpace & "smtpserverport") = 25        'SMTP服务器端口
        .Item(v_nameSpace & "sendusing") = 2              '发送端口
        .Item(v_nameSpace & "smtpauthenticate") = 1       '需要提供用户名和密码，0是不提供
        .Item(v_nameSpace & "sendusername") = i_from      '发信人邮箱地址
        .Item(v_nameSpace & "sendpassword") = i_password  '发信人邮箱密码
        .Update
    End With
    v_email.Send                           '执行发送
    Set v_email = Nothing                  '发送成功后即时释放对象
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
' Purpose   : 发送电子邮件，按文本格式发送，支持抄送和加密抄送
' Param     : i_smtpServer SMTP服务器地址
'             i_from       发信人邮箱地址
'             i_password   发信人邮箱密码
'             i_to         收信人邮箱地址，多个地址间用英文分号;隔开
'             i_subject    邮件主题
'             i_body       邮件正文，文本格式
'             i_attachment（可选）附件地址，例如："D:\1.txt"
'             i_cc        （可选）抄送人邮箱地址
'             i_bcc       （可选）加密抄送人邮箱地址
' Return    : String       发送成功则返回success，失败则返回failure加失败原因
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSendEmail2(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String, Optional ByVal i_cc As String, Optional ByVal i_bcc As String) As String
    Dim v_nameSpace As String
    Dim v_email     As New CDO.Message     'CDO.message是一个发送邮件的对象。引用路径：C:\Windows\system32\cdosys.dll
    
    On Error GoTo MSendEmail2_Error
    
    v_nameSpace = "http://schemas.microsoft.com/cdo/configuration/" '微软服务器网址，固定不用改
    v_email.From = i_from                  '发信人邮箱地址
    v_email.To = i_to                      '收信人邮箱地址
    v_email.CC = i_cc                      '抄送人邮箱地址
    v_email.BCC = i_bcc                    '加密抄送人邮箱地址
    v_email.Subject = i_subject            '邮件主题
    v_email.TextBody = i_body              '邮件正文，使用文本格式发送邮件
    If Not ModStringUtils.MIsNull(i_attachment) Then
        v_email.AddAttachment i_attachment '附件地址，例如："D:\1.txt"
    End If
    With v_email.Configuration.Fields
        .Item(v_nameSpace & "smtpserver") = i_smtpServer  'SMTP服务器地址
        .Item(v_nameSpace & "smtpserverport") = 25        'SMTP服务器端口
        .Item(v_nameSpace & "sendusing") = 2              '发送端口
        .Item(v_nameSpace & "smtpauthenticate") = 1       '需要提供用户名和密码，0是不提供
        .Item(v_nameSpace & "sendusername") = i_from      '发信人邮箱地址
        .Item(v_nameSpace & "sendpassword") = i_password  '发信人邮箱密码
        .Update
    End With
    v_email.Send                           '执行发送
    Set v_email = Nothing                  '发送成功后即时释放对象
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
' Purpose   : 发送电子邮件，按HTML格式发送
' Param     : i_smtpServer SMTP服务器地址
'             i_from       发信人邮箱地址
'             i_password   发信人邮箱密码
'             i_to         收信人邮箱地址，多个地址间用英文分号;隔开
'             i_subject    邮件主题
'             i_body       邮件正文，HTML格式
'             i_attachment（可选）附件地址，例如："D:\1.txt"
' Return    : String       发送成功则返回success，失败则返回failure加失败原因
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSendHTMLEmail(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String) As String
    Dim v_nameSpace As String
    Dim v_email     As New CDO.Message 'CDO.message是一个发送邮件的对象。引用路径：C:\Windows\system32\cdosys.dll
    
    On Error GoTo MSendHTMLEmail_Error
    
    v_nameSpace = "http://schemas.microsoft.com/cdo/configuration/" '微软服务器网址，固定不用改
    v_email.From = i_from                  '发信人邮箱地址
    v_email.To = i_to                      '收信人邮箱地址
    v_email.Subject = i_subject            '邮件主题
    v_email.HTMLBody = i_body              '邮件正文，使用html格式发送邮件
    If Not ModStringUtils.MIsNull(i_attachment) Then
        v_email.AddAttachment i_attachment '附件地址，例如："D:\1.txt"
    End If
    With v_email.Configuration.Fields
        .Item(v_nameSpace & "smtpserver") = i_smtpServer  'SMTP服务器地址
        .Item(v_nameSpace & "smtpserverport") = 25        'SMTP服务器端口
        .Item(v_nameSpace & "sendusing") = 2              '发送端口
        .Item(v_nameSpace & "smtpauthenticate") = 1       '需要提供用户名和密码，0是不提供
        .Item(v_nameSpace & "sendusername") = i_from      '发送人邮箱地址
        .Item(v_nameSpace & "sendpassword") = i_password  '发送人邮箱密码
        .Update
    End With
    v_email.Send                           '执行发送
    Set v_email = Nothing                  '发送成功后即时释放对象
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
' Purpose   : 发送电子邮件，按HTML格式发送，支持抄送和加密抄送
' Param     : i_smtpServer SMTP服务器地址
'             i_from       发信人邮箱地址
'             i_password   发信人邮箱密码
'             i_to         收信人邮箱地址，多个地址间用英文分号;隔开
'             i_subject    邮件主题
'             i_body       邮件正文，HTML格式
'             i_attachment（可选）附件地址，例如："D:\1.txt"
'             i_cc        （可选）抄送人邮箱地址
'             i_bcc       （可选）加密抄送人邮箱地址
' Return    : String       发送成功则返回success，失败则返回failure加失败原因
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSendHTMLEmail2(ByVal i_smtpServer As String, ByVal i_from As String, ByVal i_password As String, ByVal i_to As String, _
    ByVal i_subject As String, ByVal i_body As String, Optional ByVal i_attachment As String, Optional ByVal i_cc As String, Optional ByVal i_bcc As String) As String
    Dim v_nameSpace As String
    Dim v_email     As New CDO.Message 'CDO.message是一个发送邮件的对象。引用路径：C:\Windows\system32\cdosys.dll
    
    On Error GoTo MSendHTMLEmail2_Error
    
    v_nameSpace = "http://schemas.microsoft.com/cdo/configuration/" '微软服务器网址，固定不用改
    v_email.From = i_from                  '发信人邮箱地址
    v_email.To = i_to                      '收信人邮箱地址
    v_email.CC = i_cc                      '抄送人邮箱地址
    v_email.BCC = i_bcc                    '加密抄送人邮箱地址
    v_email.Subject = i_subject            '邮件主题
    v_email.HTMLBody = i_body              '邮件正文，使用html格式发送邮件
    If Not ModStringUtils.MIsNull(i_attachment) Then
        v_email.AddAttachment i_attachment '附件地址，例如："D:\1.txt"
    End If
    With v_email.Configuration.Fields
        .Item(v_nameSpace & "smtpserver") = i_smtpServer  'SMTP服务器地址
        .Item(v_nameSpace & "smtpserverport") = 25        'SMTP服务器端口
        .Item(v_nameSpace & "sendusing") = 2              '发送端口
        .Item(v_nameSpace & "smtpauthenticate") = 1       '需要提供用户名和密码，0是不提供
        .Item(v_nameSpace & "sendusername") = i_from      '发送人邮箱地址
        .Item(v_nameSpace & "sendpassword") = i_password  '发送人邮箱密码
        .Update
    End With
    v_email.Send                           '执行发送
    Set v_email = Nothing                  '发送成功后即时释放对象
    MSendHTMLEmail2 = "success"
    
    On Error GoTo 0
    Exit Function
    
MSendHTMLEmail2_Error:
    MSendHTMLEmail2 = "failure:" & Err.Number & " (" & Err.Description & ") in procedure MSendHTMLEmail_Error of Module ModNetUtils"
End Function
