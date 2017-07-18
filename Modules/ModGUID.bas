Attribute VB_Name = "ModGUID"
'---------------------------------------------------------------------------------------
' Module    : ModGUID
' Author    : YPN
' Date      : 2017-07-12 17:21
' Purpose   : ����һ��GUID
'
' UUID��GUID
'
'1. UUID: (Universally Unique Identifier) ͨ��Ψһ��ʶ��,��һ����ʶ����׼��������ܹ�,�ɿ�����������(OSF)��Ϊ�ֲ�ʽ���㻷��(DCE)��һ���ֶ������ı�׼��
'
'         UUID��Ŀ�����÷ֲ�ʽϵͳ�е�����Ԫ�ض�����Ψһ�ı�ʶ��Ѷ������Ҫ͸��������ƶ�����������Ѷ���ƶ������һ��ÿ���˶�����һ���������˲�ͬ�ı�
'
'         ʶ���������ڴ洢�����ݿ���ʱ���Ͳ��õ���������ͬ������(�����������ݿ��е�����,�������ݿ������ֻ����һ�ű�����Ч).
'
'         �����׼���ڱ��㷺Ӧ����΢���ȫ��Ψһ��ʶ����(GUID)��
'
'2. GUID:(Globally Unique Identifier) ȫ��Ψһ��ʶ��,��һ�����������������С�
'
'   GUID���ص�:
'
'    (1). ȫ��Ψһ�ԣ���������̨��������ɵ�GUID������ͬ,GUID��Ҫ����ӵ�ж���ڵ㡢��̨�������ɵļ���������ϵͳ�У��������Ψһ�Եı�־����
'
'         ��ʱ��Ϳռ��϶��ܱ�֤Ψһ�� , ��֤��ͬһʱ�䲻ͬ�ĵص����ɵ�GUIDֵ��ͬ?
'
'    (2). ��ɽṹ��ͨ���ض��㷨���ɵ�һ�������Ƴ���ΪΪ128���ַ���,����GUIDʱ�����㷨�Զ�����,����Ҫ�κλ�����������
'
'         GUID �ĸ�ʽΪ��xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx��������ÿ�� x �� 0-9 �� a-f ��Χ�ڵ�һ��ʮ�����Ƶ����֡�
'
'         ���磺6F9619FF-8B86-D011-B42D-00C04FC964FF ��Ϊ��Ч�� GUID ֵ��------>һ��16������4�������ƣ����Թ�32λ��
'
'    (3). Ӧ�ã������������û���ÿһ��Office�ĵ�����������Զ�����һ��GUIDֵ������Ϊ���Office��Ψһ��ʶ��;�������GUIDֵ����������������صģ�
'
'         �������GUIDֵ�������ǲ��ɼ���?���ߵ���Ϣ����ͨ��GUID��ֵ�ҵ�?
'---------------------------------------------------------------------------------------

Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long


'---------------------------------------------------------------------------------------
' Procedure : MGetGUID
' Author    : YPN
' Date      : 2017-07-12 17:24
' Purpose   : ����һ��GUID
' Param     :
' Return    : GetGUID(): bf8b9c642ea5426f82499bb60681671b
' Remark    : GUID:(Globally Unique Identifier) ȫ��Ψһ��ʶ��,��һ�����������������С�
'---------------------------------------------------------------------------------------
'
Public Function MGetGUID() As String
    
    Dim v_GUID As GUID
    
    If (CoCreateGuid(v_GUID) = 0) Then
        MGetGUID = _
        String(8 - Len(Hex$(v_GUID.Data1)), "0") & Hex$(v_GUID.Data1) & _
        String(4 - Len(Hex$(v_GUID.Data2)), "0") & Hex$(v_GUID.Data2) & _
        String(4 - Len(Hex$(v_GUID.Data3)), "0") & Hex$(v_GUID.Data3) & _
        IIf((v_GUID.Data4(0) < &H10), "0", "") & Hex$(v_GUID.Data4(0)) & _
        IIf((v_GUID.Data4(1) < &H10), "0", "") & Hex$(v_GUID.Data4(1)) & _
        IIf((v_GUID.Data4(2) < &H10), "0", "") & Hex$(v_GUID.Data4(2)) & _
        IIf((v_GUID.Data4(3) < &H10), "0", "") & Hex$(v_GUID.Data4(3)) & _
        IIf((v_GUID.Data4(4) < &H10), "0", "") & Hex$(v_GUID.Data4(4)) & _
        IIf((v_GUID.Data4(5) < &H10), "0", "") & Hex$(v_GUID.Data4(5)) & _
        IIf((v_GUID.Data4(6) < &H10), "0", "") & Hex$(v_GUID.Data4(6)) & _
        IIf((v_GUID.Data4(7) < &H10), "0", "") & Hex$(v_GUID.Data4(7))
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : MGetGUID2
' Author    : YPN
' Date      : 2017-07-13 15:31
' Purpose   :
' Param     : i_Format ��ʽ��"B"��"D"
' Return    : GetGUID("B"): {903c1236-fe24-43c2-b9b5-bec35d9a43a8}
'             GetGUID("D"): 17e316f4-3f5b-46a0-ad68-58abb816a969
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetGUID2(ByVal i_Format As String) As String
    
    Dim v_GUID    As GUID
    Dim v_Str     As String
    Dim v_Byte()  As Byte
    Dim v_Len     As Long
    Dim v_RealLen As Long
    
    v_Len = 40
    v_Byte = String(v_Len, 0)
    CoCreateGuid v_GUID                                                         '�ѽṹת��Ϊһ������ʾ���ַ���
    v_RealLen = StringFromGUID2(v_GUID, VarPtr(v_Byte(0)), v_Len)
    v_Str = v_Byte
    If (Asc(Mid$(v_Str, v_RealLen, 1)) = 0) Then v_RealLen = v_RealLen - 1
        
    If UCase$(Trim(i_Format)) = "B" Then
        MGetGUID2 = Left$(v_Str, v_RealLen)
    ElseIf UCase$(Trim(i_Format)) = "D" Then
        MGetGUID2 = Mid$(v_Str, 2, 36)
    End If
    
End Function
