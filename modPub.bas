Attribute VB_Name = "modPub"
'==============================================================================================
'��    �ƣ���Ӧ��ֽÿ�ո���
'��    ��������ÿ��ӱ�Ӧ������ҳ���ظ����ͼ���õ����汳��
'ʹ�÷�����˫������
'��    �̣�sysdzw ԭ���������������Ҫ��ģ���������µĻ������䷢��һ��
'�������ڣ�2016-6-15
'��    �ͣ�http://blog.csdn.net/sysdzw
'Email   ��sysdzw@gmail.com
'QQ      ��171977759
'��    ����V1.0.0   ����                                                        2016-6-15
'          V1.1.53  ���Ӧ����ͼƬ·�������޷����أ����������˶�Ӧ����          2018-9-12
'==============================================================================================
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim strWallPaperLocal As String
Const BING_PICTURE_DIR = "D:\Bing\" '��ֽ����Ŀ¼
Dim isHasUpdated As Boolean

Sub Main()
    App.TaskVisible = False
    If App.PrevInstance Then
        MsgBox "�ó����Ѿ��ں�̨�������ˣ������ظ����У�", vbExclamation
    Else
        Wait 30000
'        If Command <> "silent" Then MsgBox "������ȷ������������ں�̨�Զ�����Ӧ��ÿ�ո��±�ֽ�����õ����ĵ���������", vbInformation
        Do
            If Dir(BING_PICTURE_DIR, vbDirectory) = "" Then MkDir BING_PICTURE_DIR
            strWallPaperLocal = BING_PICTURE_DIR & Format(Now, "yyyymmdd") & ".jpg"
            If Dir(strWallPaperLocal) = "" Then
                writeToFile "������־.txt", Now & vbTab & "����Ŀ���ļ�" & strWallPaperLocal & "Ϊ�գ�׼����InternetCheckConnection������磬���������������ú���flushWallPaper����ͼƬ����������", False
                If InternetCheckConnection("http://cn.bing.com/", &H1, 0&) <> 0 Then
                    writeToFile "������־.txt", Now & vbTab & "������������ʼ���ú���flushWallPaper", False
                    If isHasUpdated Then
                        Exit Sub
'                        Shell "bingwallpaper.exe"
'                        Shell "restart_pro.bat", 0
                    Else
                        Call flushWallPaper
                        Exit Sub '������ɾ��˳�
                        isHasUpdated = True
                    End If
                    writeToFile "������־.txt", Now & vbTab & "ͼƬ������ϲ����浽" & strWallPaperLocal & "���Ѿ����������ֽ", False
                End If
            Else
                Exit Sub '��������ļ����Ѿ����ھ��˳�
            End If
            Wait 20000 '��ʱ5����һ��
'            writeToFile "������־.txt", Now & vbTab & "ѭ�����һ��", False
        Loop
    End If
End Sub
'����ǽֽ
Private Sub flushWallPaper()
'    On Error GoTo Err1
    Dim strWallPaperUrl$, i1&, i2&, strData$, XmlHttp As Object, Temp() As Byte
    
     '�õ�ǽֽ��url��ַ
    Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
    XmlHttp.Open "GET", "http://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1", False
    writeToFile "������־.txt", Now & vbTab & "��ʼִ��XmlHttp.Send", False
    XmlHttp.Send
    writeToFile "������־.txt", Now & vbTab & "XmlHttp.Sendִ����ϣ�", False
    strData = StrConv(XmlHttp.ResponseBody, vbUnicode) '�õ�ҳ��Դ����
    i1 = InStr(strData, "url"":""")
    i2 = InStr(strData, """,""urlbase")
    If i1 > 0 And i2 > 0 Then strWallPaperUrl = "https://cn.bing.com" & Mid(strData, i1 + 6, i2 - i1 - 6)
    If strWallPaperUrl <> "" Then '����ͼƬ�ļ�
        XmlHttp.Open "GET", strWallPaperUrl, False
        XmlHttp.Send
        If XmlHttp.ReadyState = 4 Then
            Temp() = XmlHttp.ResponseBody
            Open strWallPaperLocal For Binary As #1
            Put #1, , Temp()
            Close #1
        End If
        Set XmlHttp = Nothing
    
        SavePicture LoadPicture(strWallPaperLocal), BING_PICTURE_DIR & "Wallpaper1.bmp"
        SystemParametersInfo ByVal 20, True, ByVal BING_PICTURE_DIR & "Wallpaper1.bmp", 1
        Shell "rundll32 user32,UpdatePerUserSystemParameters"
    End If
    
    Exit Sub
Err1:
'    MsgBox Err.Number & vbCrLf & Err.Description, "Private Sub flushWallPaper()"
    writeToFile "������־.txt", Now & vbTab & "��������" & Err.Number & vbCrLf & Err.Description, "Private Sub flushWallPaper()"
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ����������ļ���������ֱ��д�ļ�
'��������writeToFile
'��ڲ���(����)��
'  strFileName �������ļ�����
'  strContent Ҫ���뵽�����ļ����ַ���
'  isCover �Ƿ񸲸Ǹ��ļ���Ĭ��Ϊ����
'����ֵ��True��False���ɹ��򷵻�ǰ�ߣ����򷵻غ���
'��ע��sysdzw �� 2007-5-2 �ṩ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function writeToFile(ByVal strFileName$, ByVal strContent$, Optional isCover As Boolean = True) As Boolean
    On Error GoTo Err1
    Dim fileHandl%
    fileHandl = FreeFile
    If isCover Then
        Open strFileName For Output As #fileHandl
    Else
        Open strFileName For Append As #fileHandl
    End If
    Print #fileHandl, strContent
    Close #fileHandl
    writeToFile = True
    Exit Function
Err1:
    writeToFile = False
End Function
'��ʱ����λΪ����
Public Function Wait(ByVal MilliSeconds As Long)
    Dim dSavetime As Double
    dSavetime = timeGetTime + MilliSeconds   '���¿�ʼʱ��ʱ��
    While timeGetTime < dSavetime 'ѭ���ȴ�
        DoEvents 'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼�
    Wend
End Function
