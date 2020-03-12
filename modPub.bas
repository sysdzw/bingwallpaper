Attribute VB_Name = "modPub"
'==============================================================================================
'名    称：必应壁纸每日更新
'描    述：程序每天从必应搜索首页下载高清大图设置到桌面背景
'使用方法：双击即可
'编    程：sysdzw 原创开发，如果有需要对模块扩充或更新的话请邮箱发我一份
'发布日期：2016-6-15
'博    客：http://blog.csdn.net/sysdzw
'Email   ：sysdzw@gmail.com
'QQ      ：171977759
'版    本：V1.0.0   初版                                                        2016-6-15
'          V1.1.53  因必应更新图片路径导致无法下载，本程序做了对应更新          2018-9-12
'==============================================================================================
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim strWallPaperLocal As String
Const BING_PICTURE_DIR = "D:\Bing\" '壁纸保存目录
Dim isHasUpdated As Boolean

Sub Main()
    App.TaskVisible = False
    If App.PrevInstance Then
        MsgBox "该程序已经在后台运行中了，请勿重复运行！", vbExclamation
    Else
        Wait 30000
'        If Command <> "silent" Then MsgBox "请点击“确定”，程序会在后台自动到必应上每日更新壁纸并设置到您的电脑桌面上", vbInformation
        Do
            If Dir(BING_PICTURE_DIR, vbDirectory) = "" Then MkDir BING_PICTURE_DIR
            strWallPaperLocal = BING_PICTURE_DIR & Format(Now, "yyyymmdd") & ".jpg"
            If Dir(strWallPaperLocal) = "" Then
                writeToFile "运行日志.txt", Now & vbTab & "发现目标文件" & strWallPaperLocal & "为空，准备用InternetCheckConnection检测网络，如果网络正常则调用函数flushWallPaper下载图片并更新桌面", False
                If InternetCheckConnection("http://cn.bing.com/", &H1, 0&) <> 0 Then
                    writeToFile "运行日志.txt", Now & vbTab & "网络正常，开始调用函数flushWallPaper", False
                    If isHasUpdated Then
                        Exit Sub
'                        Shell "bingwallpaper.exe"
'                        Shell "restart_pro.bat", 0
                    Else
                        Call flushWallPaper
                        Exit Sub '更新完成就退出
                        isHasUpdated = True
                    End If
                    writeToFile "运行日志.txt", Now & vbTab & "图片下载完毕并保存到" & strWallPaperLocal & "，已经更新桌面壁纸", False
                End If
            Else
                Exit Sub '如果本地文件中已经存在就退出
            End If
            Wait 20000 '延时5秒检测一次
'            writeToFile "运行日志.txt", Now & vbTab & "循环检测一次", False
        Loop
    End If
End Sub
'更新墙纸
Private Sub flushWallPaper()
'    On Error GoTo Err1
    Dim strWallPaperUrl$, i1&, i2&, strData$, XmlHttp As Object, Temp() As Byte
    
     '得到墙纸的url地址
    Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
    XmlHttp.Open "GET", "http://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1", False
    writeToFile "运行日志.txt", Now & vbTab & "开始执行XmlHttp.Send", False
    XmlHttp.Send
    writeToFile "运行日志.txt", Now & vbTab & "XmlHttp.Send执行完毕！", False
    strData = StrConv(XmlHttp.ResponseBody, vbUnicode) '得到页面源代码
    i1 = InStr(strData, "url"":""")
    i2 = InStr(strData, """,""urlbase")
    If i1 > 0 And i2 > 0 Then strWallPaperUrl = "https://cn.bing.com" & Mid(strData, i1 + 6, i2 - i1 - 6)
    If strWallPaperUrl <> "" Then '下载图片文件
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
    writeToFile "运行日志.txt", Now & vbTab & "发生错误：" & Err.Number & vbCrLf & Err.Description, "Private Sub flushWallPaper()"
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：根据所给文件名和内容直接写文件
'函数名：writeToFile
'入口参数(如下)：
'  strFileName 所给的文件名；
'  strContent 要输入到上述文件的字符串
'  isCover 是否覆盖该文件，默认为覆盖
'返回值：True或False，成功则返回前者，否则返回后者
'备注：sysdzw 于 2007-5-2 提供
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
'延时，单位为毫秒
Public Function Wait(ByVal MilliSeconds As Long)
    Dim dSavetime As Double
    dSavetime = timeGetTime + MilliSeconds   '记下开始时的时间
    While timeGetTime < dSavetime '循环等待
        DoEvents '转让控制权，以便让操作系统处理其它的事件
    Wend
End Function
