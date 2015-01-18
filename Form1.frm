VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Cubase Translate - 병작"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10035
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10035
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "무시(&I)"
      Enabled         =   0   'False
      Height          =   405
      Left            =   8760
      TabIndex        =   35
      Top             =   6820
      Width           =   1215
   End
   Begin VB.CommandButton cmdReTrans 
      Caption         =   "번역(&T)"
      Height          =   390
      Left            =   6840
      TabIndex        =   34
      Top             =   6820
      Width           =   1215
   End
   Begin VB.CommandButton cmdTxtchk 
      Caption         =   "구문 검사(&C)"
      Height          =   375
      Left            =   5520
      TabIndex        =   32
      Top             =   6820
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "정지(&S)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   33
      Top             =   6820
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "시작(&S)"
      Height          =   375
      Left            =   2880
      TabIndex        =   31
      Top             =   6820
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar Pbstatus 
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   7350
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Max             =   7342
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   330
      Left            =   0
      TabIndex        =   24
      Top             =   7290
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmTranslate 
      Height          =   4935
      Left            =   2880
      TabIndex        =   13
      Top             =   1740
      Width           =   7095
      Begin VB.TextBox txtHex 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   6855
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5640
         TabIndex        =   22
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox txtOrigin 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   4080
         Width           =   6855
      End
      Begin VB.TextBox txtScript 
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3780
         Width           =   6855
      End
      Begin VB.TextBox txtProcess 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   6855
      End
      Begin VB.TextBox txtEdit 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   16
         Top             =   4380
         Width           =   6840
      End
      Begin VB.ListBox lstProcessing 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   3190
         Width           =   6855
      End
      Begin VB.ListBox lstErr 
         BackColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   6855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3960
         TabIndex        =   21
         Top             =   240
         Width           =   780
      End
      Begin VB.Label labProcess 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   20
         Top             =   4080
         Width           =   1215
      End
   End
   Begin VB.Frame frmFocus 
      Height          =   4935
      Left            =   2880
      TabIndex        =   11
      Top             =   1740
      Visible         =   0   'False
      Width           =   7095
      Begin VB.PictureBox FadeAnswerPictureBox 
         Appearance      =   0  '평면
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4545
         ScaleWidth      =   6825
         TabIndex        =   12
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2880
      ScaleHeight     =   615
      ScaleWidth      =   7095
      TabIndex        =   4
      Top             =   1080
      Width           =   7095
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Cubase Translate"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Frame frmSetting 
      Height          =   4935
      Left            =   2880
      TabIndex        =   3
      Top             =   1740
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CheckBox UTF8 
         Caption         =   "UTF-8 모드"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2760
         Value           =   1  '확인
         Width           =   2055
      End
      Begin VB.CheckBox Korea 
         Caption         =   "한글감지 및 건너뛰기 (BETA)"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   3135
      End
      Begin VB.ComboBox lan2 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "Form1.frx":65302
         Left            =   1200
         List            =   "Form1.frx":65318
         TabIndex        =   30
         Top             =   1920
         Width           =   495
      End
      Begin VB.ComboBox lan1 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "Form1.frx":65334
         Left            =   120
         List            =   "Form1.frx":6534A
         TabIndex        =   29
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox Google 
         Caption         =   "구글번역 사용"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   2055
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   600
         Max             =   5000
         Min             =   1
         TabIndex        =   9
         Top             =   240
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox txtSpeed 
         Height          =   270
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "10"
         Top             =   240
         Width           =   480
      End
      Begin VB.CheckBox chkClip 
         Caption         =   "클립보드 감시"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox lstLanguage 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "Form1.frx":65366
         Left            =   120
         List            =   "Form1.frx":65379
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label labto 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "에서"
         Height          =   180
         Left            =   720
         TabIndex        =   28
         Top             =   2000
         Width           =   360
      End
      Begin VB.Label labSpeed 
         AutoSize        =   -1  'True
         Caption         =   "번역 속도"
         Height          =   180
         Left            =   960
         TabIndex        =   10
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "번역할 언어 영역"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   6
         Top             =   780
         Width           =   1380
      End
   End
   Begin MSComctlLib.TreeView tvSetup 
      Height          =   6120
      Left            =   120
      TabIndex        =   2
      Top             =   1095
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   10795
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      Scroll          =   0   'False
      Appearance      =   1
   End
   Begin VB.Timer Timer 
      Interval        =   10
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   120
      ScaleHeight     =   930
      ScaleWidth      =   10095
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Label labTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "큐베이스 번역기"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   23
         Top             =   600
         Width           =   1425
      End
      Begin VB.Image imgMark 
         Height          =   720
         Left            =   0
         Picture         =   "Form1.frx":65391
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   1020
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Line lnBar 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   10080
      Y1              =   950
      Y2              =   950
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&
Private Const AC_SRC_OVER = &H0

Const DT_BOTTOM = &H8
Const DT_CENTER = &H1
Const DT_LEFT = &H0
Const DT_RIGHT = &H2
Const DT_TOP = &H0
Const DT_VCENTER = &H4
Const DT_WORDBREAK = &H10

Public StringToPrint

Private Type BitmapInfoHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BitmapInfo
    Header As BitmapInfoHEADER
    Colors As RGBQUAD
End Type

Dim Pixels() As Byte
Dim BackgroundBitmap As BitmapInfo

Dim BF            As BlendFunction
Dim lBF           As Long
Dim ThisRectangle As RECT
Dim Str           As String
Dim BackGroundDC  As Long
Dim iBitmap       As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BlendFunction
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private ieGoogle As Object

Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long) 'Conver to long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BitmapInfo, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BitmapInfo, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Public Function sUTF8ToUni(bySrc() As Byte) As String
Dim lBytes As Long, lNC As Long, lRet As Long

lBytes = UBound(bySrc) - LBound(bySrc) + 1
lNC = lBytes
sUTF8ToUni = String$(lNC, Chr(0))
lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bySrc(LBound(bySrc))), lBytes, StrPtr(sUTF8ToUni), lNC)
sUTF8ToUni = Left$(sUTF8ToUni, lRet)
End Function

Private Function ConvertUTF8File(sUTF8File As String) As String
Dim iFile As Integer, bData() As Byte, sData As String, lSize As Long

lSize = FileLen(sUTF8File)
If lSize > 0 Then
ReDim bData(0 To lSize - 1)

iFile = FreeFile()
Open sUTF8File For Binary As #iFile
Get #iFile, , bData
Close #iFile

sData = sUTF8ToUni(bData)
Else
sData = ""
End If
ConvertUTF8File = sData
End Function

Public Sub cmdIgnore_Click()
Inputs = txtEdit.Text
cmdIgnore.Enabled = False
End Sub

Private Sub cmdTxtchk_Click() '문장 오류 검사
On Error GoTo Quits
lstProcessing.Clear
ChkString = False
cmdTxtchk.Enabled = False
chkprPos = 0
chkErrPos = 0
chkItemPos = 0
txtProcess.Text = ""
Fn = FreeFile

Timer.Enabled = True

Open App.Path & "\TRANSLATION.XML" For Input As #Fn

Do While Not EOF(Fn)

Line Input #Fn, strTmp

chkItemPos = chkItemPos + 1
chkprPos = chkprPos + 1

labProcess.Caption = "약 " & Format((((59913 / 100) - (chkItemPos / 100))) / 60, "#") & "분 남음"

Do
    DoEvents
Loop Until Timer.Enabled = False

Timer.Enabled = True

txtProcess.Text = txtProcess.Text & strTmp & vbCrLf
strTmp = Replace(strTmp, vbTab, "")
If Not InStr(strTmp, "<String Key=""") = 0 And ChkString = False Then
    strTmp = Replace(strTmp, "<String Key=""", "")
    If InStr(strTmp, """>") = 0 Then
        lstErr.AddItem chkItemPos & "/" & """" & Replace(strTmp, """>", "") & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    Else
        lstProcessing.AddItem Replace(strTmp, """>", "")
        lstProcessing.ListIndex = lstProcessing.ListCount - 1
        ChkString = True
    End If
ElseIf Not InStr(strTmp, "</String>") = 0 Then
    If ChkString = True And chkErrPos = 5 Then
        chkErrPos = 0
        ChkString = False
    ElseIf chkErrPos < 5 Then
        lstErr.AddItem chkItemPos & "/" & """" & chkItemPos & "의 언어문장들" & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    Else
        chkErrPos = 0
        lstErr.AddItem chkItemPos & "/" & """" & chkItemPos & "줄의 String" & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    End If
ElseIf Not InStr(strTmp, "<de>") = 0 Then
    strTmp = Replace(strTmp, "<de>", "")
    If InStr(strTmp, "</de>") = 0 Then
        lstErr.AddItem chkItemPos & "/" & """" & Replace(strTmp, "</de>", "") & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    Else
        chkErrPos = chkErrPos + 1
    End If
ElseIf Not InStr(strTmp, "<fr>") = 0 Then
    strTmp = Replace(strTmp, "<fr>", "")
    If InStr(strTmp, "</fr>") = 0 Then
        lstErr.AddItem chkItemPos & "/" & """" & Replace(strTmp, "</fr>", "") & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    Else
        chkErrPos = chkErrPos + 1
    End If
ElseIf Not InStr(strTmp, "<es>") = 0 Then
    strTmp = Replace(strTmp, "<es>", "")
    If InStr(strTmp, "</es>") = 0 Then
        lstErr.AddItem chkItemPos & "/" & """" & Replace(strTmp, "</es>", "") & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    Else
        chkErrPos = chkErrPos + 1
    End If
ElseIf Not InStr(strTmp, "<it>") = 0 Then
    strTmp = Replace(strTmp, "<it>", "")
    If InStr(strTmp, "</it>") = 0 Then
        lstErr.AddItem chkItemPos & "/" & """" & Replace(strTmp, "</it>", "") & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    Else
        chkErrPos = chkErrPos + 1
    End If
ElseIf Not InStr(strTmp, "<jp>") = 0 Then
    strTmp = Replace(strTmp, "<jp>", "")
    If InStr(strTmp, "</jp>") = 0 Then
        lstErr.AddItem chkItemPos & "/" & """" & Replace(strTmp, "</jp>", "") & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    Else
        chkErrPos = chkErrPos + 1
    End If
ElseIf Not InStr(strTmp, "<!--") = 0 Then
    strTmp = Replace(strTmp, "<!--", "")
    If InStr(strTmp, "-->") = 0 Then
        lstErr.AddItem chkItemPos & "/" & """" & Replace(strTmp, "-->", "") & """"
        lstErr.ListIndex = lstErr.ListCount - 1
    End If
    Pbstatus.Value = lstProcessing.ListCount
Else
    GoTo Nex
End If
        
Nex:
    Pbstatus.Value = lstProcessing.ListCount
Loop
    
Close #Fn


cmdTxtchk.Enabled = True

Exit Sub
Quits:
End Sub

Private Sub cmdReTrans_Click()
ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = txtOrigin
txtEdit.Text = ieGoogle.document.getElementById("result_box").innertext
End Sub

Private Sub Form_Load()

On Error GoTo ErrGo

tvSetup.Nodes.Add , tvwChild, "General", "일반"
tvSetup.Nodes.Add "General", tvwChild, "Language", "번역 설정"
tvSetup.Nodes.Add "General", tvwChild, "Focus", "포커스"
tvSetup.Nodes.Add , tvwChild, "ETC", "기타"
tvSetup.Nodes.Add "ETC", tvwChild, "Translate", "번역"
tvSetup.Nodes.Item(1).Expanded = True
tvSetup.Nodes.Item(4).Expanded = True

Timer.Enabled = True

lstLanguage.Text = "jp"
StringToPrint = "Focus On"
SetRect ThisRectangle, 0, 0, FadeAnswerPictureBox.ScaleWidth, FadeAnswerPictureBox.ScaleHeight
FadeAnswerPictureBox.FontSize = 10
FrmMain.Refresh

CopyBackGroundIntoPictureBox
CopyBackgroundToMemory

Opacity = 100
PrintTranslucentText StringToPrint, Opacity
Exit Sub
ErrGo:
MsgBox "TRANSLATION.XML 파일이 없습니다.", vbCritical, "심각한 오류"
End
Exit Sub
EndFuc:
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Ends
ieGoogle.Quit
End
Exit Sub
Ends:
End
End Sub

Private Sub cmdStart_Click()
On Error GoTo Quits

If UTF8.Value = 1 Then

lstProcessing.Clear '프로세스 정리
txtProcess.Text = "" '정리
lstLanguage.Enabled = False & Label5.Caption = "0"

'#################################################################################
'<--초기화
If chkItemPos = 0 Then chkItemPos = 0 & chkcntPos = 0 & chkprPos = 0 & chklng = ""
'-->
'#################################################################################

'<--UTF-8
Set objStream = CreateObject("ADODB.Stream")
objStream.Open
objStream.Charset = "UTF-8"
objStream.Position = 0
'-->

'<--버튼 설정
cmdStart.Enabled = False
cmdStop.Enabled = True
cmdIgnore.Enabled = True
'-->
    
Timer.Enabled = True '타이머 설정
        
'Open App.Path & "\TRANSLATION.XML" For Input As #1
'Do While Not EOF(1)
'Line Input #1, strTmp

Dim vLine As Variant, sFileBody As String

'<--UTF-8
sFileBody = ConvertUTF8File(App.Path & "\TRANSLATION.XML")
sFileBody = Mid$(sFileBody, 2)
For Each vLine In Split(sFileBody, vbCrLf)
strTmp = CStr(vLine)
'-->

'<--타이머
Do
DoEvents
Loop Until Timer.Enabled = False Or Label5.Caption = "0"
Timer.Enabled = True
'-->
    
'<--기타 정보
FadeAnswerPictureBox.Cls
txtProcess.Text = txtProcess.Text & strTmp & vbCrLf
chkItemPos = chkItemPos + 1
chkprPos = chkprPos + 1
labProcess.Caption = "약 " & Format((((59913 / 100) - (chkItemPos / 100))) / 60, "#") & "분 남음"
'-->

'###
    '<--txtScript
If Not InStr(strTmp, "<String Key=") = 0 Then
    txtScript.Text = strTmp
    txtScript.Text = Replace(txtScript.Text, "<String Key=", "")
    txtScript.Text = Replace(txtScript.Text, vbTab, "")
    txtScript.Text = Replace(txtScript.Text, """", "")
    txtScript.Text = Replace(txtScript.Text, ">", "")
End If
    '-->
'###

'<--언어 확인
If Not InStr(strTmp, "<" & lstLanguage.Text & ">") = 0 Then
    chkcntPos = chkcntPos + 1
    Pbstatus.Value = chkcntPos
    strTmp = Replace(strTmp, "<" & lstLanguage.Text & ">", "")
    strTmp = Replace(strTmp, "?", "")
    strTmp = Replace(strTmp, vbTab, "")
    '<--언어 확인
    If Not InStr(strTmp, "</" & lstLanguage.Text & ">") = 0 Then
        txtEdit.Enabled = True
        strTmp = Replace(strTmp, "</" & lstLanguage.Text & ">", "")
        
        '<--포커스 맞춤
        StringToPrint = strTmp
        SetRect ThisRectangle, 0, 0, FadeAnswerPictureBox.ScaleWidth, FadeAnswerPictureBox.ScaleHeight
        FadeAnswerPictureBox.FontSize = 10
        FrmMain.Refresh '메인페이지 다시읽기
        CopyBackGroundIntoPictureBox
        CopyBackgroundToMemory
        Opacity = 100
        PrintTranslucentText StringToPrint, Opacity
        '-->
        
        Me.Caption = chkcntPos & "/큐베이스 번역기 - 병작"
        txtOrigin.Text = strTmp
'#######
        '<--클립보드 확인
        If chkClip.Value = 1 Then
            Clipboard.Clear
            Clipboard.SetText txtScript.Text
            BackUp2 = txtScript.Text
            BackUp = txtScript.Text
            Timer1.Enabled = True
            
            '<--클립보드 감시
            Do
                DoEvents
            Loop Until Not BackUp = BackUp2 Or chkClip.Value = 0
            '-->
            
            Timer1.Enabled = False
            Inputs = BackUp2
            txtEdit.Text = BackUp2
            
            cmdIgnore.Enabled = False
        End If
        '-->
'#######

'#######
            '<--그대로 유지
        If strTmp = "=" Then
            Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
            Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & strTmp & "</" & lstLanguage.Text & ">"
            Close #2
            GoTo Nex
        End If
        '-->>
'#######

        '<--구글 번역
        If Google.Value = 1 And Not chklng = "" Then

            '<--오류 판별
'###########
            '<--한영...구분
            If lan1.Text = "en" Then '영문
                
                '<--번역 내용
                ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = txtScript
                '-->
                    
            Else '기타 언어
                
                '<--번역 내용
                ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = strTmp
                '-->
                
            End If
            '-->
'###########

        cmdIgnore.Enabled = True
        
        Do
        
        '<--번역
        txtEdit.Text = ieGoogle.document.getElementById("result_box").innertext
        '-->
                DoEvents
        Loop Until (Not txtEdit.Text = "" And Not chklng = txtEdit.Text And Not txtEdit.Text = chklng & "..." And Not chklng = strTmp) Or cmdIgnore.Enabled = False
        '-->
        
        cmdIgnore.Enabled = False
        chklng = txtEdit.Text
        
        
        '<--쓰기
        Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
        Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & chklng & "</" & lstLanguage.Text & ">"
        Close #2
        '-->
        
        GoTo Nex '넘어가기
    
        ElseIf Google.Value = 1 And chklng = "" Then '최초 번역
    

'###############
            '<--한영...구분
            If lan1.Text = "en" Then '영문
                
                '<--번역 내용
                ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = txtScript
                '-->
                    
            Else '기타 언어
                
                '<--번역 내용
                ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = strTmp
                '-->
                
            End If
            '-->
'###############

        '<--오류판별
        Do
        txtEdit.Text = ieGoogle.document.getElementById("result_box").innertext
            DoEvents
        Loop Until Not txtEdit.Text = ""
        '-->
        chklng = txtEdit.Text

        '<--쓰기
        Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
        Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & chklng & "</" & lstLanguage.Text & ">"
        Close #2
        '-->
        
        GoTo Nex '넘어가기

        ElseIf Google.Value = 0 Then
    
            txtOrigin.Text = strTmp '번역단계가 아님

        End If
        '-->
'#######
        Me.Caption = txtOrigin.Text
        lstProcessing.AddItem strTmp
        lstProcessing.ListIndex = lstProcessing.ListCount - 1
        txtEdit.SetFocus
        '<--한글 감지
            
        If Korea.Value = 1 Then
            bHan = False
            bEng = False
            bJap1 = False
            bJap2 = False
            bEtc = False
            strTxt = Replace(strTmp, " ", "")
            For a = 1 To Len(strTxt)
                strTxt = ReplCase(strTxt)
            Text5.Text = Val(Mid(strTxt, a, 1))
                
            If Not Text5.Text = "0" Then
                Do
                DoEvents
                Loop Until Text5.Text = ""
            End If
            
        If Mid(strTxt, a, 1) >= "ㄱ" And Mid(strTxt, a, 1) <= "?" Then
            bHan = True
        ElseIf Mid(strTxt, a, 1) >= "a" And Mid(strTxt, a, 1) <= "z" Then
            bEng = True
        ElseIf Mid(strTxt, a, 1) >= "A" And Mid(strTxt, a, 1) <= "Z" Then
            bEng = True
        ElseIf (Mid(strTxt, a, 1) >= "あ" And Mid(strTxt, a, 1) <= "ん") Then
            bJap1 = True
        ElseIf Mid(strTxt, a, 1) >= "ア" And Mid(strTxt, a, 1) <= "ン" Then
            bJap2 = True
        ElseIf Mid(strTxt, a, 1) >= "一" And Mid(strTxt, a, 1) <= "刀" Then
            bJap1 = True
        ElseIf Mid(strTxt, a, 1) >= "一" And Mid(strTxt, a, 1) <= "魚" Then
            bJap1 = True
        ElseIf Mid(strTxt, a, 1) >= "?" And Mid(strTxt, a, 1) <= "?" Then
            bJap1 = True
        Else
            bEtc = True
        End If
                
        Next
        Label7.Caption = bHan & bEng & bEtc & bJap1 & bJap2
            
        If _
        (bHan = True And bEng = True And bEtc = True And bJap1 = False And jap2 = False) Or _
        (bHan = True And bEng = True And bEtc = False And bJap1 = False And jap2 = False) Or _
        (bHan = True And bEng = False And bEtc = True And bJap1 = False And jap2 = False) Or _
        (bHan = True And bEng = False And bEtc = False And bJap1 = False And jap2 = False) Or _
        (bHan = False And bEng = True And bEtc = False And bJap1 = False And jap2 = False) Then
            Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
            Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & strTmp & "</" & lstLanguage.Text & ">"
            Close #2
            GoTo Nex
        End If
        
    End If
    
    cmdIgnore.Enabled = True
        
    Do
        DoEvents
    Loop Until cmdIgnore.Enabled = False
    
    cmdIgnore.Enabled = False
    
    Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
    Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & Inputs & "</" & lstLanguage.Text & ">"
    Close #2

End If

Else
    Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
        Print #2, strTmp
    Close #2
            
    GoTo Nex
End If

Nex:

Next vLine

'Loop
'Close #1
Timer.Enabled = False
cmdStop_Click
Exit Sub
Quits:
ElseIf UTF8.Value = 0 Then
On Error GoTo Quits

lstProcessing.Clear '프로세스 정리
txtProcess.Text = "" '정리
lstLanguage.Enabled = False & Label5.Caption = "0"

'<--초기화
If chkItemPos = 0 Then chkItemPos = 0 & chkcntPos = 0 & chkprPos = 0 & chklng = ""
'-->

'<--버튼 설정
cmdStart.Enabled = False
cmdStop.Enabled = True
cmdIgnore.Enabled = True
'-->

Timer.Enabled = True '타이머 설정
    
Open App.Path & "\TRANSLATION.XML" For Input As #1
Do While Not EOF(1)
Line Input #1, strTmp
    
Do
DoEvents
Loop Until Timer.Enabled = False Or Label5.Caption = "0"
Timer.Enabled = True

FadeAnswerPictureBox.Cls
txtProcess.Text = txtProcess.Text & strTmp & vbCrLf
chkItemPos = chkItemPos + 1
chkprPos = chkprPos + 1
labProcess.Caption = "약 " & Format((((59913 / 100) - (chkItemPos / 100))) / 60, "#") & "분 남음"
    
    '<--txtScript
    If Not InStr(strTmp, "<String Key=") = 0 Then
        txtScript.Text = strTmp
        txtScript.Text = Replace(txtScript.Text, "<String Key=", "")
        txtScript.Text = Replace(txtScript.Text, vbTab, "")
        txtScript.Text = Replace(txtScript.Text, """", "")
        txtScript.Text = Replace(txtScript.Text, ">", "")
    End If
    '-->
    
    '<--언어 확인
    If Not InStr(strTmp, "<" & lstLanguage.Text & ">") = 0 Then
        chkcntPos = chkcntPos + 1
        Pbstatus.Value = chkcntPos
        strTmp = Replace(strTmp, "<" & lstLanguage.Text & ">", "")
        strTmp = Replace(strTmp, "?", "")
        strTmp = Replace(strTmp, vbTab, "")
        '<--언어 확인
        If Not InStr(strTmp, "</" & lstLanguage.Text & ">") = 0 Then
            txtEdit.Enabled = True
            strTmp = Replace(strTmp, "</" & lstLanguage.Text & ">", "")
            '<--포커스 맞춤
            StringToPrint = strTmp
            SetRect ThisRectangle, 0, 0, FadeAnswerPictureBox.ScaleWidth, FadeAnswerPictureBox.ScaleHeight
            FadeAnswerPictureBox.FontSize = 10
            FrmMain.Refresh '메인페이지 다시읽기
            CopyBackGroundIntoPictureBox
            CopyBackgroundToMemory
            Opacity = 100
            PrintTranslucentText StringToPrint, Opacity
            '-->
            
            Me.Caption = chkcntPos & "/큐베이스 번역기 - 병작"
            
            txtOrigin.Text = strTmp
            
            '<--클립보드 확인
            If chkClip.Value = 1 Then
            Clipboard.Clear
            Clipboard.SetText txtScript.Text
            BackUp2 = txtScript.Text
            BackUp = txtScript.Text
            Timer1.Enabled = True
            
            Do
                DoEvents
            Loop Until Not BackUp = BackUp2 Or chkClip.Value = 0
            
            Timer1.Enabled = False
            Inputs = BackUp2
            txtEdit.Text = BackUp2
            
            cmdIgnore.Enabled = False
            End If
            '-->
            
                '<--그대로 유지
            If strTmp = "=" Then
                Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
                Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & strTmp & "</" & lstLanguage.Text & ">"
                Close #2
                GoTo Nex2
            End If
            '-->>
            
            '<--구글 번역
            If Google.Value = 1 And Not chklng = "" Then

                '<--오류 판별
                
                '<--한영...구분
                If lan1.Text = "en" Then '영문
                    
                    '<--번역 내용
                    ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = txtScript
                    '-->
                        
                Else '기타 언어
                    
                    '<--번역 내용
                    ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = strTmp
                    '-->
                    
                End If
                '-->
            cmdIgnore.Enabled = True
            
            Do
            
            '<--번역
            txtEdit.Text = ieGoogle.document.getElementById("result_box").innertext
            '-->
                    DoEvents
            Loop Until (Not txtEdit.Text = "" And Not chklng = txtEdit.Text And Not txtEdit.Text = chklng & "..." And Not chklng = strTmp) Or cmdIgnore.Enabled = False
            '-->
            
            chklng = txtEdit.Text
            
            
            '<--쓰기
            Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
            Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & chklng & "</" & lstLanguage.Text & ">"
            Close #2
            '-->
            
            GoTo Nex2 '넘어가기
            
        ElseIf Google.Value = 1 And chklng = "" Then '최초 번역
        
                '<--오류 판별
                
                '<--한영...구분
                If lan1.Text = "en" Then '영문
                    
                    '<--번역 내용
                    ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = txtScript
                    '-->
                        
                Else '기타 언어
                    
                    '<--번역 내용
                    ieGoogle.document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = strTmp
                    '-->
                    
                End If
                '-->
                
            '<--오류판별
            Do
            txtEdit.Text = ieGoogle.document.getElementById("result_box").innertext
                DoEvents
            Loop Until Not txtEdit.Text = ""
            '-->
            chklng = txtEdit.Text

            '<--쓰기
            Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
            Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & chklng & "</" & lstLanguage.Text & ">"
            Close #2
            '-->
            
            GoTo Nex2 '넘어가기
        
        ElseIf Google.Value = 0 Then
        
            txtOrigin.Text = strTmp '번역단계가 아님

        End If
        '-->

    Me.Caption = txtOrigin.Text
    lstProcessing.AddItem strTmp
    lstProcessing.ListIndex = lstProcessing.ListCount - 1
    txtEdit.SetFocus
    bHan = False
    bEng = False
    bJap1 = False
    bJap2 = False
    bEtc = False
    strTxt = Replace(strTmp, " ", "")
    For a = 1 To Len(strTxt)
        strTxt = ReplCase(strTxt)
    Text5.Text = Val(Mid(strTxt, a, 1))
        
        If Not Text5.Text = "0" Then
            Do
            DoEvents
            Loop Until Text5.Text = ""
        End If
        
        If Mid(strTxt, a, 1) >= "ㄱ" And Mid(strTxt, a, 1) <= "?" Then
            bHan = True
        ElseIf Mid(strTxt, a, 1) >= "a" And Mid(strTxt, a, 1) <= "z" Then
            bEng = True
        ElseIf Mid(strTxt, a, 1) >= "A" And Mid(strTxt, a, 1) <= "Z" Then
            bEng = True
        ElseIf (Mid(strTxt, a, 1) >= "あ" And Mid(strTxt, a, 1) <= "ん") Then
            bJap1 = True
        ElseIf Mid(strTxt, a, 1) >= "ア" And Mid(strTxt, a, 1) <= "ン" Then
            bJap2 = True
        ElseIf Mid(strTxt, a, 1) >= "一" And Mid(strTxt, a, 1) <= "刀" Then
            bJap1 = True
        ElseIf Mid(strTxt, a, 1) >= "一" And Mid(strTxt, a, 1) <= "魚" Then
            bJap1 = True
        ElseIf Mid(strTxt, a, 1) >= "?" And Mid(strTxt, a, 1) <= "?" Then
            bJap1 = True
        Else
            bEtc = True
        End If
        Next
        Label7.Caption = bHan & bEng & bEtc & bJap1 & bJap2
        
        If _
        (bHan = True And bEng = True And bEtc = True And bJap1 = False And jap2 = False) Or _
        (bHan = True And bEng = True And bEtc = False And bJap1 = False And jap2 = False) Or _
        (bHan = True And bEng = False And bEtc = True And bJap1 = False And jap2 = False) Or _
        (bHan = True And bEng = False And bEtc = False And bJap1 = False And jap2 = False) Or _
        (bHan = False And bEng = True And bEtc = False And bJap1 = False And jap2 = False) Then
            Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
            Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & strTmp & "</" & lstLanguage.Text & ">"
            Close #2
            GoTo Nex2
        End If
    
    Do
        DoEvents
    Loop Until cmdIgnore.Enabled = False
    
    cmdIgnore.Enabled = True
    Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
    Print #2, vbTab & vbTab & vbTab & "<" & lstLanguage.Text & ">" & Inputs & "</" & lstLanguage.Text & ">"
    Close #2

    End If
Else
    Open App.Path & "\TRANSLATION(TRANS).XML" For Append Access Write As #2
        Print #2, strTmp
    Close #2
            
    GoTo Nex2
End If
Nex2:

Loop
Close #1
Timer.Enabled = False
cmdStop_Click
Exit Sub
Quits2:
End If
End Sub

Private Sub cmdStop_Click()
txtEdit.Enabled = False
cmdIgnore.Enabled = False
Label5.Caption = "1"
Timer.Enabled = False
cmdStop.Enabled = False
cmdStart.Enabled = True
End Sub

Private Sub Google_Click()
If Google.Value = 1 Then
    Set ieGoogle = CreateObject("InternetExplorer.Application")
      
    With ieGoogle
        .Silent = True
        .Navigate "http://translate.google.com/"
        .Visible = False
    End With
    
    If ieReady Then
        Google.Visible = True
    End If
Else
    ieGoogle.Quit
End If
End Sub

Private Sub tvSetup_NodeClick(ByVal Node As MSComctlLib.Node)
Select Case Node.Index
    Case 2
        frmSetting.Visible = True
        frmFocus.Visible = False
        frmTranslate.Visible = False
    Case 3
        frmSetting.Visible = False
        frmFocus.Visible = True
        frmTranslate.Visible = False
        StringToPrint = "Focus On"
        SetRect ThisRectangle, 0, 0, FadeAnswerPictureBox.ScaleWidth, FadeAnswerPictureBox.ScaleHeight
        FadeAnswerPictureBox.FontSize = 30
        FrmMain.Refresh
        CopyBackGroundIntoPictureBox
        CopyBackgroundToMemory
        Opacity = 100
        PrintTranslucentText StringToPrint, Opacity
    Case 5
        frmSetting.Visible = False
        frmFocus.Visible = False
        frmTranslate.Visible = True
End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtEdit.Text = "" Then
        cmdIgnore_Click
        cmdIgnore.Enabled = False
    Else
        Inputs = txtEdit.Text
        txtEdit.Enabled = False
        cmdIgnore.Enabled = False
        txtEdit.Text = ""
    End If
End If
End Sub

Private Sub Timer_Timer()
If chkprPos > 8 Then
    txtProcess.Text = ""
    chkprPos = 0
End If
Timer.Enabled = False
End Sub

Private Sub Timer1_Timer()
BackUp2 = Clipboard.GetText
End Sub

Private Sub VScroll1_Change()
txtSpeed.Text = VScroll1.Value
Timer.Interval = VScroll1.Value
End Sub

Private Sub lan1_Click()
On Error GoTo errlng
  ieGoogle.document.getElementById("gt-submit").Click
  ieGoogle.document.getElementById("gt-sl").Value = GetLanguage(lan1)
  ieGoogle.document.getElementById("gt-submit").Click
  ieGoogle.document.Forms(0).submit
Exit Sub
errlng:
  lan1.Text = ""
End Sub

Private Sub lan2_Click()
On Error GoTo errlng
  ieGoogle.document.getElementById("gt-submit").Click
  ieGoogle.document.getElementById("gt-tl").Value = GetLanguage(lan2)
  ieGoogle.document.getElementById("gt-submit").Click
  ieGoogle.document.Forms(0).submit
Exit Sub
errlng:
  lan2.Text = ""
End Sub

Public Function GetLanguage(sLanguage As String) As String
Select Case sLanguage
  Case "en"
    GetLanguage = "en"
  Case "fr"
    GetLanguage = "fr"
  Case "de"
    GetLanguage = "de"
  Case "it"
    GetLanguage = "it"
  Case "jp"
    GetLanguage = "ja"
  Case "ko"
    GetLanguage = "ko"
  Case "es"
    GetLanguage = "es"
End Select
End Function

Public Function ieReady() As Boolean
Dim ie_Ready As Long
Dim doc_Ready As String

  ie_Ready = 4
  doc_Ready = "complete"
 
  Do Until ieGoogle.readyState = ie_Ready
    DoEvents
  Loop
  
  Do Until ieGoogle.document.readyState = doc_Ready
    DoEvents
  Loop
  
  ieReady = True
  Exit Function
  
ErrExit:
  ieReady = False
End Function


Public Sub CopyBackGroundIntoPictureBox()
Dim ThisWidth   As Integer
Dim ThisHeight  As Integer
Dim XCoord      As Integer
Dim YCoord      As Integer
XCoord = FadeAnswerPictureBox.Left
YCoord = FadeAnswerPictureBox.Top
ThisWidth = FadeAnswerPictureBox.ScaleWidth
ThisHeight = FadeAnswerPictureBox.ScaleHeight
FadeAnswerPictureBox.Visible = False
BitBlt FadeAnswerPictureBox.hdc, 0, 0, ThisWidth, ThisHeight, FrmMain.hdc, XCoord, YCoord, vbSrcCopy
FadeAnswerPictureBox.Visible = True
End Sub

Public Sub PrintTranslucentText(ByVal ThisText As String, ThisOpacity As Integer)
CopyBackgroundFromMemory
FadeAnswerPictureBox.ForeColor = RGB(129, 0, 0)
DrawText FadeAnswerPictureBox.hdc, StringToPrint, Len(StringToPrint), ThisRectangle, DT_WORDBREAK
AlphaBlendWithBackground (ThisOpacity)
FadeAnswerPictureBox.Refresh
End Sub

Public Sub AlphaBlendWithBackground(ByVal BlendValue As Integer)
Dim ThisWidth   As Integer
Dim ThisHeight  As Integer
BF.BlendOp = AC_SRC_OVER
BF.BlendFlags = 0
BF.SourceConstantAlpha = 255 - BlendValue
BF.AlphaFormat = 0
RtlMoveMemory lBF, BF, 4
ThisWidth = FadeAnswerPictureBox.ScaleWidth
ThisHeight = FadeAnswerPictureBox.ScaleHeight
AlphaBlend FadeAnswerPictureBox.hdc, 0, 0, ThisWidth, ThisHeight, BackGroundDC, 0, 0, ThisWidth, ThisHeight, lBF
End Sub

Public Sub CopyBackgroundFromMemory()
SetDIBits FadeAnswerPictureBox.hdc, FadeAnswerPictureBox.Image, 0, FadeAnswerPictureBox.ScaleHeight, Pixels(1, 1, 1), BackgroundBitmap, DIB_RGB_COLORS
FadeAnswerPictureBox.Picture = FadeAnswerPictureBox.Image
End Sub

Public Sub CopyBackgroundToMemory()
Dim ThisWidth   As Integer
Dim ThisHeight  As Integer
Dim XCoord      As Integer
Dim YCoord      As Integer
Dim Bytes_per_scanLine As Integer
Dim X, Y As Integer
XCoord = FadeAnswerPictureBox.Left
YCoord = FadeAnswerPictureBox.Top
ThisWidth = FadeAnswerPictureBox.ScaleWidth
ThisHeight = FadeAnswerPictureBox.ScaleHeight
With BackgroundBitmap.Header
    .biSize = 40
    .biWidth = ThisWidth
    .biHeight = -ThisHeight
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    Bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    .biSizeImage = Bytes_per_scanLine * Abs(.biHeight)
End With
ReDim Pixels(1 To 4, 1 To FadeAnswerPictureBox.ScaleWidth, 1 To FadeAnswerPictureBox.ScaleHeight)
BackGroundDC = CreateCompatibleDC(0)
iBitmap = CreateDIBSection(BackGroundDC, BackgroundBitmap, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
SelectObject BackGroundDC, iBitmap
GetDIBits FadeAnswerPictureBox.hdc, FadeAnswerPictureBox.Image, 0, FadeAnswerPictureBox.ScaleHeight, Pixels(1, 1, 1), BackgroundBitmap, DIB_RGB_COLORS
SetDIBits BackGroundDC, iBitmap, 0, FadeAnswerPictureBox.ScaleHeight, Pixels(1, 1, 1), BackgroundBitmap, DIB_RGB_COLORS
End Sub

Private Sub LoadXML(ByRef oNode As MSXML2.IXMLDOMNode, Optional ByVal sParent As String = "")
    Dim intCount As Integer, intCount2 As Integer
    Dim objNode As MSXML2.IXMLDOMNode
    Dim strPath As String
    Dim objNodeList As MSXML2.IXMLDOMNodeList
    Dim strName As String
    
    Timer.Enabled = True
    
    If oNode Is Nothing Then Exit Sub
    
    If sParent = "" Then
        sParent = "*[1]"
        strCurrentKey = "*[1]"
        Text1.Text = Text1.Text & oNode.xml
    End If
       
    For intCount = 0 To oNode.childNodes.Length - 1
    
    Do
        DoEvents
    Loop Until Timer.Enabled = False
    Timer.Enabled = True
    
        Set objNode = oNode.childNodes(intCount)
        If InStr(1, Exclude, objNode.nodeName & ";", vbTextCompare) = 0 Then
            strPath = sParent & "/*[" & intCount + 1 & "]"
            strName = objNode.nodeName
            
            If Len(strTitleAttributes) > 0 Then
                Set objNodeList = objNode.selectNodes("@*[contains('" & strTitleAttributes & "',name())]")
                If objNodeList.Length > 0 Then
                    strName = strName & " ("
                    For intCount2 = 0 To objNodeList.Length - 1
                        strName = strName & objNodeList.Item(intCount2).nodeValue
                        If intCount < objNodeList.Length - 1 Then
                            strName = strName & ","
                        End If
                    Next
                    strName = strName & ")"
                End If
                Set objNodeList = Nothing
            End If
            
            Text2.Text = Text2.Text & strName

        
            If objNode.hasChildNodes Then
                'Call LoadTree(objNode, strPath)
            End If
        End If
    Next
End Sub
