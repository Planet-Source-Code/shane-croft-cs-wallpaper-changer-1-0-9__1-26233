VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CS Wallpaper Changer"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6165
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10742
            Text            =   "© Crofts Software - Networking Software & More"
            TextSave        =   "© Crofts Software - Networking Software & More"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "8/15/2001"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:26 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Advanced Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   3495
      Begin VB.CheckBox Check3 
         Caption         =   "Minimize To System Tray At Startup"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Launch At Startup"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2880
         Picture         =   "FrmMain.frx":1D2A
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Wallpaper Options"
      Height          =   1815
      Left            =   128
      TabIndex        =   6
      Top             =   2880
      Width           =   5280
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   4800
         Top             =   1320
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Don't Change Wallpaper"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   720
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Change Wallpaper At Startup"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Random Order"
         Height          =   210
         Left            =   2640
         TabIndex        =   18
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable Wallpaper Preview"
         Height          =   210
         Left            =   2640
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Change Wallpaper Hourly"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Change Wallpaper Daily"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Change Wallpaper Weekly"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Change Wallpaper Monthly"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Wallpaper List"
      Height          =   2775
      Left            =   128
      TabIndex        =   3
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Command3 
         Caption         =   "Set as Wallpaper"
         Height          =   255
         Left            =   7200
         TabIndex        =   19
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   120
         Top             =   240
      End
      Begin VB.ListBox List1 
         Height          =   2160
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Total Wallpapers in list: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   8535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   3255
      Left            =   5408
      TabIndex        =   0
      Top             =   2880
      Width           =   3495
      Begin VB.PictureBox picScreen 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         Height          =   2940
         Left            =   120
         ScaleHeight     =   192
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   216
         TabIndex        =   1
         Top             =   240
         Width           =   3300
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Date && Time Wallpaper Will Change."
      Height          =   615
      Left            =   3720
      TabIndex        =   24
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label LblTime 
      Alignment       =   2  'Center
      Caption         =   "Time"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      ToolTipText     =   "Date & Time To Change The Wallpaper"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label LblDate 
      Alignment       =   2  'Center
      Caption         =   "Date"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      ToolTipText     =   "Date & Time To Change The Wallpaper"
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu MenuShow 
         Caption         =   "&Show/Hide"
      End
      Begin VB.Menu MenuChange 
         Caption         =   "&Change Wallpaper"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "&Exit Program"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 32
End Type

'constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256
Private nid As NOTIFYICONDATA

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1 ' Unicode nul terminated String
    Const REG_DWORD = 4 ' 32-bit number


Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Private Declare Function ClipCursor Lib "user32" _
    (lpRect As Any) As Long

Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private FilePathName As String
Public Sub CleanUpSystray()
Shell_NotifyIcon NIM_DELETE, nid
End Sub
Public Sub SaveSettings()
Dim fFile As Integer
fFile = FreeFile
'save settings
Open App.Path & "\Settings.inf" For Output As fFile
Print #fFile, "[settings]"
Print #fFile, "monthly=" & Option1.Value
Print #fFile, "weekly=" & Option2.Value
Print #fFile, "daily=" & Option3.Value
Print #fFile, "hourly=" & Option4.Value
Print #fFile, "startup=" & Option5.Value
Print #fFile, "nochange=" & Option6.Value
Print #fFile, "preview=" & Check1.Value
Print #fFile, "random=" & Check4.Value
Print #fFile, "launchatstartup=" & Check2.Value
Print #fFile, "minimize=" & Check3.Value
Print #fFile, "lastpic=" & List1.ListIndex
Print #fFile, "sdate=" & LblDate.Caption
Print #fFile, "stime=" & LblTime.Caption
Close fFile
DoEvents
End Sub
Public Sub List_Add(list As ListBox, txt As String)
On Error Resume Next
    List1.AddItem txt
End Sub
Public Sub List_Load(thelist As ListBox, FileName As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call List_Add(List1, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub

Public Sub List_Save(thelist As ListBox, FileName As String)
    'Save a listbox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List1.list(Save)
    Next Save
    Close fFile
End Sub
Private Sub SetPicture(ByVal FileName As String)
    On Error GoTo Dawm
    Dim xFile As String
    xFile = WinPath & "CS WallPaper.bmp"

    SavePicture picScreen.Picture, xFile
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, ByVal xFile, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
Dawm:
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Call AddToRun("CS Wallpaper Changer", App.Path & "\" & App.EXEName & ".exe")
End If
If Check2.Value = 0 Then
Call RemoveFromRun("CS Wallpaper Changer")
End If
End Sub

Private Sub Command1_Click()
On Error GoTo dangit
CD1.Filter = "Supported Picture Files|*.jpg;*.bmp;*.gif"
CD1.CancelError = True
CD1.ShowOpen
List1.AddItem CD1.FileName
Exit Sub
dangit:

End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
DoEvents
End Sub

Private Sub Command3_Click()
SetPicture List1.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim AppDir As String
AppDir = App.Path

Me.Caption = "CS Wallpaper Changer v" & App.Major & "." & App.Minor & "." & App.Revision
Call List_Load(List1, App.Path & "\WallpaperList.ini")
DoEvents

FilePathName = AppDir + "\Settings.inf"
monthly = GetPrivateProfileString("settings", "monthly", "", FilePathName)
weekly = GetPrivateProfileString("settings", "weekly", "", FilePathName)
daily = GetPrivateProfileString("settings", "daily", "", FilePathName)
hourly = GetPrivateProfileString("settings", "hourly", "", FilePathName)
startup = GetPrivateProfileString("settings", "startup", "", FilePathName)
nochange = GetPrivateProfileString("settings", "nochange", "", FilePathName)
preview = GetPrivateProfileString("settings", "preview", "", FilePathName)
random = GetPrivateProfileString("settings", "random", "", FilePathName)
launchatstartup = GetPrivateProfileString("settings", "launchatstartup", "", FilePathName)
minimize = GetPrivateProfileString("settings", "minimize", "", FilePathName)
lastpic = GetPrivateProfileString("settings", "lastpic", "", FilePathName)
sdate = GetPrivateProfileString("settings", "sdate", "", FilePathName)
stime = GetPrivateProfileString("settings", "stime", "", FilePathName)

DoEvents
Option1.Value = monthly
Option2.Value = weekly
Option3.Value = daily
Option4.Value = hourly
Option5.Value = startup
Option6.Value = nochange
Check1.Value = preview
Check2.Value = launchatstartup
Check3.Value = minimize
Check4.Value = random
List1.ListIndex = lastpic
LblDate.Caption = sdate
LblTime.Caption = stime
DoEvents
DoEvents

If Check3.Value = 1 Then
Me.Hide
End If

With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
nid.szTip = "CS Wallpaper Changer v" & App.Major & "." & App.Minor & "." & App.Revision
End With
Shell_NotifyIcon NIM_ADD, nid
Timer2.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
    
Select Case msg
Case WM_LBUTTONDBLCLK    '515 restore form window
If Me.Visible = True Then
Me.Visible = False
Else
Me.Visible = True
Me.SetFocus
End If
Case WM_RBUTTONUP        '517 display popup menu
Me.PopupMenu Me.MenuFile
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
Cancel = True
End If
Call List_Save(List1, App.Path & "\WallpaperList.ini")
DoEvents
FrmClose.Show
End Sub

Private Sub List1_Click()
On Error GoTo dangit
If Check1.Value = 1 Then
    If 1 = 2 Then
        picScreen.Cls
        picScreen.PaintPicture LoadPicture(List1.Text), 0, 0, picScreen.ScaleWidth, picScreen.ScaleHeight
    Else
        Set picScreen.Picture = LoadPicture(List1.Text)
    End If
    picPreview.PaintPicture picScreen.Picture, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight
    picPreview.Refresh
End If
Exit Sub
dangit:
    If 1 = 2 Then
        picScreen.Cls
        picScreen.PaintPicture LoadPicture(App.Path & "\Error.jpg"), 0, 0, picScreen.ScaleWidth, picScreen.ScaleHeight
    Else
        Set picScreen.Picture = LoadPicture(App.Path & "\Error.jpg")
    End If
    picPreview.PaintPicture picScreen.Picture, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight
    picPreview.Refresh
End Sub

Private Sub MenuChange_Click()
On Error Resume Next
    If Check4.Value = 1 Then
    List1.ListIndex = Int(Rnd * List1.ListCount)
    Else
    If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
    Else
    List1.ListIndex = List1.ListIndex + 1
    End If
    End If
DoEvents
Call Command3_Click
End Sub

Private Sub MenuExit_Click()
Call List_Save(List1, App.Path & "\WallpaperList.ini")
DoEvents
Call FrmMain.SaveSettings
Call FrmMain.CleanUpSystray
End
End Sub

Private Sub MenuShow_Click()
If Me.Visible = True Then
Me.Visible = False
Else
Me.Visible = True
Me.SetFocus
End If
End Sub

Private Sub Option1_Click()
Timer2.Enabled = True
LblDate = Date + 30
LblTime = "Time - Never"
End Sub

Private Sub Option2_Click()
Timer2.Enabled = True
LblDate = Date + 7
LblTime = "Time - Never"
End Sub

Private Sub Option3_Click()
Timer2.Enabled = True
LblDate = Date + 1
LblTime = "Time - Never"
End Sub

Private Sub Option4_Click()
Timer2.Enabled = True
LblDate.Caption = "Date - Never"
LblTime.Caption = DateAdd("h", 1, Time)
End Sub

Private Sub Option5_Click()
Timer2.Enabled = False
LblDate.Caption = "Program Startup"
LblTime.Caption = "Program Startup"
End Sub

Private Sub Option6_Click()
Timer2.Enabled = False
LblDate.Caption = "Date - Never"
LblTime.Caption = "Time - Never"
End Sub

Private Sub Timer1_Timer()
If List1.ListCount = 0 Then
Command3.Enabled = False
Else
Command3.Enabled = True
End If
Me.Label1.Caption = "Total Wallpapers in list: " & List1.ListCount
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Dim xxx As Date
Dim yyy As Date
xxx = LblDate.Caption
yyy = LblTime.Caption

If Option1.Value = True Then
Timer2.Enabled = True
If xxx <= Date Then
    If Check4.Value = 1 Then
    List1.ListIndex = Int(Rnd * List1.ListCount)
    Else
    If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
    Else
    List1.ListIndex = List1.ListIndex + 1
    End If
    End If
LblDate = Date + 30
DoEvents
Call Command3_Click
End If
    
End If

If Option2.Value = True Then
Timer2.Enabled = True
If xxx <= Date Then
    If Check4.Value = 1 Then
    List1.ListIndex = Int(Rnd * List1.ListCount)
    Else
    If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
    Else
    List1.ListIndex = List1.ListIndex + 1
    End If
    End If
LblDate = Date + 7
DoEvents
Call Command3_Click
End If
End If

If Option3.Value = True Then
Timer2.Enabled = True
If xxx <= Date Then
    If Check4.Value = 1 Then
    List1.ListIndex = Int(Rnd * List1.ListCount)
    Else
    If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
    Else
    List1.ListIndex = List1.ListIndex + 1
    End If
    End If
LblDate = Date + 1
DoEvents
Call Command3_Click
End If
End If

If Option4.Value = True Then
Timer2.Enabled = True
If yyy <= Time Then
    If Check4.Value = 1 Then
    List1.ListIndex = Int(Rnd * List1.ListCount)
    Else
    If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
    Else
    List1.ListIndex = List1.ListIndex + 1
    End If
    End If
LblTime.Caption = DateAdd("h", 1, Time)
DoEvents
Call Command3_Click
End If
End If

If Option5.Value = True Then
Timer2.Enabled = False
    If Check4.Value = 1 Then
    List1.ListIndex = Int(Rnd * List1.ListCount)
    Else
    If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
    Else
    List1.ListIndex = List1.ListIndex + 1
    End If
    End If
Call Command3_Click
End If
End Sub
Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***

   Dim szTmp                     As String
   Dim nRet                      As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)

End Function
Private Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   ' *** Get an entry in the WIN inifile ***

   Dim szTmp                    As String
   Dim nRet                     As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)

End Function
Public Sub AddToRun(ProgramName As String, FileToRun As String)
    'Add a program to the 'Run at Startup' r
    '     egistry keys
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun)
End Sub


Public Sub RemoveFromRun(ProgramName As String)
    'Remove a program from the 'Run at Start
    '     up' registry keys
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName)
End Sub
Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub


Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal hKey As HKeyTypes, ByVal strPath As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry")
    '
    Dim keyhand As Long
    r = RegDeleteKey(hKey, strPath)
End Function

