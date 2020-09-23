VERSION 5.00
Begin VB.Form FrmClose 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Close or Minimize"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Minimize To System Tray"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close Program"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "What Would You Like To Do?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "FrmClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call FrmMain.CleanUpSystray
Call FrmMain.SaveSettings
End
End Sub

Private Sub Command2_Click()
Call FrmMain.SaveSettings
FrmMain.Hide
DoEvents
Unload Me
End Sub
