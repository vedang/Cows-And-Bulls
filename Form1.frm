VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmINTRO 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "COWZ and BULLZ..."
   ClientHeight    =   6390
   ClientLeft      =   195
   ClientTop       =   1365
   ClientWidth     =   7410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6390
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TimTIMER 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4200
      Top             =   4320
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image Image3 
      Height          =   3060
      Left            =   8640
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Image Image4 
      Height          =   3060
      Left            =   1680
      Picture         =   "Form1.frx":10CB5
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Image ico4 
      Height          =   480
      Left            =   4200
      Picture         =   "Form1.frx":14F6C
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico3 
      Height          =   480
      Left            =   4200
      Picture         =   "Form1.frx":15276
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico2 
      Height          =   480
      Left            =   4200
      Picture         =   "Form1.frx":15580
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico1 
      Height          =   480
      Left            =   4200
      Picture         =   "Form1.frx":1588A
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label exit1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   5790
      MouseIcon       =   "Form1.frx":15B94
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label fmanual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Manual"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   4800
      MouseIcon       =   "Form1.frx":171DE
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   7080
      Width           =   3075
   End
   Begin VB.Label options 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Logs"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   795
      Left            =   5160
      MouseIcon       =   "Form1.frx":18828
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6120
      Width           =   2595
   End
   Begin VB.Label newgame 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   5040
      MouseIcon       =   "Form1.frx":19E72
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5160
      Width           =   2835
   End
   Begin VB.Image Image2 
      Height          =   1320
      Left            =   9720
      Picture         =   "Form1.frx":1B4BC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   600
      Picture         =   "Form1.frx":1BD86
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1440
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cows && Bulls"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Image bull 
      Height          =   3825
      Left            =   6360
      Picture         =   "Form1.frx":1C090
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   5640
   End
   Begin VB.Image cow 
      Height          =   3825
      Left            =   0
      Picture         =   "Form1.frx":2DB1F
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   6345
   End
End
Attribute VB_Name = "frmINTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********** FORM : INTRO FORM ********


Dim flag As Integer
Dim flagtim As Integer
Dim ch As Integer
'-----------------------------------------------------

Private Sub fmanual_Click()
    MMControl1.Command = "close"
    MMControl1.Notify = True
    MMControl1.FileName = "c:\ball2ball.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "play"
    ch = 3
    TimTIMER.Enabled = True
End Sub
'-----------------------------------------------------

Private Sub fmanual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    newgame.ForeColor = &HFF00&
    options.ForeColor = &HFF00&
    fmanual.ForeColor = &HFF&
    exit1.ForeColor = &HFF00&
    fmanual.FontSize = 26
    ico1.Visible = False
    ico2.Visible = False
    ico3.Visible = True
    ico4.Visible = False
    If flag = 0 Then
        MMControl1.Command = "play"
        flag = 1
    End If
End Sub
'-----------------------------------------------------

Private Sub exit1_Click()
    MMControl1.Command = "close"
    MMControl1.Notify = True
    MMControl1.FileName = "c:\ball2ball.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "play"
    ch = 4
    TimTIMER.Enabled = True
End Sub
'-----------------------------------------------------

Private Sub exit1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    newgame.ForeColor = &HFF00&
    options.ForeColor = &HFF00&
    fmanual.ForeColor = &HFF00&
    exit1.ForeColor = &HFF&
    exit1.FontSize = 26
    ico1.Visible = False
    ico2.Visible = False
    ico3.Visible = False
    ico4.Visible = True
    If flag = 0 Then
        MMControl1.Command = "play"
        flag = 1
    End If
End Sub
'-----------------------------------------------------


Private Sub Form_Load()
    MMControl1.FileName = "c:\click.wav"
    MMControl1.Command = "Open"
End Sub
'-----------------------------------------------------


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    newgame.ForeColor = &HFF00&
    options.ForeColor = &HFF00&
    fmanual.ForeColor = &HFF00&
    exit1.ForeColor = &HFF00&
    newgame.FontSize = 24
    options.FontSize = 24
    fmanual.FontSize = 24
    exit1.FontSize = 24
    ico1.Visible = False
    ico2.Visible = False
    ico3.Visible = False
    ico4.Visible = False
    flag = 0
    MMControl1.Command = "prev"
End Sub
'-----------------------------------------------------

Private Sub newgame_Click()
    MMControl1.Command = "close"
    MMControl1.Notify = True
    MMControl1.FileName = "c:\ball2ball.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "play"
    ch = 1
    TimTIMER.Enabled = True
End Sub
'-----------------------------------------------------

Private Sub newgame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    newgame.ForeColor = &HFF&
    options.ForeColor = &HFF00&
    fmanual.ForeColor = &HFF00&
    exit1.ForeColor = &HFF00&
    newgame.FontSize = 26
    ico1.Visible = True
    ico2.Visible = False
    ico3.Visible = False
    ico4.Visible = False
    If flag = 0 Then
        MMControl1.Command = "play"
        flag = 1
    End If
End Sub
'-----------------------------------------------------

Private Sub options_Click()
    MMControl1.Command = "close"
    MMControl1.Notify = True
    MMControl1.FileName = "c:\ball2ball.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "play"
    ch = 2
    TimTIMER.Enabled = True
End Sub
'-----------------------------------------------------

Private Sub options_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    newgame.ForeColor = &HFF00&
    options.ForeColor = &HFF&
    fmanual.ForeColor = &HFF00&
    exit1.ForeColor = &HFF00&
    options.FontSize = 26
    ico1.Visible = False
    ico2.Visible = True
    ico3.Visible = False
    ico4.Visible = False
    If flag = 0 Then
        MMControl1.Command = "play"
        flag = 1
    End If
End Sub
'-----------------------------------------------------

Private Sub TimTIMER_Timer()
    Select Case ch
    
    Case 1:    'GAMECHOICE FORM
        TimTIMER.Enabled = False
        frmGAMECHOICE.Visible = True
        Unload Me
    Case 2:   'USER INFO
        TimTIMER.Enabled = False
        drptUSERLOG.Visible = True
        Unload Me
    Case 3:   'MANUAL
        TimTIMER.Enabled = False
        frmMANUAL.Visible = True
        Unload Me
    Case 4:    'EXIT
        TimTIMER.Enabled = False
        MMControl1.Command = "close"
        End
        
    End Select
End Sub

'-----------------------------------------------------
