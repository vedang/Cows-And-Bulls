VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmGAMECHOICE 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "What Form Of the Game?"
   ClientHeight    =   5805
   ClientLeft      =   840
   ClientTop       =   6195
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5805
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimTIMER 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   3120
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.OptionButton optDIFFICULTY 
      BackColor       =   &H80000007&
      Caption         =   "Difficult"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Index           =   1
      Left            =   4800
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.OptionButton optDIFFICULTY 
      BackColor       =   &H80000007&
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   0
      Left            =   4800
      TabIndex        =   5
      Top             =   3720
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.Image ico4 
      Height          =   480
      Left            =   2760
      Picture         =   "Form2.frx":030A
      Stretch         =   -1  'True
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblWRDDBASE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word Database"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3360
      MouseIcon       =   "Form2.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   7440
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Form2.frx":1C5E
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico3 
      Height          =   480
      Left            =   2760
      Picture         =   "Form2.frx":1F68
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label srch 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "Search Utility "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3720
      MouseIcon       =   "Form2.frx":2272
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6360
      Width           =   3855
   End
   Begin VB.Label back 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<<---  Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   600
      Left            =   8640
      MouseIcon       =   "Form2.frx":38BC
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   7800
      Width           =   2685
   End
   Begin VB.Image Image4 
      Height          =   3060
      Left            =   9000
      Picture         =   "Form2.frx":4F06
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   3060
      Left            =   240
      Picture         =   "Form2.frx":158B1
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Image ico2 
      Height          =   480
      Left            =   2760
      Picture         =   "Form2.frx":19B68
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ico1 
      Height          =   480
      Left            =   2760
      Picture         =   "Form2.frx":19E72
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label comp_user 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I'll Guess....."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3840
      MouseIcon       =   "Form2.frx":1A17C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Label user_comp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Guess..... "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3840
      MouseIcon       =   "Form2.frx":1B7C6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2880
      Width           =   3915
   End
   Begin VB.Image bull 
      Height          =   1680
      Left            =   10560
      Picture         =   "Form2.frx":1CE10
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1440
   End
   Begin VB.Image cow 
      Height          =   2040
      Left            =   0
      Picture         =   "Form2.frx":1D6DA
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1905
   End
   Begin VB.Label newgame 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NEW GAME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   12075
   End
End
Attribute VB_Name = "frmGAMECHOICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******* FORM : SELECT GAME TYPE *******

Public difficulty As Byte
Dim ch As Integer
'-----------------------------------------------------

Private Sub back_Click()
    MMControl1.Command = "close"
    frmINTRO.Visible = True
    Unload Me
End Sub
'-----------------------------------------------------

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 26
    back.ForeColor = &HFFFF&
End Sub
'-----------------------------------------------------

Private Sub comp_user_Click()
    MMControl1.Command = "close"
    MMControl1.Notify = True
    MMControl1.FileName = "c:\ball2ball.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "play"
    ch = 2
    TimTIMER.Enabled = True
    
End Sub
'-----------------------------------------------------

Private Sub comp_user_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    user_comp.ForeColor = &HFF&
    comp_user.ForeColor = &HFF00&
    comp_user.FontSize = 28
    srch.ForeColor = &HFF&
    lblWRDDBASE.ForeColor = &HFF&
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

Private Sub Form_Load()
    MMControl1.FileName = "c:\click.wav"
    MMControl1.Command = "Open"
End Sub
'-----------------------------------------------------

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    user_comp.ForeColor = &HFF&
    comp_user.ForeColor = &HFF&
    srch.ForeColor = &HFF&
    lblWRDDBASE.ForeColor = &HFF&
    user_comp.FontSize = 26
    comp_user.FontSize = 26
    srch.FontSize = 26
    lblWRDDBASE.FontSize = 26
    back.FontSize = 22
    back.ForeColor = &HC0C0&
    ico1.Visible = False
    ico2.Visible = False
    ico3.Visible = False
    ico4.Visible = False
    flag = 0
    MMControl1.Command = "prev"
End Sub
'-----------------------------------------------------

Private Sub lblWRDDBASE_Click()
    MMControl1.Command = "close"
    MMControl1.Notify = True
    MMControl1.FileName = "c:\ball2ball.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "play"
    ch = 4
    TimTIMER.Enabled = True
    
End Sub
'-----------------------------------------------------

Private Sub lblWRDDBASE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    user_comp.ForeColor = &HFF&
    comp_user.ForeColor = &HFF&
    srch.ForeColor = &HFF&
    lblWRDDBASE.ForeColor = &HFF00&
    lblWRDDBASE.FontSize = 28
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


Private Sub srch_Click()
    frmSEARCH.call_from = 0
    MMControl1.Command = "close"
    MMControl1.Notify = True
    MMControl1.FileName = "c:\ball2ball.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "play"
    ch = 3
    TimTIMER.Enabled = True
End Sub
'-----------------------------------------------------

Private Sub srch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    user_comp.ForeColor = &HFF&
    comp_user.ForeColor = &HFF&
    srch.ForeColor = &HFF00&
    srch.FontSize = 28
    lblWRDDBASE.ForeColor = &HFF&
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

Private Sub user_comp_Click()
    If optDIFFICULTY(0).Value = True Then
        difficulty = 0
    Else
        difficulty = 1
    End If
    MMControl1.Command = "close"
    MMControl1.Notify = True
    MMControl1.FileName = "c:\ball2ball.wav"
    MMControl1.Command = "Open"
    MMControl1.Command = "play"
    ch = 1
    TimTIMER.Enabled = True
End Sub
'-----------------------------------------------------

Private Sub user_comp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    user_comp.ForeColor = &HFF00&
    comp_user.ForeColor = &HFF&
    user_comp.FontSize = 28
    srch.ForeColor = &HFF&
    lblWRDDBASE.ForeColor = &HFF&
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

Private Sub TimTIMER_Timer()
    Select Case ch
    
    Case 1:   'USER GUESSES
        TimTIMER.Enabled = False
        frmUSER.Visible = True
        Unload Me
    Case 2:   'COMPUTER GUESSES
        TimTIMER.Enabled = False
        frmCOMP.Visible = True
        Unload Me
    Case 3:   'SEARCH UTILITY
        TimTIMER.Enabled = False
        frmSEARCH.Visible = True
        Unload Me
    Case 4:   'WORD DATABABSE
        TimTIMER.Enabled = False
        dprtWORD_DB.Visible = True
        Unload Me
    
    End Select
End Sub
'-----------------------------------------------------
