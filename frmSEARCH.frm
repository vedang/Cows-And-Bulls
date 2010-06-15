VERSION 5.00
Begin VB.Form frmSEARCH 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Search The Database"
   ClientHeight    =   1830
   ClientLeft      =   360
   ClientTop       =   2595
   ClientWidth     =   3810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmSEARCH.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   1830
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtWORD 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Text            =   "Enter the Word"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Frame frmSRCH 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   4200
      TabIndex        =   0
      Top             =   4080
      Width           =   3375
      Begin VB.Label lblOK 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   600
         MouseIcon       =   "frmSEARCH.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1560
         Width           =   2115
      End
   End
   Begin VB.Image Image4 
      Height          =   3060
      Left            =   8880
      Picture         =   "frmSEARCH.frx":1954
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   3060
      Left            =   960
      Picture         =   "frmSEARCH.frx":122FF
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
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
      Left            =   8400
      MouseIcon       =   "frmSEARCH.frx":165B6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   7680
      Width           =   2805
   End
   Begin VB.Image bull 
      Height          =   1920
      Left            =   9720
      Picture         =   "frmSEARCH.frx":17C00
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image cow 
      Height          =   2040
      Left            =   720
      Picture         =   "frmSEARCH.frx":184CA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1905
   End
   Begin VB.Label newgame 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SEARCH "
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
      TabIndex        =   2
      Top             =   480
      Width           =   12015
   End
End
Attribute VB_Name = "frmSEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******* FORM : SEARCH THE DATABSE *****

Public search_flag As Byte
Public search_result As String
Public call_from As Byte
Dim flag_txt As Integer

Private Sub back_Click()
    frmGAMECHOICE.Visible = True
    Unload Me
End Sub
'-----------------------------------------------------

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 26
    back.ForeColor = &HFFFF&
End Sub
'-----------------------------------------------------

Private Sub Form_Click()
    txtWORD.ForeColor = &H808080
    txtWORD.FontBold = False
    txtWORD.FontSize = 14
    txtWORD.Text = "Enter the Word"
    flag_txt = 0
End Sub
'-----------------------------------------------------

Private Sub frmSRCH_Click()
    flag_txt = 0
End Sub
'-----------------------------------------------------

Private Sub frmSRCH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOK.FontSize = 26
    lblOK.ForeColor = &HFF00&
End Sub
'-----------------------------------------------------

'CHECKS FOR THE WOORD IN THE DATABASE
Public Sub lblOK_Click()
Dim i As Byte
Dim j As Byte
Dim arr(1 To 4) As Byte
Dim str As String
Dim connection As New ADODB.connection
Dim recordset As New ADODB.recordset
   
search_flag = 0
connection.Open "dsn=adodc", "scott", "tiger"
recordset.CursorLocation = adUseClient
recordset.CursorType = adOpenKeyset
recordset.Open "select * from tp_word", connection
    
    'CHECK LENGTH
    If Len(txtWORD.Text) <> 4 Then
        MsgBox "rtfm"
        txtWORD.SetFocus
        txtWORD = ""
        Exit Sub
    End If

For i = 1 To 4
    arr(i) = CByte(AscB(Mid(txtWORD.Text, i, 1)))
Next i

'REPEATED LETTERS ARE NOT ALLOWED
For i = 1 To 3
    For j = i + 1 To 4
        If arr(i) = arr(j) Then
            MsgBox "rtfm"
            txtWORD.SetFocus
            txtWORD = ""
            Exit Sub
        End If
    Next j
Next i

'ACTUAL CHCKING
recordset.MoveFirst
Do While Not recordset.EOF
    str = recordset!word
    If txtWORD.Text = str Then
        search_flag = 1
        search_result = txtWORD.Text
        Exit Do
    End If
    recordset.MoveNext
Loop
recordset.Close
connection.Close

'CALLED FROM THE COMPUTER GUESSES UTILITY
If call_from = 1 Then
    frmSEARCH.Visible = False
    frmCOMP.Visible = True
    Call frmCOMP.search_answers
    Unload frmCOMP
    frmSEARCH.Visible = False
    frmINTRO.Visible = True
    
'USED AS AN INDEPENDANT FUNCTIONALITY
Else
    If search_flag = 1 Then
        MsgBox "Word IS in the database. Go Ahead.. Enjoy the game..."
    Else
        MsgBox "BooHoo... My Conceited Makers didn't deem it necessary to add this word to the Database!!"
    End If
    flag_txt = 0
    frmGAMECHOICE.Visible = True
    Unload Me
End If

End Sub
'-----------------------------------------------------

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 22
    back.ForeColor = &HC0C0&
    lblOK.FontSize = 26
    lblOK.ForeColor = &HFF00&
End Sub
'-----------------------------------------------------

Private Sub lblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOK.FontSize = 30
    lblOK.ForeColor = &HFFFF&
End Sub
'-----------------------------------------------------

Private Sub txtWORD_Click()
    txtWORD.Text = ""
    flag_txt = 1
End Sub
'-----------------------------------------------------

'THIS SUB ALLOWS ONLY ALLOWED CHARACTERS
Private Sub txtWORD_KeyPress(KeyAscii As Integer)
    If flag_txt = 0 Then
        txtWORD.Text = Left(txtWORD.Text, 0)
        flag_txt = 1
    End If
    txtWORD.ForeColor = &H80000007
    txtWORD.FontBold = True
    txtWORD.FontSize = 16

If KeyAscii = 8 Then
    Exit Sub
End If

If KeyAscii = 13 Then
    Call lblOK_Click 'cmdOK_Click
    Exit Sub
End If

If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
        KeyAscii = 0
End If

If KeyAscii >= 65 And KeyAscii <= 90 Then
    KeyAscii = KeyAscii + 32
End If

End Sub
'-----------------------------------------------------
