VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmUSER 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "You Guess"
   ClientHeight    =   7080
   ClientLeft      =   2355
   ClientTop       =   5970
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmUSER.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGIVEUP 
      Caption         =   "&Give up!! :("
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdBACK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "<<--- &Back"
      Height          =   495
      Left            =   3840
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdGUESS 
      Caption         =   "Guess"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid fgridVAL 
      Height          =   2655
      Left            =   4200
      TabIndex        =   1
      ToolTipText     =   "Previously entered words"
      Top             =   4920
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   11
      Cols            =   4
      FixedCols       =   0
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   0
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.TextBox txtINPUT 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   570
      Left            =   4320
      TabIndex        =   0
      Text            =   "Start guessing"
      Top             =   3120
      Width           =   3495
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label lblQUIT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quit ???"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   960
      MouseIcon       =   "frmUSER.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7800
      Width           =   2265
   End
   Begin VB.Image bull 
      Height          =   1680
      Left            =   10560
      Picture         =   "frmUSER.frx":1954
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1440
   End
   Begin VB.Image cow 
      Height          =   2040
      Left            =   0
      Picture         =   "frmUSER.frx":221E
      Stretch         =   -1  'True
      Top             =   480
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
      TabIndex        =   7
      Top             =   480
      Width           =   12075
   End
   Begin VB.Label lblguess 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guess"
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
      Left            =   4680
      MouseIcon       =   "frmUSER.frx":2528
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3840
      Width           =   2655
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
      Left            =   8760
      MouseIcon       =   "frmUSER.frx":3B72
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   7800
      Width           =   2565
   End
   Begin VB.Image Image4 
      Height          =   3060
      Left            =   9120
      Picture         =   "frmUSER.frx":51BC
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   3060
      Left            =   1080
      Picture         =   "frmUSER.frx":15B67
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   2055
   End
End
Attribute VB_Name = "frmUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******* FORM : USER GUESSES *******

'THIS MODULE IS DESIGNED FOR THE USER GUESSES FUNCTIONALITY
'OF THE GAME

Dim num_of_entries As Byte   'KEEPS TRACK OF NUMBER OF ENTRIES MADE
Dim str_to_be_guessed As String  'THE WORD HELD BY THE COMP
Public connection As ADODB.connection
Public recordset As ADODB.recordset
Dim rec As ADODB.recordset
Public auth_flag As Byte
Dim flag_txt As Integer
'-------------------------------------------------------

'THIS MODULE UPDATES THE USER LOG AND EXITS
Private Sub update_user_info(ByVal str_result As String, ByVal num_of_chances As Byte)
Dim no_of_wins As Integer

recordset.Close
recordset.LockType = adLockPessimistic
recordset.Open "select * from userinfo", connection


    'CASE : USER WINS
    If str_result = "win" Then
        Do While Not recordset.EOF
            If recordset!Name = frmUSERLOGIN.struser Then
                recordset!login_time = frmUSERLOGIN.user_time
                no_of_wins = recordset!win_perc * recordset!play_count / 100
                no_of_wins = no_of_wins + 1
                recordset!play_count = recordset!play_count + 1
                recordset.Update
                recordset!win_perc = no_of_wins * 100 / recordset!play_count
                If Left(recordset!best_effort, 2) = "NA" Then
                    recordset!best_effort = CStr(num_of_chances)
                Else
                    If Val(recordset!best_effort) > num_of_chances Then
                       recordset!best_effort = CStr(num_of_chances)
                    End If
                End If
               recordset.Update
               GoTo save
            End If
            recordset.MoveNext
        Loop
        MsgBox "User not found !! (first time users require a restart for the changes to take effect)"
        GoTo save
        
    'CASE : USER LOSES
    Else
        Do While Not recordset.EOF
            If recordset!Name = frmUSERLOGIN.struser Then
                recordset!login_time = frmUSERLOGIN.user_time
                no_of_wins = recordset!win_perc * recordset!play_count / 100
                recordset!play_count = recordset!play_count + 1
                recordset.Update
                recordset!win_perc = no_of_wins * 100 / recordset!play_count
                recordset.Update
                GoTo save
            End If
            recordset.MoveNext
        Loop
        MsgBox "User not found !! (first time users require a restart for the changes to take effect)"
        GoTo save
    End If
save:
recordset.Close
connection.Close
frmGAMECHOICE.Visible = True
Unload Me
End Sub
'-------------------------------------------------------


'THE DATABASE IS MANUALLY DESIGNED SO, IT IS POSSIBLE THT
'SOME WORDS MAY BE MISSING
'THIS MODULE AUTHENTICATES THE USER AND ADDS THE NEW WORD
'TO THE DBASE
Public Sub addnewword()

Set rec = New ADODB.recordset

rec.CursorType = adOpenKeyset
rec.LockType = adLockPessimistic
rec.Open "select * from tp_word", connection

    'LOGIN FAILED
    If frmLogin.LoginSucceeded = False Then
                'USER TRIES TO HACK !!
                If auth_flag = 2 Then
                    MsgBox "Don't Tell me I didn't warn you..."
                    End
                End If
                
                'UNAUTHORIZED USER
                If auth_flag = 1 Then
                    MsgBox "Try hacking again and you'll b kicked out !"
                                    
                End If
                
    'USER AUTHORIZED
    Else
            Dim num As Integer
            rec.AddNew
            rec!word = txtINPUT.Text
            rec!freq = 50
            rec.Update
    End If
    
rec.Close
recordset.Requery

'NOW THE NEW WORD IS ADDED
'SO SIMULATE THE GUESS CLICK
If frmLogin.LoginSucceeded = True Then
    Call lblGUESS_Click
    If num_of_entries >= 10 Then
        Exit Sub
    End If
End If
    
    txtINPUT.Text = ""
    txtINPUT.SetFocus
        
End Sub
'-----------------------------------------------------
                
'SHOW LOGIN BOX
Private Sub authenticate()
    frmUSER.Visible = False
    frmLogin.Visible = True
    frmLogin.SetFocus
End Sub
'-----------------------------------------------------

'THIS SUB SELECTS A RANDOM WORD TO BE GUESSED BY THE USER
Public Function Randomer(str As String)
Dim num As Integer

If frmGAMECHOICE.difficulty = 1 Then
    num = Int((228 - 1) * Rnd) + 1
Else
   num = Int((frmUSER.recordset.RecordCount - 228 - 1) * Rnd) + 228 + 1
End If

frmUSER.recordset.Move num - frmUSER.recordset.Bookmark
str = recordset!word
End Function
'-----------------------------------------------------

'THIS SUB CALCULATES THE NUMBER OF COWZ AND BULLZ
Private Function Cowz_Bullz(test_str As String, num_cowz As Byte, num_bullz As Byte) As String
    
    Dim ch_to_be_guessed As Byte
    Dim ch_test As Byte
    Dim i As Byte
    Dim j As Byte
    
    num_cowz = 0
    num_bullz = 0
    
    For i = 1 To 4
        ch_test = CByte(AscB(Mid(test_str, i, 1)))
        For j = 1 To 4
            ch_to_be_guessed = CByte(AscB(Mid(str_to_be_guessed, j, 1)))
            If i <> j And ch_test = ch_to_be_guessed Then
                num_cowz = num_cowz + 1
            ElseIf i = j And ch_test = ch_to_be_guessed Then
                num_bullz = num_bullz + 1
            End If
        Next j
    Next i
    
End Function
'-----------------------------------------------------

Private Sub back_Click()
Dim flag As VbMsgBoxResult
    flag = MsgBox("Game in progress. Give up?", vbYesNo)
    If flag = vbYes Then
        MMControl1.Command = "close"
        Call update_user_info("lose", 0)
        frmGAMECHOICE.Visible = True
        Unload Me
    End If
End Sub
'-----------------------------------------------------

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 26
    back.ForeColor = &HFFFF&
End Sub
'-----------------------------------------------------

Private Sub cmdGIVEUP_Click()
    Dim flag As VbMsgBoxResult
    flag = MsgBox("Are You Sure You Want To Quit? ", vbYesNo)
    If flag = vbYes Then
        Call update_user_info("lose", 0)
        frmINTRO.Visible = True
        Unload Me
    End If
End Sub
'-----------------------------------------------------

'THIS SUB RECORDS THE ENTRY MADE BY THE USER
'AND CHECKS FOR POSSIBLE ERRORS
Private Sub lblGUESS_Click()

Dim arr(1 To 4) As Byte
Dim str As String
Dim i As Byte
Dim j As Byte
Dim num_cowz As Byte
Dim num_bullz As Byte
Dim flag As Byte
Dim rep_flag As Byte
Dim temp As String
Dim decision As VbMsgBoxResult

'LENGTH OF WORD
If Len(txtINPUT.Text) <> 4 Then
        MsgBox "rtfm (read the fantastic manual)"
        txtINPUT.SetFocus
        txtINPUT = ""
        Exit Sub
End If

    
For i = 1 To 4
    arr(i) = CByte(AscB(Mid(txtINPUT.Text, i, 1)))
Next i

'CHECK FOR REPETITION
For i = 1 To 3
    For j = i + 1 To 4
        If arr(i) = arr(j) Then
            MsgBox "rtfm (read the fantastic manual)"
            txtINPUT.SetFocus
            txtINPUT = ""
            Exit Sub
        End If
    Next j
Next i
                 
'ACCESS THE DBASE AND CHECK FOR EXISTENCE OF THE WORD
recordset.MoveFirst
Do While Not recordset.EOF
    str = recordset!word
    If txtINPUT.Text = str Then
        flag = 1
        Exit Do
    End If
    recordset.MoveNext
Loop


'IS THE WORD IN THE DBASE ?
If flag = 0 Then
     decision = MsgBox("Do you wanna add thid new word to database?", vbYesNo)
        If decision = vbYes Then
            temp = txtINPUT.Text
            Call authenticate
            Exit Sub
        Else
            txtINPUT.Text = ""
            txtINPUT.SetFocus
            Exit Sub
        End If
End If


'NO MORE GUESSES ALLOWED !!
If num_of_entries > 10 Then
    MsgBox "Can't you see ?? NO more space on the FlexGrid"
    Exit Sub
End If

'CHECK FOR REPETITION
For i = 1 To num_of_entries
    If txtINPUT.Text = fgridVAL.TextMatrix(i, 1) Then
        MsgBox "please consult the FLEXGRID!!!!"
        txtINPUT = ""
        Exit Sub
    End If
Next i


'HUSSSSSH ... THE WORD IS ALLOWED !!
insert num_of_entries, 0, CStr(num_of_entries)
insert num_of_entries, 1, txtINPUT.Text
Cowz_Bullz txtINPUT.Text, num_cowz, num_bullz
insert num_of_entries, 2, CStr(num_cowz)
insert num_of_entries, 3, CStr(num_bullz)


'USER GUESSES THE WORD SUCCESSFULLY !!
If num_bullz = 4 Then
    MsgBox "OK! You saved the World !!"
    Call update_user_info("win", num_of_entries)
    frmGAMECHOICE.Visible = True
    Unload Me
    Exit Sub
End If

'USER LOSES
If num_of_entries = 10 Then
    MsgBox "You are Hopeless!! The word is " & str_to_be_guessed
    Call update_user_info("lose", 0)
    Exit Sub
End If

num_of_entries = num_of_entries + 1
txtINPUT.SetFocus
txtINPUT.Text = ""
    
        
End Sub
'-----------------------------------------------------

Private Sub Form_Click()
    txtINPUT.ForeColor = &H80000004
    txtINPUT.FontBold = False
    txtINPUT.FontSize = 18
    txtINPUT.Text = "Start guessing"
    flag_txt = 0
End Sub
'-----------------------------------------------------

Private Sub Form_Load()

    flag_txt = 0
    
    MMControl1.FileName = "c:\click.wav"
    MMControl1.Command = "Open"
    
    insert 0, 0, "Number"
    insert 0, 1, "Words"
    insert 0, 2, "Cowz"
    insert 0, 3, "Bullz"
    
    Set frmUSER.connection = New ADODB.connection
    Set frmUSER.recordset = New ADODB.recordset
       
    With frmUSER.recordset
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
    End With
    
    frmUSER.connection.Open "dsn=adodc", "scott", "tiger"
    frmUSER.recordset.Open "select * from tp_word", connection
        
    Randomize
    num_of_entries = 1
    Randomer str_to_be_guessed
End Sub
'-----------------------------------------------------

'INSERT IN FLEX-GRID
Private Sub insert(row As Byte, col As Byte, str As String)
    With fgridVAL
        .row = row
        .col = col
        .Text = str
    End With
End Sub
'-----------------------------------------------------

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 22
    back.ForeColor = &HC0C0&
    lblQUIT.FontSize = 22
    lblQUIT.ForeColor = &HFF&
    lblguess.FontSize = 26
    lblguess.ForeColor = &HFF00&
    MMControl1.Command = "prev"
End Sub
'-----------------------------------------------------

Private Sub Form_Unload(Cancel As Integer)
If recordset.State = adStateOpen Then recordset.Close
If connection.State = adStateOpen Then connection.Close
End Sub
'-----------------------------------------------------

Private Sub lblguess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblguess.FontSize = 30
    lblguess.ForeColor = &HFF&
    MMControl1.Command = "play"
End Sub
'-----------------------------------------------------

Private Sub lblQUIT_Click()
    Dim flag As VbMsgBoxResult
    flag = MsgBox("Are You Sure You Want To Quit? ", vbYesNo)
    If flag = vbYes Then
        Call update_user_info("lose", 0)
        frmINTRO.Visible = True
        Unload Me
    End If
End Sub
'-----------------------------------------------------

Private Sub lblQUIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblQUIT.FontSize = 26
    lblQUIT.ForeColor = &HFFFF&
End Sub
'-----------------------------------------------------

Private Sub txtINPUT_Click()
    txtINPUT.Text = ""
    flag_txt = 1
End Sub
'-----------------------------------------------------

Private Sub txtINPUT_KeyPress(KeyAscii As Integer)
    If flag_txt = 0 Then
        txtINPUT.Text = Left(txtINPUT.Text, 0)
        flag_txt = 1
    End If
    txtINPUT.ForeColor = &H80000007
    txtINPUT.FontBold = True
    txtINPUT.FontSize = 20

If KeyAscii = 8 Then
    Exit Sub
End If

If KeyAscii = 13 Then
    Call lblGUESS_Click 'cmdGUESS_Click
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

' END OF USER GUESSES
'=====================================================
