VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCOMP 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "I'll Guess"
   ClientHeight    =   6090
   ClientLeft      =   240
   ClientTop       =   1485
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCOMP.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6090
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBINGO 
      Caption         =   "BING&O!!!!!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   8040
      Width           =   3135
   End
   Begin VB.ComboBox cmbBULLZ 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmCOMP.frx":030A
      Left            =   7200
      List            =   "frmCOMP.frx":031D
      MousePointer    =   1  'Arrow
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Enter Number of Bulls here"
      Top             =   6600
      Width           =   495
   End
   Begin VB.ComboBox cmbCOWZ 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmCOMP.frx":0330
      Left            =   6120
      List            =   "frmCOMP.frx":0343
      MousePointer    =   1  'Arrow
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Enter Number of Cows Here"
      Top             =   6600
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid fgridVAL2 
      Height          =   2655
      Left            =   4200
      TabIndex        =   0
      Top             =   3120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   11
      Cols            =   4
      FixedCols       =   0
      GridColor       =   -2147483637
      ScrollBars      =   0
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label lblCOMMIT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Commit..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   600
      Left            =   5070
      MouseIcon       =   "frmCOMP.frx":0356
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   7200
      Width           =   2115
   End
   Begin VB.Label lblQUIT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   525
      Left            =   1080
      MouseIcon       =   "frmCOMP.frx":19A0
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   7320
      Width           =   2025
   End
   Begin VB.Image Image3 
      Height          =   3060
      Left            =   1080
      Picture         =   "frmCOMP.frx":2FEA
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Image Image4 
      Height          =   3060
      Left            =   8880
      Picture         =   "frmCOMP.frx":72A1
      Stretch         =   -1  'True
      Top             =   3480
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
      Left            =   8520
      MouseIcon       =   "frmCOMP.frx":17C4C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7320
      Width           =   2565
   End
   Begin VB.Image bull 
      Height          =   1680
      Left            =   10560
      Picture         =   "frmCOMP.frx":19296
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1440
   End
   Begin VB.Image cow 
      Height          =   2040
      Left            =   0
      Picture         =   "frmCOMP.frx":19B60
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
      TabIndex        =   8
      Top             =   600
      Width           =   12075
   End
   Begin VB.Label lblBULLZ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bulls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   7080
      TabIndex        =   7
      Top             =   6000
      Width           =   705
   End
   Begin VB.Label lblCOWZ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   6000
      TabIndex        =   6
      Top             =   6000
      Width           =   780
   End
   Begin VB.Label lblCAPTION 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Word"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   4200
      TabIndex        =   4
      Top             =   6000
      Width           =   1290
   End
   Begin VB.Label lblGUESS 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
End
Attribute VB_Name = "frmCOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************* FORM : COMPUTER GUESSES **********

'THIS FORM IS FOR THE COMPUTER GUESSES FUNCTIONALITY
'THIS MODULE SIMULATES THE HUMAN THINKING

Option Explicit
Dim rec2 As New ADODB.recordset
Dim con2 As New ADODB.connection

'FLAGS FOR POSITIONALLY DISCARDING ALPHABETS
Dim discard(1 To 4, 1 To 26) As Boolean

'FLAGS FOR PROBABLE ALPHABETS
Dim probable(1 To 26) As Boolean
Dim no_of_guess As Byte
Dim cnum As Byte  'NO OF COWS
Dim bnum As Byte  'NO OF BULLS
Dim word_num As Byte 'SHORTLISTED WORD COUNT
Dim guess_word(1 To 3, 1 To 3) As String  'HARDCODED SETS OF WORDS
Dim total_letters As Byte
Dim rec As New ADODB.recordset
Dim string_db(1 To 150) As String   'SHORTLISTED WORDS
Dim str_count As Byte
Dim try As Byte   'FLAG WHICH DECIDES WHETHER TO ACCESS DBASE OR STRING_DB
Dim word_frequency(1 To 150) As Byte

'------------------------------------------------------
'THIS SUB INITIALIZES ALL DEFINED VARIABLES
Private Sub init_all()
    Dim i As Byte
    Dim j As Byte
    
    For i = 1 To 4
        For j = 1 To 26
            discard(i, j) = False
            probable(j) = False
        Next j
    Next i
    
    For i = 1 To 150
        string_db(j) = ""
    Next i
    
    try = 0
    cnum = 0
    bnum = 0
    total_letters = 0
        
End Sub
'------------------------------------------------------

'THIS SUB VERIFIES THE ENTRIES MADE BY THE USER IN CASE THE COMPUTER GOES OUT OF GUESSES
Private Sub Verify_search()
    Dim ch_to_be_guessed As Byte
    Dim ch_test As Byte
    Dim i As Byte
    Dim j As Byte
    Dim row_num As Byte
    Dim num_cowz As Byte
    Dim num_bullz As Byte
    Dim search_string As String
    
    'WORD TO BE SEARCHED
    search_string = frmSEARCH.search_result
     
For row_num = 1 To fgridVAL2.row - 1
    num_cowz = 0
    num_bullz = 0
    For i = 1 To 4
        ch_test = CByte(AscB(Mid(fgridVAL2.TextMatrix(row_num, 1), i, 1)))
        For j = 1 To 4
            ch_to_be_guessed = CByte(AscB(Mid(search_string, j, 1)))
            If i <> j And ch_test = ch_to_be_guessed Then
                num_cowz = num_cowz + 1
            ElseIf i = j And ch_test = ch_to_be_guessed Then
                num_bullz = num_bullz + 1
            End If
        Next j
    Next i
    If num_cowz <> Int(fgridVAL2.TextMatrix(row_num, 2)) Or num_bullz <> Int(fgridVAL2.TextMatrix(row_num, 3)) Then
    'MISTAKE FOUND!!
        MsgBox "You have made a mistake!!"
        MsgBox fgridVAL2.TextMatrix(row_num, 1) & " " & num_cowz & "c " & num_bullz & "b"
        Exit Sub
    End If
Next row_num
        
MsgBox "You Win!! Is there a bug in the code?? Report To Jitubhai or Vedang..."

End Sub
'------------------------------------------------------

'THIS SUB VERIFIES WHETHER THE ENTERED WORD IS IN THE DBASE OR NOT
Public Sub search_answers()
    If frmSEARCH.search_flag = 1 Then
        MsgBox "Word IS in the database! Verifying Entries..."
        Call Verify_search
    Else
        MsgBox "Verified that Word is not in database..."
        frmINTRO.Visible = True
        Unload Me
    End If
End Sub
'------------------------------------------------------

'THIS SUB EXTRACTS THE MAXIMUM FREQUENCY WORD. THIS IS THE NEXT WORD TO BE GUESSED
Private Function maximum(count As Byte) As Byte
    
    Dim i As Byte
    Dim max As Byte
    max = 1
    
    For i = 1 To count - 1
        If word_frequency(i) > word_frequency(max) Then
            max = i
        End If
    Next i
    maximum = max
End Function
'------------------------------------------------------

'THIS SUB RETURNS MAX FREQ WORD
Private Function sort(str_count As Byte) As String
    
    Dim total As Byte
    Dim max As Byte
    Dim i As Byte
    Dim j As Byte
    Dim query As String
    Dim char As Byte
    Dim pos As Byte
    
    pos = 1
    
    rec.CursorLocation = adUseClient
    rec.CursorType = adOpenKeyset
        
    For i = 1 To str_count - 1
        total = 0
        
        For j = 1 To 4
            char = Asc(Mid(string_db(i), j, 1))
            
            If j = 1 Then
                query = "select pos1 from letter_freq where chr='" & chr(char) & "'"
                rec.Open query, con2
                rec.Requery
                'POS 1
                total = total + rec!pos1
                rec.Close
            ElseIf j = 2 Then
                query = "select pos2 from letter_freq where chr='" & chr(char) & "'"
                rec.Open query, con2
                rec.Requery
                'POS 2
                total = total + rec!pos2
                rec.Close
            ElseIf j = 3 Then
                query = "select pos3 from letter_freq where chr='" & chr(char) & "'"
                rec.Open query, con2
                rec.Requery
                'POS 3
                total = total + rec!pos3
                rec.Close
            ElseIf j = 4 Then
                query = "select pos4 from letter_freq where chr='" & chr(char) & "'"
                rec.Open query, con2
                rec.Requery
                'POS 4
                total = total + rec!pos4
                rec.Close
            End If
        Next j
        
        word_frequency(i) = total
        
        If total > max Then
            max = total
            pos = i
        End If
    
    Next i
    
    sort = string_db(pos)

End Function

'-------------------------------------------------------

'INSERTS WORD INTO FLEXGRID
Private Sub insert(row As Byte, col As Byte, ByVal str As String)
    With fgridVAL2
        .row = row
        .col = col
        .Text = str
    End With
End Sub
'-------------------------------------------------------

'CHECKS THE PRESENT WORD AGANIST THE PREVIOUSLY MADE ENTRIES
Private Sub Letter_Check(flag As Byte, str As String)

Dim i As Byte
Dim temp As String
Dim chr As String
Dim total As Byte
Dim j As Byte
  
For i = 1 To fgridVAL2.row - 1
    total = 0
    temp = fgridVAL2.TextMatrix(i, 1)
    For j = 1 To 4
        chr = Mid(str, j, 1)
        If InStr(1, temp, chr) <> 0 Then
            total = total + 1
        End If
    Next j
    
    If total <> Int(fgridVAL2.TextMatrix(i, 2)) + _
    Int(fgridVAL2.TextMatrix(i, 3)) Then
          flag = 1
        Exit Sub
    End If
Next i
    
End Sub
'-------------------------------------------------------

'ACCESSES THE DABSE AND SHORTLISTS PROBABLE WORDS
Private Sub filter_database(str As String)

Dim i As Byte
Dim chr As Byte
Dim count As Byte

Dim flag As Byte

str_count = 1

With rec2
    .MoveFirst
    Do While Not .EOF
        Dim tempstr As String
        tempstr = !word
        count = 0
        For i = 1 To 4
        'CHECK THE WORD WITH PREVIOUSLY ENTERED VALUES
            chr = Asc(Mid(tempstr, i, 1))
            chr = chr - 96
            
            If probable(chr) = True And discard(i, chr) = False Then
                count = count + 1
            End If
        Next i
        
               
        If count = total_letters Then
            flag = 0
            For i = 1 To fgridVAL2.row - 1
                If tempstr = fgridVAL2.TextMatrix(i, 1) Then
                    flag = 1
                End If
            Next i
            
            Call Letter_Check(flag, tempstr)
            
            If flag = 0 Then
            'THE WORD IS PROBABLE
                string_db(str_count) = tempstr
                str_count = str_count + 1
            End If
        End If
        .MoveNext
    Loop
End With

'RETURNS THE WORD
str = sort(str_count)

End Sub
'-------------------------------------------------------

'ACCESSES THE LIST OF SHORT-LISTED WORDS AND FURHTER SHORT-LISTS THEM
Private Function filter_array(str As String)

Dim i As Byte
Dim chr As Byte
Dim count As Byte
Dim new_count As Byte
Dim flag As Byte
Dim j As Byte

new_count = 1
For j = 1 To str_count - 1
    flag = 0
    For i = 1 To fgridVAL2.row - 1
        If string_db(j) = fgridVAL2.TextMatrix(i, 1) Then
            flag = 1
        End If
    Next i
            
    'CHECKS AGANIST PREV CONDITIONS
    Call Letter_Check(flag, string_db(j))
            
    If flag = 0 Then
        string_db(new_count) = string_db(j)
        word_frequency(new_count) = word_frequency(j)
        new_count = new_count + 1
    End If
Next j

str_count = new_count
If str_count = 1 Then
    MsgBox "Word is not in the database"
    frmSEARCH.call_from = 1
    frmCOMP.Visible = False
    frmSEARCH.Visible = True
Else
     'GUESS WORD
     str = string_db(maximum(str_count))
End If
    
End Function
'-------------------------------------------------------

'THIS IS THE SUB WHICH ACTUALLY IMPOSES THE CONDITIONS THAT THE
'OTHER WORDS MUST CONFORM WITH
'THIS SUB IS THE HEART OF THE CODE
Private Sub Upgrade_Tables()

Dim i As Byte
Dim j As Byte
Dim chr As Byte

'CASE : COWS = 0 AND BULLS <> 0
If cnum = 0 And bnum <> 0 Then
    For i = 1 To 4
        
        chr = Asc(Mid(lblguess.Caption, i, 1))
        chr = chr - 96
        For j = 1 To 4
            If i <> j Then
                discard(j, chr) = True
            End If
        Next j
        probable(chr) = True
   Next i

'CASE : COWS <> 0 AND BULLS  = 0
ElseIf cnum <> 0 And bnum = 0 Then
   For i = 1 To 4
    
    chr = Asc(Mid(lblguess.Caption, i, 1))
    chr = chr - 96
    For j = 1 To 4
        If i = j Then
          discard(j, chr) = True
        End If
    Next j
    probable(chr) = True
   Next i


'CASE : COWS = 0 AND BULLS = 0
ElseIf cnum = 0 And bnum = 0 Then
   For i = 1 To 4
    
    chr = Asc(Mid(lblguess.Caption, i, 1))
    chr = chr - 96
    For j = 1 To 4
        discard(j, chr) = True
    Next j
   Next i

'CASE : COWS <> 0 AND BULLS <> 0
ElseIf cnum <> 0 And bnum <> 0 Then
    For i = 1 To 4
    
    chr = Asc(Mid(lblguess.Caption, i, 1))
    chr = chr - 96
    For j = 1 To 4
        probable(chr) = True
    Next j
   Next i

End If

End Sub
'-------------------------------------------------------

'THIS IS THE THINKING PART OF THE CODE
'THIS FUNCTIONS CO-ORDINATES ALL OTHER SUBS
Private Sub My_Logic()

Dim i As Integer
Dim j As Integer
Dim str As String

'UPDATE DISCARD AND PROBABLE TABLES
Call Upgrade_Tables

'TRY THE HARDCODED WORDS
If no_of_guess <= 3 Then
    total_letters = 0
    For i = 1 To fgridVAL2.row - 1
        total_letters = total_letters + Int(fgridVAL2.TextMatrix(i, 2)) _
        + Int(fgridVAL2.TextMatrix(i, 3))
    Next i
    If total_letters < 3 And no_of_guess < 3 Then
        lblguess.Caption = guess_word(no_of_guess + 1, word_num)
    End If
    
End If

'ACTUALLY THINK AND GUESS THE NEXT WORD
If no_of_guess >= 3 Or total_letters >= 3 Then
    If try = 0 Then
        Call filter_database(str)
        try = 1
    Else
        Call filter_array(str)
    End If
    
    lblguess.Caption = str

End If

no_of_guess = no_of_guess + 1

End Sub
'-------------------------------------------------------

Private Sub back_Click()
    frmGAMECHOICE.Visible = True
    Unload Me
End Sub
'-------------------------------------------------------

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 26
    back.ForeColor = &HFFFF&
End Sub
'-------------------------------------------------------

Private Sub cmdBINGO_Click()
MsgBox "I WIN!! Hurrah! Yeah! LOSAH!...Come Again"
frmGAMECHOICE.Visible = True
Unload Me
End Sub
'-------------------------------------------------------

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 22
    back.ForeColor = &HC0C0&
    lblQUIT.FontSize = 22
    lblQUIT.ForeColor = &HFF&
    lblCOMMIT.FontSize = 22
    lblCOMMIT.ForeColor = &HFF00&
End Sub '-------------------------------------------------------

'ENTER THE NUMBER OF COWS AND BULLS
Private Sub lblCOMMIT_Click()

Dim new_row As Byte
Dim str As String
Dim flag As Byte

    If Len(lblguess.Caption) <> 4 Then
        MsgBox "Wait for comp !!"
        Exit Sub
    End If
    
    If cmbCOWZ.ListIndex = -1 Then
        MsgBox "lol .. rofl .. lol.. how many cowz??"
        Exit Sub
    End If
    
    If cmbBULLZ.ListIndex = -1 Then
        MsgBox "lol .. rofl .. lol.. how many bullz??"
        Exit Sub
    End If
    
    If cmbCOWZ.ListIndex + cmbBULLZ.ListIndex > 4 Then
        MsgBox "you need to check your sense of humour!"
        Exit Sub
    End If
    
    If cmbBULLZ.Text = 4 Then
        Call cmdBINGO_Click
        Exit Sub
    End If

    new_row = fgridVAL2.row
    insert new_row, 0, new_row
    insert new_row, 1, lblguess.Caption
    insert new_row, 2, cmbCOWZ.Text
    insert new_row, 3, cmbBULLZ.Text
    fgridVAL2.row = new_row + 1
    
    cnum = cmbCOWZ.Text
    bnum = cmbBULLZ.Text
    
    cmbCOWZ.ListIndex = -1
    cmbBULLZ.ListIndex = -1
    cmbCOWZ.SetFocus
    Call My_Logic
   
End Sub
'-------------------------------------------------------



Private Sub insert_flex(row As Byte, col As Byte, str As String)
    With fgridVAL2
        .row = row
        .col = col
        .Text = str
    End With
End Sub
'-------------------------------------------------------

Private Sub Form_Load()
Dim str As String

guess_word(1, 1) = "sulk"
guess_word(1, 2) = "silk"
guess_word(1, 3) = "musk"

guess_word(2, 1) = "idea"
guess_word(2, 2) = "dean"
guess_word(2, 3) = "iota"

guess_word(3, 1) = "myth"
guess_word(3, 2) = "tyro"
guess_word(3, 3) = "help"

insert_flex 0, 0, "Number"
insert_flex 0, 1, "Words"
insert_flex 0, 2, "Cowz"
insert_flex 0, 3, "Bullz"
fgridVAL2.row = fgridVAL2.row + 1

Call init_all

Set con2 = New ADODB.connection
Set rec2 = New ADODB.recordset

With rec2
       .CursorType = adOpenKeyset
       .CursorLocation = adUseClient
       .LockType = adLockPessimistic
       
End With
    
con2.Open "dsn=adodc", "scott", "tiger"
rec2.Open "select * from tp_word", con2
    
Randomize
Call AIRandom(3)
lblguess.Caption = guess_word(1, word_num)

no_of_guess = 1

End Sub
'-------------------------------------------------------

'SELECTS A RANDOM SET OF WORDS TO START WITH
Private Sub AIRandom(upper_bound As Byte)
    word_num = Int((upper_bound - 1 + 1) * Rnd) + 1
    
End Sub
'-------------------------------------------------------

Private Sub Form_Unload(Cancel As Integer)
If rec2.State = adStateOpen Then rec2.Close
If con2.State = adStateOpen Then con2.Close
End Sub
'-------------------------------------------------------

Private Sub lblCOMMIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCOMMIT.FontSize = 26
    lblCOMMIT.ForeColor = &HFFFF&
End Sub
'-------------------------------------------------------

Private Sub lblQUIT_Click()
    frmINTRO.Visible = True
    Unload Me
End Sub
'-------------------------------------------------------

Private Sub lblQUIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblQUIT.FontSize = 26
    lblQUIT.ForeColor = &HFFFF&
End Sub
'-------------------------------------------------------

'END OF COMPUTER GUESSES MODULE
'======================================================
