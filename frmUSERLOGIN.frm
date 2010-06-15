VERSION 5.00
Begin VB.Form frmUSERLOGIN 
   Caption         =   "Login"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQUIT 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtNAME 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Enter Your Name Here"
      Top             =   600
      Width           =   2415
   End
   Begin VB.ComboBox cmbLOGIN 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdLOGIN 
      Caption         =   "&LOGIN"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblNEW 
      AutoSize        =   -1  'True
      Caption         =   "New User:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   750
   End
   Begin VB.Label lblUSER 
      AutoSize        =   -1  'True
      Caption         =   "Existing User:"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmUSERLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** FORM : USER LOGIN FOR THE GAME ****

Dim connection As New ADODB.connection
Dim recordset As New ADODB.recordset
Public struser As String
Public user_time As Date
Public int_rating As Integer
Public play_count As Integer
Public win_perc As Single
Public best_effort As Byte
'-----------------------------------------------------


'THIS SUB CREATES A NEW LOGIN IF THE USER IS NOT IN THE DBASE
'AND LOGS-IN THE EXISTING USER
Private Sub cmdLOGIN_Click()
    If cmbLOGIN.ListIndex = -1 Then
        If Len(txtNAME.Text) = 0 Then
           Exit Sub
        End If
        
        '2 USERS WITH SAME NAME ARE NOT ALLOWED
        With recordset
            .MoveFirst
            Do While Not .EOF
                If !Name = txtNAME.Text Then
                    MsgBox "User exists !!"
                    txtNAME.Text = ""
                    txtNAME.SetFocus
                    Exit Sub
                End If
            .MoveNext
            Loop
        End With
        'CREATE THE LOGIN !
        recordset.AddNew
        recordset!Name = txtNAME.Text
        recordset!int_rating = 0
        recordset!play_count = 0
        recordset!win_perc = 0
        recordset!best_effort = "NA"
        recordset.Update
        struser = txtNAME.Text
    Else
        struser = cmbLOGIN.Text
    End If
    recordset.Close
    connection.Close
    
    'SAVE THE TIME OF LOGIN
    user_time = Date
            
    frmUSERLOGIN.Visible = False
    frmINTRO.Visible = True
    
End Sub
'-----------------------------------------------------

Private Sub cmdQUIT_Click()
    End
End Sub
'-----------------------------------------------------

Private Sub Form_Load()
    connection.Open "dsn=adodc", "scott", "tiger"
    recordset.CursorLocation = adUseClient
    recordset.CursorType = adOpenKeyset
    recordset.LockType = adLockPessimistic
    
    recordset.Open "select * from userinfo", connection
    Do While Not recordset.EOF
        cmbLOGIN.AddItem (recordset!Name)
        recordset.MoveNext
    Loop
    
    
End Sub
'-----------------------------------------------------

