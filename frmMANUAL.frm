VERSION 5.00
Begin VB.Form frmMANUAL 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "The Manual"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMANUAL.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6975
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRULES 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5055
      Left            =   960
      Locked          =   -1  'True
      MouseIcon       =   "frmMANUAL.frx":030A
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMANUAL.frx":0614
      Top             =   2520
      Width           =   9615
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
      Left            =   8160
      MouseIcon       =   "frmMANUAL.frx":093C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   7800
      Width           =   2565
   End
   Begin VB.Label lblHEADER 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "THE MANUAL OF GAME RULES"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   787
      TabIndex        =   1
      Top             =   360
      Width           =   5700
   End
   Begin VB.Label lblRULES 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Cowz And Bullz!!! "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   3480
   End
End
Attribute VB_Name = "frmMANUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***** FORM : MANUAL ****

Private Sub back_Click()
    frmINTRO.Visible = True
    Unload Me
End Sub
'-----------------------------------------------------

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 26
    back.ForeColor = &HFFFF&
End Sub
'-----------------------------------------------------

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    back.FontSize = 22
    back.ForeColor = &HC0C0&
End Sub
'-----------------------------------------------------
