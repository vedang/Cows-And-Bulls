VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1380
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Frequency 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblCAPTION 
      AutoSize        =   -1  'True
      Caption         =   "Enter Frequency of the word:"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
