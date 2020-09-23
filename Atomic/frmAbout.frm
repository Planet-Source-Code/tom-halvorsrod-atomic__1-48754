VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Atomic v. 1.0"
   ClientHeight    =   3615
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7215
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00000000&
      Caption         =   "OK"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    Dim str As String
    str = "About Atomic - v. 1.0" & vbCrLf
    str = str & "Made by Tom Halvorsroed, Oslo, Norway, 2003" & vbCrLf
    str = str & vbCrLf
    str = str & "Idea and graphics used shamelessly from the KAtomic(Linux) game by Andreas Würst" & vbCrLf
    str = str & "Ini-routines by sitush@aol.com on http://www.Planet-Source-Code.com" & vbCrLf
    str = str & vbCrLf
    str = str & "Realeased at www.planetsourcecode.com as GPL. "
    str = str & "Improvements and stuff can be sent to tomwh@online.no"
    lblMessage.Caption = str
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
