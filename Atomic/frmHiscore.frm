VERSION 5.00
Begin VB.Form frmHiscore 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hiscore"
   ClientHeight    =   3120
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4635
   Icon            =   "frmHiscore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPlayer 
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPlayer 
      Height          =   375
      Index           =   4
      Left            =   840
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPlayer 
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPlayer 
      Height          =   405
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPlayer 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   315
      ItemData        =   "frmHiscore.frx":0442
      Left            =   120
      List            =   "frmHiscore.frx":0444
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblPlayer 
      BackColor       =   &H00000000&
      Caption         =   "Label15"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   21
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "Label14"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   20
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "Label13"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   19
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "Label12"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   18
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "Label11"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   17
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblPlayer 
      BackColor       =   &H00000000&
      Caption         =   "Label10"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblPlayer 
      BackColor       =   &H00000000&
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblPlayer 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   14
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblPlayer 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   13
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "#5"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "#3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "#4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "#2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "#1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmHiscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MoveScore(ByVal iStart As Integer, ByVal iEnd As Integer)
    Dim j As Integer
    For j = iStart To iEnd Step -1
        txtPlayer(j).Text = txtPlayer(j - 1).Text
        lblPlayer(j).Caption = lblPlayer(j - 1).Caption
        lblScore(j).Caption = lblScore(j - 1).Caption
    Next j
    
    txtPlayer(iEnd - 1).Visible = True
    lblScore(iEnd - 1).Caption = CStr(nCurrentScore)
End Sub

Private Sub Form_Load()
    Dim i As Integer
        
    For i = 1 To 67
        cmbLevel.AddItem "Level " & CStr(i)
    Next i
    
    cmbLevel.ListIndex = theGame.nLevel - 1
    
    If nCurrentScore > 0 Then
        If nCurrentScore < CInt(lblScore(1).Caption) Then
            MoveScore 5, 2
        ElseIf nCurrentScore < CInt(lblScore(2).Caption) Then
            MoveScore 5, 3
        ElseIf nCurrentScore < CInt(lblScore(3).Caption) Then
            MoveScore 5, 4
        ElseIf nCurrentScore < CInt(lblScore(4).Caption) Then
            MoveScore 5, 5
        ElseIf nCurrentScore < CInt(lblScore(5).Caption) Then
            MoveScore 5, 6
        End If
    End If
End Sub

Private Sub cmbLevel_Click()
    Dim nLevel As String
    Dim str As String
    Dim i As Integer

    nLevel = "Level" & CStr(cmbLevel.ListIndex + 1)
    For i = 1 To 5
        str = "Name" & CStr(i)
        lblPlayer(i).Caption = mfncGetFromIni(nLevel, str, strScoreFile)
        txtPlayer(i).Text = mfncGetFromIni(nLevel, str, strScoreFile)
        str = "Score" & CStr(i)
        lblScore(i).Caption = mfncGetFromIni(nLevel, str, strScoreFile)
    Next i
End Sub

Private Sub OKButton_Click()
    Dim nLevel As String
    Dim i As Integer
    Dim nRet As Integer
    
    If nCurrentScore > 0 Then
        nLevel = "Level" & CStr(cmbLevel.ListIndex + 1)
        For i = 1 To 5
            nRet = mfncWriteIni(nLevel, "Name" & CStr(i), txtPlayer(i).Text, strScoreFile)
            If nRet = 0 Then GoTo errhandler
            
            nRet = mfncWriteIni(nLevel, "Score" & CStr(i), lblScore(i).Caption, strScoreFile)
            If nRet = 0 Then GoTo errhandler
        Next i
    End If
    
    Unload Me
    Exit Sub
errhandler:
    MsgBox "huh"
End Sub

