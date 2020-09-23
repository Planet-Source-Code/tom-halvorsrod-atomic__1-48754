VERSION 5.00
Begin VB.Form frmGraphics 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atomic v. 1.0"
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGraphics.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   6720
      Max             =   67
      Min             =   1
      TabIndex        =   3
      Top             =   2400
      Value           =   1
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9480
      Top             =   5400
   End
   Begin VB.Label lblControls 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   6765
      TabIndex        =   2
      Top             =   2640
      Width           =   2670
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Solution"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6770
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Width           =   9495
   End
   Begin VB.Menu mnuGam 
      Caption         =   "&Game"
      Begin VB.Menu mnuGame 
         Caption         =   "&Hiscore"
         Index           =   0
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Previous level"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "N&ext level"
         Index           =   2
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Reset level"
         Index           =   3
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuGame 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPref 
      Caption         =   "&Prefs"
      Begin VB.Menu mnuPrefs 
         Caption         =   "&Arrows"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPrefs 
         Caption         =   "&Moves"
         Index           =   1
      End
   End
   Begin VB.Menu mnuHel 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuHelp 
         Caption         =   "&About..."
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bMarkerOnOff As Boolean

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If theGame.bGamePresent Then
        'process the key
        Select Case KeyCode
            Case 8 'backspace
                theGame.UndoMove
            Case 37 'left
                theGame.MoveAtom DIR_LEFT
            Case 38 'up
                theGame.MoveAtom DIR_UP
            Case 39 'right
                theGame.MoveAtom DIR_RIGHT
            Case 40 'down
                theGame.MoveAtom DIR_DOWN
            Case 67 'C prev atom
                theGame.StepAtom -1
            Case 86 'V next atom
                theGame.StepAtom 1
        End Select
    
        Form_Paint
    End If
End Sub

Private Sub Form_Load()
    Dim str As String
    
    'create new game
    Set theGame = New CGame
    
    'load level dir from atomic.ini
    theGame.strLevelDir = mfncGetFromIni("FileAndDir", "levels", ".\atomic.ini")
    
    'set the filename and path to the current score file
    strScoreFile = mfncGetFromIni("FileAndDir", "scoredat", ".\atomic.ini")
    If Dir(strScoreFile) = "" Then
        MsgBox "Couldn't find entry for score file"
        End
    End If
    
    'load and create bitmaps from AtomicGraphics.res
    theGame.LoadResources
    
    'display instructions
    str = "Controls:" & vbCrLf
    str = str & "C - prev atom" & vbCrLf
    str = str & "V - next atom" & vbCrLf
    str = str & "Else use the arrows to move the current atom to make it look like the solution" & vbCrLf & vbCrLf
    str = str & "Backspace for undo"
    str = str & vbCrLf & vbCrLf
    str = str & "You can also use mouse-clicks on atoms, arrows and numbers"
    lblControls.Caption = str
    
    'load level 1
    theGame.nLevel = 1
    NewLevel
    
    'set blinking marker for current atom to on
    bMarkerOnOff = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not theGame.bGamePresent Then Exit Sub
    Dim str As String
    Dim i As Integer, j As Integer
    Dim pos As Integer
    Dim sTempBoard As String
    Dim cell As String
    Dim iSanity As Integer
    
    'make a copy to play with
    If mnuPrefs(0).Checked Then
        sTempBoard = theGame.getBoardWithArrows()
    Else
        sTempBoard = theGame.getBoardWithMoves()
    End If
    
    'get pos
    i = ((x - 225) / 450)
    j = ((y - 225) / 450)
    
    If i < 0 Or i > 14 Then Exit Sub
    If j < 0 Or j > 14 Then Exit Sub
    
    pos = j * 15 + i + 1
    
    'get cell
    cell = Mid(sTempBoard, pos, 1)
        
    'if it's a wall or a floor; no need do anything
    If cell = "." Then
        Exit Sub
    ElseIf cell = "#" Then
        Exit Sub
        
    'several atoms to be moved, graphically displayed by numbers
    ElseIf cell = "T" Then
        MoveSeveralAtoms pos
        
    'process directions
    ElseIf cell = "U" Then
        theGame.MoveAtom DIR_UP
    ElseIf cell = "R" Then
        theGame.MoveAtom DIR_RIGHT
    ElseIf cell = "D" Then
        theGame.MoveAtom DIR_DOWN
    ElseIf cell = "L" Then
        theGame.MoveAtom DIR_LEFT
        
    'it has to be an atom, so select that one
    Else
        theGame.GotoAtom (pos)
    End If
    
    'paint the board
    Form_Paint
End Sub

Private Sub Form_Paint()
    If theGame.bGamePresent Then
        If mnuPrefs(0).Checked Then
            theGame.PaintBoard 1
        Else
            If bMarkerOnOff Then
                theGame.PaintBoard 2
            Else
                theGame.PaintBoard 3
            End If
        End If
        
        'paint the two pictures; board and solution
        BitBlt Me.hdc, 0, 0, 450, 450, theGame.dcBoard, 0, 0, vbSrcCopy
        StretchBlt Me.hdc, 451, 20, 175, 150, theGame.dcSolution, 0, 0, 350, 300, vbSrcCopy
    
        lblScore.Caption = "Hiscore: " & getHighestScore(theGame.nLevel) & " Score: " & CStr(theGame.nScore)
        
        If theGame.CheckSolution() = True Then
            Timer1.Enabled = False
            MsgBox "Solution found"
            nCurrentScore = theGame.nScore
            'update the score
            frmHiscore.Show vbModal, Me
            'goto next board
            If theGame.IncreaseLevel() = True Then NewLevel
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    theGame.DestroyHandles
End Sub

Private Sub HScroll1_Change()
    theGame.nLevel = HScroll1.Value
    NewLevel
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0
            frmAbout.Show vbModal, Me
    End Select
End Sub

Private Sub NewLevel()
    Dim i As Integer
    Dim sStr As String
    'parse the level, create graphics, display level name and setup timer and stuff
    theGame.ParseLevel
    If Not theGame.bGamePresent Then Exit Sub
    theGame.CreateBitmaps
    Caption = "Atomic v. 1.0 - Level " & CStr(theGame.nLevel) & " " & theGame.strName
    If mnuPrefs(1).Checked Then
        bMarkerOnOff = True
        Timer1.Enabled = True
    End If
    HScroll1.Value = theGame.nLevel
    Form_Paint
End Sub

Private Sub mnuGame_Click(iIndex As Integer)
    Dim sStr As String
    Select Case iIndex
        Case 0 'hiscore
            If theGame.bGamePresent Then
                nCurrentScore = -1
                frmHiscore.Show vbModal, Me
            End If
        Case 1  'prev level
            If theGame.DecreaseLevel() Then NewLevel
        Case 2 'next level
            If theGame.IncreaseLevel() Then NewLevel
        Case 3 'reset level
            If theGame.bGamePresent Then
                theGame.ResetBoard
                Form_Paint
            End If
        Case 5
            End
    End Select
End Sub

Private Sub mnuPrefs_Click(Index As Integer)
    mnuPrefs(0).Checked = Not mnuPrefs(0).Checked
    mnuPrefs(1).Checked = Not mnuPrefs(1).Checked
    
    If mnuPrefs(1).Checked Then
        bMarkerOnOff = True
        Timer1.Enabled = True
    End If
    Form_Paint
End Sub

'when you set Prefs->Moves on the menu you can click on the numbers to move the current atom that many steps
'the function get the move sequence and move the atom accordingly
'update the screen after each move
Public Sub MoveSeveralAtoms(ByVal pos As Integer)
    Dim strSequence As String
    Dim i As Integer
    Dim nLength As Integer
    Dim marker As String
    
    strSequence = theGame.getMoves(pos)
    nLength = Len(strSequence)
    
    For i = 1 To nLength
        marker = Mid(strSequence, i, 1)
        If i > 1 Then SleepEx 400, 0
        If marker = "L" Then theGame.MoveAtom DIR_LEFT
        If marker = "U" Then theGame.MoveAtom DIR_UP
        If marker = "R" Then theGame.MoveAtom DIR_RIGHT
        If marker = "D" Then theGame.MoveAtom DIR_DOWN
        Form_Paint
        DoEvents
    Next
End Sub

'make the current atom "blink"
Private Sub Timer1_Timer()
    Form_Paint
    bMarkerOnOff = Not bMarkerOnOff
End Sub
