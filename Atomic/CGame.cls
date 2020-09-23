VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'image position var's
Private anBoardAtomPos() As Integer
Private anResAtomPos() As Integer
Private anResTickPos() As Integer
 
'graphics var's
Private dcRes As Long
Private dcMask As Long
Private dcAtom As Long
Private bmpBoard As Long
Private bmpSolution As Long
Private bmpAtom As Long

'game var's
Private astrBoard() As String
Private astrAtoms() As String
Private anAtomIndex() As Integer
Private strMoves() As String

'local variable(s) to hold property value(s)
Private m_bGamePresent As Boolean 'local copy
Private m_strName As String 'local copy
Private m_nNumAtoms As Integer 'local copy
Private m_strSolution As String 'local copy
Private m_strLevelDir As String 'local copy
Private m_nLevel As Integer 'local copy
Private m_nScore As Integer 'local copy
Private m_dcBoard As Long 'local copy
Private m_dcSolution As Long 'local copy

Public Property Get dcSolution() As Long
    dcSolution = m_dcSolution
End Property

Public Property Get dcBoard() As Long
    dcBoard = m_dcBoard
End Property

Public Property Get nScore() As Integer
    nScore = m_nScore
End Property

Public Property Let nLevel(ByVal vData As Integer)
    m_nLevel = vData
End Property

Public Property Get nLevel() As Integer
    nLevel = m_nLevel
End Property

Public Property Let strLevelDir(ByVal vData As String)
    m_strLevelDir = vData
    If Right(m_strLevelDir, 1) <> "\" Then m_strLevelDir = m_strLevelDir & "\"
    If Dir(m_strLevelDir) = "" Then
        MsgBox "Couldn't find entry for levels dir"
        End
    End If
End Property

Public Property Get strLevelDir() As String
    strLevelDir = m_strLevelDir
End Property

Public Property Get strSolution() As String
    strSolution = m_strSolution
End Property

Public Property Get nNumAtoms() As Integer
    nNumAtoms = m_nNumAtoms
End Property

Public Property Get strName() As String
    strName = m_strName
End Property

Public Property Get bGamePresent() As Boolean
    bGamePresent = m_bGamePresent
End Property

Public Function getBoardWithMoves() As String
    Dim strTempBoard As String
    Dim strTemp() As String
    ReDim strMoves(226)
    ReDim strTemp(226)
    Dim str As String
    Dim pos As Integer
    Dim i As Integer
    Dim j As Integer
    
    'make a copy to play with
    strTempBoard = astrBoard(nScore)
    
    'get current info
    pos = getCurrentMarker()
    str = getCurrentBoard()
    
    'calculate first move
    strMoves(SlideAtom(str, pos, -15)) = "U"
    strMoves(SlideAtom(str, pos, 1)) = "R"
    strMoves(SlideAtom(str, pos, 15)) = "D"
    strMoves(SlideAtom(str, pos, -1)) = "L"
    
    'remove the marker from further calculations
    Mid(str, pos, 1) = "."
    
    'calculate up to 225 moves
    For j = 1 To 225
        'check the 225 cells
        For i = 1 To 225
            If Len(strMoves(i)) = j And pos <> i Then
                    'calculate next moves
                    strTemp(SlideAtom(str, i, -15)) = strMoves(i) & "U"
                    strTemp(SlideAtom(str, i, 1)) = strMoves(i) & "R"
                    strTemp(SlideAtom(str, i, 15)) = strMoves(i) & "D"
                    strTemp(SlideAtom(str, i, -1)) = strMoves(i) & "L"
            End If
        Next
    
        'update strmoves(i) if empty
        For i = 1 To 225
            If strMoves(i) = "" Then strMoves(i) = strTemp(i)
        Next
    Next
    
    'remove the current position from the list
    strMoves(pos) = ""
    
    'set a "T" marker if necessary
    For i = 1 To 225
        If strMoves(i) <> "" Then Mid(strTempBoard, i, 1) = "T"
    Next
    
    'return the processed board
    getBoardWithMoves = strTempBoard
End Function

Public Function getBoardWithArrows() As String
    Dim strTempBoard As String
    
    'make a copy to play with
    strTempBoard = astrBoard(nScore)
    
    'set the arrows for the current marker
    'over
    If Mid(strTempBoard, anAtomIndex(nScore) - 15, 1) = "." Then Mid(strTempBoard, anAtomIndex(nScore) - 15, 1) = "U"
    
    'right
    If Mid(strTempBoard, anAtomIndex(nScore) + 1, 1) = "." Then Mid(strTempBoard, anAtomIndex(nScore) + 1, 1) = "R"
    
    'down
    If Mid(strTempBoard, anAtomIndex(nScore) + 15, 1) = "." Then Mid(strTempBoard, anAtomIndex(nScore) + 15, 1) = "D"
    
    'left
    If Mid(strTempBoard, anAtomIndex(nScore) - 1, 1) = "." Then Mid(strTempBoard, anAtomIndex(nScore) - 1, 1) = "L"
        
    getBoardWithArrows = strTempBoard
End Function

Public Function getAtom(ByVal Index As Integer) As String
    getAtom = astrAtoms(Index)
End Function

Public Function getCurrentMarker() As Integer
    getCurrentMarker = anAtomIndex(nScore)
End Function

Public Function getCurrentBoard() As String
    getCurrentBoard = astrBoard(nScore)
End Function

Public Function ParseLevel() As Boolean
    Dim strFilNavn As String
    Dim strLiner As String
    Dim i As Integer
    ReDim astrAtoms(127)
    ReDim astrBoard(1000)
    ReDim anAtomIndex(1000)
    
    ParseLevel = False
    
    'set the correct filename
    strFilNavn = strLevelDir & "level_" & CStr(nLevel)
    
    'extract levelname
    m_strName = mfncGetFromIni("Level", "Name", strFilNavn)
       
    'get the atoms
    m_nNumAtoms = 0
    
    '1-9
    'parse atoms
    For i = 49 To 106
        strLiner = "atom_" & Chr(i)
        strLiner = mfncGetFromIni("Level", strLiner, strFilNavn)
        If strLiner <> "" Then
            astrAtoms(i) = strLiner
            m_nNumAtoms = m_nNumAtoms + 1
        End If
        If i = 57 Then i = 96
    Next i
    
    'get the board
    astrBoard(0) = ""
    For i = 0 To 14
        strLiner = "feld_"
        If i < 10 Then
            strLiner = strLiner & "0"
        End If
        strLiner = strLiner & CStr(i)
        strLiner = mfncGetFromIni("Level", strLiner, strFilNavn)
        astrBoard(0) = astrBoard(0) & strLiner
    Next i
    
    'get solution
    m_strSolution = ""
    For i = 0 To 9
        strLiner = "mole_" & i
        strLiner = mfncGetFromIni("Level", strLiner, strFilNavn)
        If strLiner <> "" Then
            strLiner = strLiner & String(15, ".")
            strLiner = Mid(strLiner, 1, 15)
            m_strSolution = strSolution & strLiner
        End If
    Next i
    
    'set score
    m_nScore = 0
    
    'find first marker
    anAtomIndex(0) = 1
    If StepAtom(1) Then ParseLevel = True
    
    m_bGamePresent = ParseLevel
End Function

'check for a valid solution
Public Function CheckSolution() As Boolean
    Dim strBoardTemp As String
    
    CheckSolution = False
    
    'make the board compatible with the solution by removing the walls
    strBoardTemp = Replace(astrBoard(nScore), "#", ".")
    
    'if board_temp contain the string solution then we've found it
    If InStr(strBoardTemp, strSolution) > 0 Then CheckSolution = True
End Function

Public Function IncreaseLevel() As Boolean
    IncreaseLevel = False
    If nLevel < 67 Then
        IncreaseLevel = True
        nLevel = nLevel + 1
    End If
End Function

Public Function DecreaseLevel() As Boolean
    DecreaseLevel = False
    If nLevel > 1 Then
        DecreaseLevel = True
        nLevel = nLevel - 1
    End If
End Function

Public Function MoveAtom(ByVal dt As Integer) As Boolean
    Dim nSanity As Integer
    MoveAtom = False
    
    nSanity = 0
    
    'can we move
    If Mid(astrBoard(nScore), anAtomIndex(nScore) + dt, 1) <> "." Then Exit Function
    
    'store current situation
    MoveAtom = True
    m_nScore = m_nScore + 1
    astrBoard(nScore) = astrBoard(m_nScore - 1)
    anAtomIndex(nScore) = anAtomIndex(m_nScore - 1)
    
    'move it
    While (nSanity < 100)
       If Mid(astrBoard(m_nScore), anAtomIndex(m_nScore) + dt, 1) = "." Then
            Mid(astrBoard(m_nScore), anAtomIndex(m_nScore) + dt, 1) = Mid(astrBoard(m_nScore), anAtomIndex(m_nScore), 1)
            Mid(astrBoard(m_nScore), anAtomIndex(m_nScore), 1) = "."
            anAtomIndex(m_nScore) = anAtomIndex(m_nScore) + dt
        Else
            Exit Function
        End If
        'just as precaution; to avoid inf. loop
        nSanity = nSanity + 1
    Wend
    
End Function

'helper function for StepAtom
Private Function FindAtom(ByVal nStart As Integer, ByVal nEnd As Integer, ByVal nStep As Integer) As Boolean
    Dim i As Integer
    Dim strTemp As String
    
    FindAtom = False
    For i = nStart To nEnd Step nStep
        If Mid(astrBoard(nScore), i, 1) <> "#" And Mid(astrBoard(nScore), i, 1) <> "." Then
            anAtomIndex(nScore) = i
            strTemp = getBoardWithArrows
            If (InStr(1, strTemp, "D") + InStr(1, strTemp, "L") + InStr(1, strTemp, "R") + InStr(1, strTemp, "U")) > 0 Then
                FindAtom = True
                Exit Function
            End If
        End If
    Next i
End Function

'find next/prev atom
'iDir=1  => next atom
'iDir=-1 => prev atom
Public Function StepAtom(ByVal nDir As Integer) As Boolean
    StepAtom = True
        
    'find the next atom
    If nDir = 1 Then
        If FindAtom(anAtomIndex(nScore) + 1, 225, nDir) Then Exit Function
        If FindAtom(1, anAtomIndex(nScore) - 1, nDir) Then Exit Function
    'find the prev atom
    Else
        If FindAtom(anAtomIndex(nScore) - 1, 1, nDir) Then Exit Function
        If FindAtom(225, anAtomIndex(nScore) + 1, nDir) Then Exit Function
    End If
    
    StepAtom = False
End Function

Public Sub ResetBoard()
    m_nScore = 0
End Sub

Public Function UndoMove() As Boolean
    UndoMove = False
    If nScore > 0 Then
        m_nScore = m_nScore - 1
        UndoMove = True
    End If
End Function

Private Function SlideAtom(ByVal str As String, ByVal pos As Integer, ByVal direct As Integer) As Integer
    SlideAtom = 0
    If Mid(str, pos + direct, 1) = "." Then
        While Mid(str, pos + direct, 1) = "."
            pos = pos + direct
        Wend
        SlideAtom = pos
    End If
End Function

Public Sub CreateBitmaps()
    Dim i As Integer
    Dim nLen As Integer
    
    'create the atombitmap
    For i = 40 To 126
        CreateAtom i
    Next i
    
    'clear the solution dc
    BitBlt dcSolution, 0, 0, 500, 500, 0, 0, 0, vbWhiteness
    
    nLen = Len(strSolution)
    
    For i = 1 To nLen
        PaintAtom dcSolution, Asc(Mid(strSolution, i, 1)), (i Mod 15) * 30, (i \ 15) * 30, 0
    Next i
End Sub

Public Function LoadResources() As Boolean
    ReDim anBoardAtomPos(127, 2)
    ReDim anResAtomPos(127, 2)
    ReDim anResTickPos(127, 2)
    
    'left and top positions in the bitmaps
    anResAtomPos(50, 0) = 31
    anResAtomPos(51, 0) = 62
    anResAtomPos(52, 0) = 93
    anResAtomPos(53, 0) = 124
    anResAtomPos(54, 0) = 155
    anResAtomPos(55, 0) = 186
    anResAtomPos(56, 0) = 217
    anResAtomPos(57, 0) = 248
    anResAtomPos(65, 0) = 62
    anResAtomPos(66, 0) = 93
    anResAtomPos(67, 0) = 124
    anResAtomPos(68, 0) = 155
    anResAtomPos(97, 0) = 279
    anResAtomPos(111, 0) = 31
        
    anResAtomPos(50, 1) = 0
    anResAtomPos(51, 1) = 0
    anResAtomPos(52, 1) = 0
    anResAtomPos(53, 1) = 0
    anResAtomPos(54, 1) = 0
    anResAtomPos(55, 1) = 0
    anResAtomPos(56, 1) = 0
    anResAtomPos(57, 1) = 0
    anResAtomPos(65, 1) = 93
    anResAtomPos(66, 1) = 93
    anResAtomPos(67, 1) = 93
    anResAtomPos(68, 1) = 93
    anResAtomPos(97, 1) = 0
    anResAtomPos(111, 1) = 93

    anResTickPos(65, 0) = 0
    anResTickPos(66, 0) = 31
    anResTickPos(67, 0) = 62
    anResTickPos(68, 0) = 93
    anResTickPos(69, 0) = 124
    anResTickPos(70, 0) = 155
    anResTickPos(71, 0) = 186
    anResTickPos(72, 0) = 217
    anResTickPos(97, 0) = 0
    anResTickPos(98, 0) = 31
    anResTickPos(99, 0) = 62
    anResTickPos(100, 0) = 93
    anResTickPos(101, 0) = 124
    anResTickPos(102, 0) = 155
    anResTickPos(103, 0) = 186
    anResTickPos(104, 0) = 217

    anResTickPos(65, 1) = 62
    anResTickPos(66, 1) = 62
    anResTickPos(67, 1) = 62
    anResTickPos(68, 1) = 62
    anResTickPos(69, 1) = 62
    anResTickPos(70, 1) = 62
    anResTickPos(71, 1) = 62
    anResTickPos(72, 1) = 62
    anResTickPos(97, 1) = 31
    anResTickPos(98, 1) = 31
    anResTickPos(99, 1) = 31
    anResTickPos(100, 1) = 31
    anResTickPos(101, 1) = 31
    anResTickPos(102, 1) = 31
    anResTickPos(103, 1) = 31
    anResTickPos(104, 1) = 31

    anBoardAtomPos(35, 0) = 248
    anBoardAtomPos(46, 0) = 186
    anBoardAtomPos(68, 0) = 248
    anBoardAtomPos(76, 0) = 217
    anBoardAtomPos(82, 0) = 279
    anBoardAtomPos(83, 0) = 279
    anBoardAtomPos(84, 0) = 279
    anBoardAtomPos(85, 0) = 248

    anBoardAtomPos(35, 1) = 31
    anBoardAtomPos(46, 1) = 93
    anBoardAtomPos(68, 1) = 93
    anBoardAtomPos(76, 1) = 93
    anBoardAtomPos(82, 1) = 93
    anBoardAtomPos(83, 1) = 62
    anBoardAtomPos(84, 1) = 31
    anBoardAtomPos(85, 1) = 62
                
    'left image positions for painting
    anBoardAtomPos(49, 0) = 1470 '1
    anBoardAtomPos(50, 0) = 1500 '2
    anBoardAtomPos(51, 0) = 1530 '3
    anBoardAtomPos(52, 0) = 1560 '4
    anBoardAtomPos(53, 0) = 1590 '5
    anBoardAtomPos(54, 0) = 1620 '6
    anBoardAtomPos(55, 0) = 1650 '7
    anBoardAtomPos(56, 0) = 1680 '8
    anBoardAtomPos(57, 0) = 1710 '9
    anBoardAtomPos(97, 0) = 2910 'a
    anBoardAtomPos(98, 0) = 2940 'b
    anBoardAtomPos(99, 0) = 2970 'c
    anBoardAtomPos(100, 0) = 3000 'd
    anBoardAtomPos(101, 0) = 3030 'e
    anBoardAtomPos(102, 0) = 3060 'f
    anBoardAtomPos(103, 0) = 3090 'g
    anBoardAtomPos(104, 0) = 3120 'h
    anBoardAtomPos(105, 0) = 3150 'i
    anBoardAtomPos(106, 0) = 3180 'j
        
    'load graphics
    dcRes = CreateCompatibleDC(GetDC(0))
    SelectObject dcRes, LoadResPicture("ID_BITMAP_RES", 0)
    
    dcMask = CreateCompatibleDC(GetDC(0))
    SelectObject dcMask, LoadResPicture("ID_BITMAP_MASK", 0)
    
    'create graphics handles to work with
    m_dcBoard = CreateCompatibleDC(GetDC(0))
    bmpBoard = CreateCompatibleBitmap(GetDC(0), 700, 700)
    SelectObject m_dcBoard, bmpBoard

    m_dcSolution = CreateCompatibleDC(GetDC(0))
    bmpSolution = CreateCompatibleBitmap(GetDC(0), 700, 700)
    SelectObject m_dcSolution, bmpSolution
    
    dcAtom = CreateCompatibleDC(GetDC(0))
    bmpAtom = CreateCompatibleBitmap(GetDC(0), 3810, 30)
    SelectObject dcAtom, bmpAtom
End Function

Public Sub PaintBoard(ByVal nType As Integer)
    Dim i As Integer, j As Integer
    Dim x As Long
    Dim y As Long
    Dim strTempBoard As String
    
    'make a copy to play with
    If nType = 1 Then
        strTempBoard = getBoardWithArrows()
    Else
        strTempBoard = getBoardWithMoves()
    End If
    
    'display the screen
    For i = 1 To 225
        x = ((i - 1) Mod 15) * 30
        y = ((i - 1) \ 15) * 30
        
        PaintAtom dcBoard, Asc(Mid(strTempBoard, i, 1)), x, y, i
        
        'paint the outline of the marker
        If i = getCurrentMarker() And nType = 2 Then
            BitBlt dcBoard, x, y, 30, 30, dcMask, anBoardAtomPos(83, 0), anBoardAtomPos(83, 1), vbSrcAnd
            BitBlt dcBoard, x, y, 30, 30, dcRes, anBoardAtomPos(83, 0), anBoardAtomPos(83, 1), vbSrcPaint
        End If
    Next i
End Sub

Private Sub CreateAtom(ByVal nIndex As Integer)
    Dim nAtomLeft As Integer
    Dim nAtomTop As Integer
        
    Dim lRes As Long
    Dim nLength As Integer
    
    Dim i As Integer
    Dim nLeftPos As Integer
    Dim nTopPos As Integer
    Dim strTick As String
    Dim strAtom As String
    
    strAtom = getAtom(nIndex)
    If strAtom = "" Then Exit Sub
    nLength = Len(strAtom)
    
    nAtomLeft = anResAtomPos(Asc(Mid(strAtom, 1, 1)), 0)
    nAtomTop = anResAtomPos(Asc(Mid(strAtom, 1, 1)), 1)
   
    'get atom
    BitBlt dcAtom, nIndex * 30, 0, 30, 30, dcRes, nAtomLeft, nAtomTop, vbSrcCopy
    
    'paint ticks
    For i = 3 To nLength
        nLeftPos = anResTickPos(Asc(Mid(strAtom, i, 1)), 0)
        nTopPos = anResTickPos(Asc(Mid(strAtom, i, 1)), 1)
        BitBlt dcAtom, nIndex * 30, 0, 30, 30, dcMask, nLeftPos, nTopPos, vbSrcAnd
        BitBlt dcAtom, nIndex * 30, 0, 30, 30, dcRes, nLeftPos, nTopPos, vbSrcPaint
    Next i
End Sub

Private Sub PaintAtom(ByVal destDC As Long, ByVal nAtomIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nIndex As Integer)
    Dim dc As Long
    Dim i As Integer
    
    'assume dcAtom as current dc
    dc = dcAtom
    
    'if it's an new position, arrow, wall or floor then use the original bitmap
    Select Case nAtomIndex
        Case 35, 46, 68, 76, 82, 84, 85
            dc = dcRes
    End Select
    
    'if painting possible moves
    If nAtomIndex = 84 Then
        i = (Len(strMoves(nIndex)) - 1) * 31
        If i > 248 Then i = 248
        BitBlt destDC, x, y, 30, 30, dc, i, 124, vbSrcCopy
    'painting everything else
    Else
        BitBlt destDC, x, y, 30, 30, dc, anBoardAtomPos(nAtomIndex, 0), anBoardAtomPos(nAtomIndex, 1), vbSrcCopy
    End If
End Sub

Public Sub DestroyHandles()
    'destroy graphics handles
    DeleteDC dcRes
    DeleteDC dcMask
    DeleteDC dcBoard
    DeleteDC dcSolution
    DeleteDC dcAtom
    DeleteObject bmpSolution
    DeleteObject bmpBoard
    DeleteObject bmpAtom
End Sub

Private Sub Class_Initialize()
    m_bGamePresent = False
End Sub

'return the moves list for nPos
Public Function getMoves(ByVal nPos As Integer) As String
    getMoves = strMoves(nPos)
End Function

Public Function GotoAtom(ByVal nPos As Integer) As Boolean
    Dim strTemp As String
    strTemp = Mid(astrBoard(nScore), nPos, 1)
    
    GotoAtom = False
    
    'just as a precaution
    If strTemp = "#" Or strTemp = "." Then Exit Function
    
    'if the atom has a floor next to it then it can be selected for movement
    If Mid(astrBoard(nScore), nPos - 15, 1) = "." Then GotoAtom = True
    If Mid(astrBoard(nScore), nPos - 1, 1) = "." Then GotoAtom = True
    If Mid(astrBoard(nScore), nPos + 1, 1) = "." Then GotoAtom = True
    If Mid(astrBoard(nScore), nPos + 15, 1) = "." Then GotoAtom = True
    
    'if we can select the atom then change the position of the currently selected atom
    If GotoAtom = True Then anAtomIndex(nScore) = nPos
End Function
