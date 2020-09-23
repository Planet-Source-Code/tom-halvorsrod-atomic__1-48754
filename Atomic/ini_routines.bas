Attribute VB_Name = "ini_routines"
Option Explicit
'ini functions
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'**************************************
' Name: .INI read/write routines
' Description:.INI read/write routines
' mfncGetFromIni-- Reads from an *.INI file strFileName (full path & file name)
' mfncWriteIni--Writes To an *.INI file called strFileName (full path & file name)
' sitush@ aol.com
' By: Newsgroup Posting
'
' Returns:mfncGetFromIni--The string sto
'     red in [strSectionHeader], line beginnin
'     g strVariableName
' mfncWriteIni--Integer indicating failure (0) or success (other) To write
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.584/lngWId.1/qx/vb/scripts/ShowCode.h
'     tm'for details.'**************************************

Public Function mfncGetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    '*** DESCRIPTION:Reads from an *.INI fil
    '     e strFileName (full path & file name)
    '*** RETURNS:The string stored in [strSectionHeader], line  beginning strVariableName=
    '*** NOTE: Requires declaration of API call GetPrivateProfileString
    'Initialise variable
    Dim strReturn As String
    'Blank the return string
    strReturn = String(255, Chr(0))
    'Get requested information, trimming the returned string
    mfncGetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function

Public Function mfncParseString(strIn As String, intOffset As Integer, strDelimiter As String) As String
    '*** DESCRIPTION:Parses the passed string, returning the value indicated
    '***by the offset specified, eg: the string "Hello, World ","
    '***offset 2 = "World".
    '*** RETURNS:See description.
    '*** NOTE: The offset starts at 1 and the delimiter is the Character
    '***which separates the elements of the string.
    'Trap any bad calls

    If Len(strIn) = 0 Or intOffset = 0 Then
        mfncParseString = ""
        Exit Function
    End If
    'Declare local variables
    Dim intStartPos As Integer
    ReDim intDelimPos(10) As Integer
    Dim intStrLen As Integer
    Dim intNoOfDelims As Integer
    Dim intCount As Integer
    Dim strQuotationMarks As String
    Dim intInsideQuotationMarks As Integer
    strQuotationMarks = Chr(34) & Chr(147) & Chr(148)
    intInsideQuotationMarks = False


    For intCount = 1 To Len(strIn)
        'If character is a double-quote then tog
        '     gle the In Quotation flag


        If InStr(strQuotationMarks, Mid$(strIn, intCount, 1)) <> 0 Then
            intInsideQuotationMarks = (Not intInsideQuotationMarks)
        End If
        If (Not intInsideQuotationMarks) And (Mid$(strIn, intCount, 1) = strDelimiter) Then
        intNoOfDelims = intNoOfDelims + 1
        'If array filled then enlarge it, keepin
        '     g existing contents


        If (intNoOfDelims Mod 10) = 0 Then
            ReDim Preserve intDelimPos(intNoOfDelims + 10)
        End If
        intDelimPos(intNoOfDelims) = intCount
    End If
Next intCount
'Handle request for value not present (over-run)

If intOffset > (intNoOfDelims + 1) Then
    mfncParseString = ""
    Exit Function
End If
'Handle boundaries of string


If intOffset = 1 Then
    intStartPos = 1
End If
'Requesting last value - handle null


If intOffset = (intNoOfDelims + 1) Then
    If Right$(strIn, 1) = strDelimiter Then
        intStartPos = -1
        intStrLen = -1
        mfncParseString = ""
        Exit Function
    Else
        intStrLen = Len(strIn) - intDelimPos(intOffset - 1)
    End If
End If
'Set start and length variables if not handled by boundary check above


If intStartPos = 0 Then
    intStartPos = intDelimPos(intOffset - 1) + 1
End If


If intStrLen = 0 Then
    intStrLen = intDelimPos(intOffset) - intStartPos
End If
'Set the return string
mfncParseString = Mid$(strIn, intStartPos, intStrLen)
End Function


Public Function mfncWriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    '*** DESCRIPTION:Writes to an *.INI file
    '     called strFileName (full path & file name)
    '*** RETURNS:Integer indicating failure
    '     (0) or success (other) To write
    '*** NOTE: Requires declaration of API call WritePrivateProfileString
    'Call the API
    mfncWriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

