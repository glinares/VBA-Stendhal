'VBA-Stendhal.BAS - Office EMF Steganography Commands by @Laughing_Mantis"
'Beta Version - Non-weaponized
'July 2020

'This is an ongoing project and can change rapidly and without notice
'If you use this in research or talks pls just give me a shoutout
'Feedback and Comments always welcome DM me on Twitter @Laughing_Mantis


Const Version = "Version Beta 1.0.0 - July 2020"

Const debugmode = True
Const TargetFile = "C:\tools\shape-CMD-Test1.emf" ' CHANGE THIS TO YOUR FILE I will include this for testing
Const C2URL = "http://pbs.twimg.com/profile_images/631923531719151616/IpTXHz_t_400x400.jpg"
Const NukeShapes = True ' Deletes all shapes in the document prior to tests - super helpful
Const DataFlagTest = "LAUGHINGMANTISWASHERE"

Public colResults As Collection ' This will store the results of what Stendhal finds


Public Sub StendhalDemo()
Dim cmdCol As Collection
Set cmdCol = New Collection
Dim Password As String
Dim CmdShape As InlineShape
Dim BaseDataArray() As Byte
Dim CmdByteArray() As Byte
Dim CmdArray() As Byte
Dim CmdFile As String
Dim BaseFile As String
Dim Offset As Long
Dim CmdString As String
Dim FF As Long

Debug.Print "VBA-Stendhal.BAS - Office EMF Steganography Commands by @Laughing_Mantis"
Debug.Print Version
Debug.Print "==============================================================================="
Debug.Print ""
Password = CleanEncryptSTR("LaughingMantis", "LaughingMantis", True)
cmdCol.Add Password
cmdCol.Add CleanEncryptSTR("CMD", Password, True)
cmdCol.Add CleanEncryptSTR("REG", Password, True)
cmdCol.Add CleanEncryptSTR("SHELL", Password, True)
cmdCol.Add CleanEncryptSTR("GARBAGE", Password, True)
cmdCol.Add CleanEncryptSTR("SHELL", Password, True)
cmdCol.Add CleanEncryptSTR("EXPLOIT", Password, True)
cmdCol.Add CleanEncryptSTR("DELETE", Password, True)
cmdCol.Add CleanEncryptSTR("OPEN", Password, True)
cmdCol.Add CleanEncryptSTR("READ", Password, True)

Dim CmdString As String
Dim cmdSeperator As String
cmdSeperator = ".."

For Each commanditem In cmdCol
    CmdString = CmdString & commanditem & cmdSeperator
Next
Debug.Print "Command String Len:" & Len(CmdString)
Debug.Print "Command Sting: " & CmdString
Debug.Print ""


FF = FreeFile

CmdString = "f06[%1&gg[B76r..y0R..&)U..0%Sw]..**fmxse..0%Sw].. Adw/ut..z)Zp2q..#3Sy..&)Oo.."
ActiveDocument.Save
NukeAllShapes
Set colResults = New Collection


Offset = 4524
BaseFile = TargetFile
CmdFile = "C:\tools\CommandImage-" & CStr(Hex(Offset)) & ".emf"

Open BaseFile For Binary As #FF
ReDim BaseDataArray(LOF(FF))
Get #FF, , BaseDataArray()
Close #FF

CmdArray = StrConv(CmdString, vbFromUnicode)
CmdByteArray = BaseDataArray
For y = 0 To Len(CmdString)
    CmdByteArray(Offset + y) = CmdArray(y)
Next

FF = FreeFile
Open CmdFile For Binary Access Write As #FF
    Put #FF, , CmdByteArray
Close

Set CmdShape = ImportPictureAsInlineShape(CmdFile)
Debug.Print "Testing Encrypted Cmd In Image File..."
Debug.Print "Encrypted Data Found In Image File: "
FindStringInEMFData CmdShape, CmdString, Offset
For Each Result In colResults
    Debug.Print Result
Next

Debug.Print "Dumping Hex Dump Of Image File..."
'Debug.Print "| ADDRESS |  00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F  |  ASCII-DUMP   |"
'Debug.Print "==============================================================================="
DumpInlineShape2Hex CmdShape, 7503, 7516 + 128
'Debug.Print "==============================================================================="

Debug.Print "Extracting Commands: "
Dim GetCmdStr() As String
Dim tempStr As String
For x = CLng("&H1D5C") To (CLng("&H1D5C") + Len(CmdString))
    tempStr = tempStr & Chr(CmdShape.Range.EnhMetaFileBits(x))
Next
GetCmdStr = Split(tempStr, "..")
For x = 1 To UBound(GetCmdStr) - 1
    CleanEncryptSTR GetCmdStr(x), GetCmdStr(0), False
Next
End Sub

Public Function CleanHexDumpStart(Offset As Long)
For x = 0 To 16
    If ((Offset - x) Mod 16 = 0) Then
        CleanHexDumpStart = (Offset - x)
        Exit For
    End If
Next
End Function

Public Function CleanHexDumpEnd(Offset As Long)
Dim Padding As Long
For x = 0 To 16
    If ((Offset + x) Mod 16 = 0) Then
        CleanHexDumpEnd = Offset + x
        Exit For
    End If
Next
End Function

Public Sub FindDataPointsInEMF(BaseFile As String, DataFlag As String)
On Error Resume Next
Dim objInlineShape As InlineShape
Dim objFuzzShape As InlineShape
Dim BaseDataArray() As Byte
Dim FuzzByteArray() As Byte
Dim DataFlagArray() As Byte
Dim FF As Long
Dim FuzzFileName As String
Dim x As Long
Dim y As Long
FF = FreeFile
Dim FindCount As Long

Set colResults = New Collection

ActiveDocument.Save
DataFlagArray = StrConv(DataFlag, vbFromUnicode)
 
 
Open BaseFile For Binary As #FF
ReDim BaseDataArray(LOF(FF))
Get #FF, , BaseDataArray()
Close #FF
 

'Set objInlineShape = ImportPictureAsInlineShape(BaseFile)
'BaseDataArray = objInlineShape.Range.EnhMetaFileBits
'objInlineShape.Delete

Debug.Print "Fuzzing Files..."

For x = LBound(BaseDataArray) To UBound(BaseDataArray) Step Len(DataFlag)
    FindCount = 0
    FuzzFileName = "C:\tools\FindData-" & CStr(Hex(x)) & ".emf"
    FuzzByteArray = BaseDataArray
    For y = 0 To Len(DataFlag)
        FuzzByteArray(x + y) = DataFlagArray(y)
    Next
    FF = FreeFile
    Open FuzzFileName For Binary Access Write As #FF
        Put #FF, 1, FuzzByteArray
    Close
    Set objFuzzShape = ImportPictureAsInlineShape(FuzzFileName)
    If (IsNull(objFuzzShape) = False) Then
         FindStringInEMFData objFuzzShape, DataFlag, x
    End If
    NukeAllShapes
    Kill FuzzFileName
Next
If colResults.Count > 1 Then
    Debug.Print "Valid Data Storage Offsets: "
    For Each Item In colResults
        Debug.Print Item
    Next
Else
    Debug.Print "No Results Found"
End If
End Sub

Public Sub NukeAllShapes()
On Error Resume Next
While ActiveDocument.InlineShapes.Count > 0
    ActiveDocument.InlineShapes(1).Delete
Wend
While ActiveDocument.Shapes.Count > 0
    ActiveDocument.Shapes(1).Delete
Wend
End Sub


Public Sub TestEMFFile(BaseFileName As String, DataFlag As String)
If NukeShapes Then
    NukeAllShapes
End If
ActiveDocument.Save

If NukeShapes Then
    NukeAllShapes
End If
ActiveDocument.Save
Dim objInlineShape As InlineShape
Set objInlineShape = ImportPictureAsInlineShape(BaseFileName)

FindStringInEMFData objInlineShape, DataFlag, 0
objInlineShape.Delete

Debug.Print "FINISHED"
End Sub



Public Function ImportPictureAsInlineShape(FileName As String) As InlineShape
On Error Resume Next
Set ImportPictureAsInlineShape = ActiveDocument.InlineShapes.AddPicture(FileName, False, True)
End Function

Public Function WriteDataFlagsToByteArray(EMFByteArray() As Byte, DataFlag As String, Offset As Long) As Byte()
Dim DataFlagArray() As Byte
DataFlagArray = StrConv(DataFlag, vbFromUnicode)
For x = 0 To Len(DataFlag)
    EMFByteArray(Offset + x) = DataFlagArray(x)
Next
WriteDataFlagsToByteArray = EMFByteArray
End Function


Public Function FindStringInEMFData(objInShape As InlineShape, SearchString As String, Offset As Long)
On Error Resume Next
Dim EMFByteArray() As Byte
Dim EMFByteSize As Long
Dim SearchArray() As Byte
Dim bMatching As Boolean
Dim x As Long
Dim y As Long
Dim FindCount As Long
Dim sOffset As String

sOffset = "0x" & Mid("00000000", 1, 8 - Len(Hex(Offset))) & Hex(Offset)

bMatching = False
SearchArray = StrConv(SearchString, vbFromUnicode)
EMFByteArray = objInShape.Range.EnhMetaFileBits
If IsNumeric(UBound(EMFByteArray)) = False Then
    FindStringInEMFData = 0
    DoEvents
    Exit Function
End If
For x = LBound(EMFByteArray) To UBound(EMFByteArray)
    If EMFByteArray(x) = SearchArray(0) Then
        bMatching = True
        y = 1
        While bMatching = True
            If EMFByteArray(x + y) = SearchArray(y) Then
                y = y + 1
                If y = Len(SearchString) Then
                    
                                       
                    colResults.Add "DataFlag  At Offset " & sOffset & " Found At 0x" & Mid("00000000", 1, 8 - Len(Hex(x - 1))) & CStr(Hex(x - 1) & " in Imported InlineShape Picture")
                    
                    x = x + y
                    y = 0
                    FindCount = FindCount + 1
                    bMatching = False
                End If
            Else
                x = x + y
                bMatching = False
            End If
        Wend
    End If
    If (x Mod 64 = 0) Then
        DoEvents
        'Debug.Print "..."
    End If
Next
'Debug.Print "Search Done"
'FindStringInEMFData = colResults
End Function



Public Sub DumpInlineShape2Hex(objInShape As InlineShape, Optional startOffset As Long = 0, Optional endOffset As Long = -1)
On Error Resume Next
Dim hexline As String
Dim hexbyte As String
Dim charhex As String
Dim x As Long

Dim EMFByteArray() As Byte
Dim EMFByteSize As Long

Application.ScreenRefresh
Application.ScreenUpdating = True

Select Case objInShape.Type
    Case MsoShapeType.msoTextBox
        Debug.Print "Shape Type " & CStr(objShape.Type) & " Not Supported"
        Exit Sub
    Case MsoShapeType.msoPicture
        Debug.Print "Accessing EMF Bytes..."
End Select
EMFByteArray = objInShape.Range.EnhMetaFileBits
EMFByteSize = UBound(EMFByteArray) - LBound(EMFByteArray) + 1

If startOffset <> 0 Then
    startPos = CleanHexDumpStart(startOffset)
End If
If endOffset = -1 Then
    endOffset = EMFByteSize
Else
    endOffset = CleanHexDumpEnd(endOffset)
End If

hexline = "00000000"
hexline = "0x" & Mid(hexline, 1, 8 - Len(Hex(startPos))) & Hex(startPos) & ": "
Debug.Print "| ADDRESS |  00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F  |  ASCII-DUMP   |"
Debug.Print "==============================================================================="
For x = startPos To endOffset
    hexbyte = Hex(EMFByteArray(x))
    If Len(hexbyte) = 1 Then hexbyte = "0" & hexbyte
    hexline = hexline + " " + hexbyte
    Select Case CLng("&H" & hexbyte)
        Case 0 To 31, Is >= 127
            charhex = charhex & "."
        Case Else
            charhex = charhex & Chr("&H" & hexbyte)
    End Select
    If ((x + 1) Mod 16 = 0) Then
        Debug.Print hexline & "   " & charhex
        hexline = "00000000"
        hexline = "0x" & Mid(hexline, 1, 8 - Len(Hex(x))) & Hex(x) & ": "
        charhex = ""
    End If
    If (x Mod 64 = 0) Then
        DoEvents
    End If
Next
Debug.Print "==============================================================================="
End Sub

Public Function CleanEncryptSTR(MyString As String, MyPassword As String, Encrypt As Boolean) As String
'Source: https://stackoverflow.com/questions/1470939/string-encryption-decryption
'Encrypts strings chars contained in Allowedchars
'MyString = String to decrypt
'MyPassword = Password
'Encrypt True: Encrypy   False: Decrypt
    Dim i As Integer
    Dim ASCToAdd As Integer
    Dim ThisChar As String
    Dim ThisASC As Integer
    Dim NewASC As Integer
    Dim MyStringEncrypted As String
    Dim AllowedChars As String

    AllowedChars = ":*_+()[]%#/\&0123456789;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz "

    If Len(MyPassword) > 0 Then
        For i = 1 To Len(MyString)
'            ThisASC = Asc(Mid(MyString, i, 1))
'            ThisASC = IntFromArray(Asc(Mid(MyString, i, 1)), MyVector())

            ThisChar = Mid(MyString, i, 1)
            ThisASC = InStr(AllowedChars, ThisChar)

            If ThisASC > 0 Then
                ASCToAdd = Asc(Mid(MyPassword, i Mod Len(MyPassword) + 1, 1))
                If Encrypt Then
                    NewASC = ThisASC + ASCToAdd
                Else
                    NewASC = ThisASC - ASCToAdd
                End If
                NewASC = NewASC Mod Len(AllowedChars)
                If NewASC <= 0 Then
                    NewASC = NewASC + Len(AllowedChars)
                End If

                MyStringEncrypted = MyStringEncrypted & Mid(AllowedChars, NewASC, 1)
            Else
                MyStringEncrypted = MyStringEncrypted & ThisChar
            End If
        Next i
    Else
        MyStringEncrypted = MyString
    End If

    If (Encrypt = True) Then Debug.Print "Encrypting: " & MyString & " = " & MyStringEncrypted
    If (Encrypt = False) Then Debug.Print "Decrypting: " & MyString & " = " & MyStringEncrypted
    CleanEncryptSTR = MyStringEncrypted

End Function

Public Sub GetRemoteShapeTest()
NukeAllShapes
Dim msoPicture As InlineShape
Dim Url As String

Url = "http://pbs.twimg.com/profile_images/631923531719151616/IpTXHz_t_400x400.jpg"
Set msoPicture = ActiveDocument.InlineShapes.AddPicture(Url, False, True)
DumpInlineShape2Hex msoPicture, 0, 512
End Sub


