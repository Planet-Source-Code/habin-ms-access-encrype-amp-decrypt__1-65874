Attribute VB_Name = "modAccessDBEncryptDecrypt"
Global arrEncrypt()    As Byte
Global arrDecrypt()    As Byte
Global i               As Integer
Global strEncrypt      As String
Global intEncrypt      As Integer
Global strTemp


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, ByVal lpszOp As String, _
                 ByVal lpszFile As String, ByVal lpszParams As String, _
                 ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                 As Long

Public Sub StartEncDec()
strEncrypt = "HHJRSOFTCOMPANY."

ReDim arrEncrypt(Len(strEncrypt) - 1)


For i = 0 To Len(strEncrypt) - 1
    intEncrypt = Asc(Mid(strEncrypt, i + 1, 1))
    arrEncrypt(i) = intEncrypt
Next


ReDim arrDecrypt(0 To 15)

strTemp = Split("0,1,0,0,83,116,97,110,100,97,114,100,32,74,101,116", ",")

For i = 0 To 15
    
    arrDecrypt(i) = strTemp(i)
    
Next
End Sub


Public Sub DecryptMDB(filename As String)

Dim frFile As Integer
Dim bytes() As Byte
    
    
    frFile = FreeFile
    
    If Dir(filename) = "" Then Exit Sub
    
    Open filename For Binary As #frFile Len = 16
    
    
    ReDim bytes(0 To 15)
    
   
    Put #frFile, 1, arrDecrypt
    
    Close #frFile
    
End Sub


Public Sub EncryptMDB(filename As String)

Dim frFile As Integer
Dim bytes() As Byte
    
    
    frFile = FreeFile
    
    If Dir(filename) = "" Then Exit Sub
    
    Open filename For Binary As #frFile Len = 16
    
    
    ReDim bytes(0 To 15)
    

    
    Put #frFile, 1, arrEncrypt
    
    Close #frFile
    
End Sub
Sub CloseAll()
    On Error Resume Next
    Dim intFrmNum As Integer
    intFrmNum = Forms.Count


    Do Until intFrmNum = 0
        Unload Forms(intFrmNum - 1)
        intFrmNum = intFrmNum - 1
    Loop
End Sub
