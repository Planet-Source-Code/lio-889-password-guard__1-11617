Attribute VB_Name = "SubMain"
' =================================================================
' Password Guard source code
' Version 1.1
' Copyright (C) 2000 Khaery Rida
' =================================================================

' Thanx for using Password Guard!
' Please log on http://www.geocities.com/lio889 for more great VB programs!
' Comments or Questions? Please do NOT hesitate at emailing me:
' lio_889@ziplip.com

' * AutoSave feature ensures that all the changes made to a specific record by the end-user, are
' immediately considered.

' Declare Windows' API functions
Public Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCursorPos Lib "User32" (ByRef lpPoint As Point) As Long
Public Declare Function WindowFromPoint Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Global Constants
Global Const MainTitle = "Password Guard"
Global Const MasterKey = "PGmECk"
Global Const Key1 = "PswrdGrd"
Global Const Key2 = "CherAlog"
Global Const FileTitle = "Password Gaurd Data File"
Global Const conKey = ""
Global Const sDivide = "ƒ"

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

' Color Constants
Global Const LightOn = &HC0&
Global Const LightOff = &H800000
Global Const Active = &H80000005
Global Const NoActive = &H8000000F

Global Const Navy = &H800000
Global Const SRCCOPY = &HCC0020

' Message Box Constants
Global Const MB_YESNO = 4
Global Const MB_ICONQUESTION = 32
Global Const MB_DEFBUTTON1 = &H0&
Global Const MB_DEFBUTTON2 = 256
Global Const IDYES = 6

' User-defined types
Public Type DataRecord
    Description As String
    Server As String
    UserName As String
    Password As String
    Notes As String
End Type

Public Type SearchRecord
    keyName As String * 8
End Type

Public Type Point
    X As Long
    Y As Long
End Type

' Public Variables
Public tDescription As String
Public tServer As String
Public tUserName As String
Public tPassword As String
Public tNotes As String

Public currentRecord As Long
Public currentKey As String
Public Saved As Boolean

Public currentIndex As Long
Public FileNum As Integer
Public TestFileNum As Integer
Public ItemCount As Long
Public MainList As ListItem
Public Keyword As String

Public regString As String
Public regInt As Integer

Public tmpString As String
Public tmpString2 As String
Public tmpString3 As String

Public Index As String
Public sourcePath As String

Public UserKeyword As String
Public UserIDKeyword As String
Public MasterPasswordKeyword As String

Public UserIndex As String
Public ItemIndex As String

Public UserRegSection As String
Public UserID As String
Public MasterPassword As String

Public DgDef, Msg, Response, Title


Public Sub Main()

' Load forms:
    Load frmLogIn
    Load frmMain
    Load frmUserID
    Load frmAbout
    
    Dim cIndex As String
    Dim cUserIDIndex As String
    Dim cOptionsIndex As String
    Dim cOptions As String
    
    sourcePath$ = CurDir$ & "\"
    frmLogIn.cmdSubmit.Enabled = False
    regString = GetSetting(MainTitle, "Settings", "Index", "/NEWRUN/")
    
    If regString = "/NEWRUN/" Then
        Index = "/NEWRUN/"
        UserCount = 0
        frmLogIn.Show
        Exit Sub
    End If
    
    Index = decrypt(regString, Key1 & Key2)
    
    For currentIndex = 1 To Len(Index) Step 6
        UserCount = UserCount + 1
        UserRegSection = Mid$(Index, currentIndex, 6)
        regString = GetSetting(MainTitle, UserRegSection, "Index")
        cIndex = decrypt(regString, Key2 & UserRegSection)
        cOptionsIndex = Mid$(cIndex, 17, 8)
        regString = GetSetting(MainTitle, UserRegSection, cOptionsIndex)
        cOptions = decrypt(regString, Key1 & Mid$(cOptionsIndex, 3, 1) & Mid$(cOptionsIndex, 5, 1))
        
        If Left$(cOptions, 1) = "1" Then
            cUserIndex = Left$(cIndex, 8)
            regString = GetSetting(MainTitle, UserRegSection, cUserIndex)
            frmLogIn.lstUserID.AddItem decrypt(regString, Left$(UserRegSection, 2) & Right$(cUserIndex, 2) & Key1)
        End If
    Next
    
    frmLogIn.lstUserID.Text = ""
    frmLogIn.Show
    Exit Sub
    
End Sub
Public Function FileExists(Path$) As Integer

' This function is used to ensure that a file is openable.
  
    X = FreeFile

    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X

End Function
Public Function IsValidUserID(UserID As String) As String
    
    Dim vUserIndex As String
    Dim vUserRegSection As String
    
    For currentIndex = 1 To Len(Index) Step 6
        vUserRegSection = Mid$(Index, currentIndex, 6)
        regString = GetSetting(MainTitle, vUserRegSection, "Index")
        vUserIndex = decrypt(regString, Key2 & vUserRegSection)
        regString = GetSetting(MainTitle, vUserRegSection, Left$(vUserIndex, 8))
        UserIDKeyword = Left$(vUserRegSection, 2) & Right$(Left$(vUserIndex, 8), 2) & Key1
        
        If decrypt(regString, UserIDKeyword) = UserID Then
            IsValidUserID = vUserRegSection & vUserIndex
            Exit Function
        End If
    Next
    
    IsValidUserID = ""
    
End Function

Public Function IsValidMasterPassword(regSection As String, regKey As String, UserID As String, MasterPassword As String) As String

    regString = GetSetting(MainTitle, regSection, regKey)
    MasterPasswordKeyword = Mid$(regSection, 3, 1) & Left$(UserID, 1) & Left$(MasterPassword, Len(MasterPassword) - 5) & Right$(MasterPassword, 1) & Right$(regKey, 1) & Right$(UserID, 1)
    tmpString = decrypt(regString, MasterPasswordKeyword)
    
    If Trim(MasterPassword) = tmpString Then
        IsValidMasterPassword = Left$(UserID, 1) & Mid$(regSection, 5, 2) & MasterPassword
    Else
        IsValidMasterPassword = ""
    End If
    
End Function

Public Sub LogIn()
    Screen.MousePointer = 11
    If Not frmUserID.txtUserID.Text = "" Then GoTo NoUserIDFormSet
    Load frmUserID
    
    ' Fill in frmUserID
    frmUserID.txtUserID.Text = UserID
    frmUserID.txtUserID.Tag = Left$(UserIndex, 8)
    frmUserID.txtUserID.Locked = True
    
    frmUserID.txtMasterPassword1.Text = MasterPassword
    frmUserID.txtMasterPassword2.Text = MasterPassword
    frmUserID.txtMasterPassword1.Tag = Mid$(UserIndex, 9, 8)
    
    frmUserID.frameLog.Tag = Mid$(UserIndex, 17, 8)
    frmUserID.txtQuestion.Tag = Mid$(UserIndex, 25, 8)
    frmUserID.txtAnswer.Tag = Mid$(UserIndex, 33, 8)
    
    frmUserID.txtUserID.Text = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.txtUserID.Tag), UserIDKeyword)
    frmUserID.txtMasterPassword1.Text = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.txtMasterPassword1.Tag), MasterPasswordKeyword)
    frmUserID.txtMasterPassword2.Text = frmUserID.txtMasterPassword1.Text
    frmUserID.txtQuestion.Text = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.txtQuestion.Tag), Key1 & Mid$(frmUserID.txtQuestion.Tag, 3, 1) & Mid$(frmUserID.txtQuestion.Tag, 5, 1))
    frmUserID.txtAnswer.Text = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.txtAnswer.Tag), Key1 & Mid$(frmUserID.txtAnswer.Tag, 3, 1) & Mid$(frmUserID.txtAnswer.Tag, 5, 1))
    regString = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.frameLog.Tag), Key1 & Mid$(frmUserID.frameLog.Tag, 3, 1) & Mid$(frmUserID.frameLog.Tag, 5, 1))
    frmUserID.chkDisplay.Value = Left$(regString, 1)
    frmUserID.chkPassword.Value = Mid$(regString, 2, 1)
    frmUserID.chkRemove.Value = Mid$(regString, 3, 1)
    frmUserID.chkLog.Value = Mid$(regString, 4, 1)
    frmUserID.chkLogAll.Value = Mid$(regString, 5, 1)
    frmUserID.chkEncrypt.Value = Mid$(regString, 6, 1)
    
    frmUserID.chkLog_Click
    
    If Len(regString) > 6 Then frmUserID.txtLogFile.Text = Mid$(regString, 7)
    
NoUserIDFormSet:
   ItemCount = 0
   regString = GetSetting(MainTitle, UserRegSection, "Item", "")
   If Not regString = "" Then ItemIndex = decrypt(regString, UserKeyword) Else ItemIndex = ""
   
   frmMain.lstItem.ListItems.Clear
   frmMain.txtDescription.Text = ""
   frmMain.txtServer.Text = ""
   frmMain.txtUserName.Text = ""
   frmMain.txtPassword.Text = ""
   frmMain.txtNotes.Text = ""
   
   If Len(ItemIndex) = 0 Then
        frmMain.txtDescription.Enabled = False
        frmMain.txtServer.Enabled = False
        frmMain.txtUserName.Enabled = False
        frmMain.txtPassword.Enabled = False
        frmMain.txtNotes.Enabled = False
        frmMain.mnuRemoveRecord.Enabled = False
        frmMain.mnuSearch.Enabled = False
        frmMain.imgRemove.Visible = False
        frmMain.lblRemove.Visible = False
        frmMain.imgSearch.Visible = False
        frmMain.lblSearch.Visible = False
        
        GoTo NoItems
    End If
    
    Dim curKey As Long
    Dim KeyString As String
    Dim tmp1stRecord As DataRecord
    
    ' Add all items to the ListView control
    For curKey = 1 To Len(ItemIndex) Step 8
        KeyString = Mid$(ItemIndex, curKey, 8)
        regString = GetSetting(MainTitle, UserRegSection, KeyString)
        tmpString = decrypt(regString, UserKeyword & Mid$(KeyString, 3, 1) & Mid$(KeyString, 5, 1))
        currentChr = 0
        tmpString2 = ""
        Do Until tmpString2 = sDivide
            currentChr = currentChr + 1
            tmpString2 = Mid$(tmpString, currentChr, 1)
        Loop
        Set MainList = frmMain.lstItem.ListItems.Add(, KeyString, Left$(tmpString, currentChr - 1), 1, 1)
        ItemCount = ItemCount + 1
    Next
    
    tmp1stRecord = ReadRecord(Left$(ItemIndex, 8))
    tDescription = tmp1stRecord.Description
    tServer = tmp1stRecord.Server
    tUserName = tmp1stRecord.UserName
    tPassword = tmp1stRecord.Password
    tNotes = tmp1stRecord.Notes
    
    frmMain.txtDescription.Text = tDescription
    frmMain.txtServer.Text = tServer
    frmMain.txtUserName.Text = tUserName
    frmMain.txtPassword.Text = tPassword
    frmMain.txtNotes.Text = tNotes
            
    frmMain.lstItem.ListItems(1).Selected = True
    currentKey = Left$(ItemIndex, 8)
    currentRecord = 1
    Saved = True
    
NoItems:
    frmMain.Caption = MainTitle & " - [Welcome " & UserID & "]"
    Screen.MousePointer = 0
    frmMain.Show
    
End Sub


Public Function RandomPinString(PinNum As Integer) As String
    
    Dim tOffset As Integer, currentPin As Long, tmpPin As String
    Dim PinNumCount As Long
    
GenerateRndPinString:
        Randomize
        For currentPin = 1 To PinNum
            tOffset = (Rnd * 10000 Mod 255) + 1
            RandomPinString = RandomPinString & Format$(Hex$(tOffset), "@@")
        Next
        
        ' The Format$ function is used to make sure that always 2 bytes are returned.
        ' For example, instead of returning "B", the Format$ function, returns "B "
        ' in this way, the resulting RandomPinString will always consist of 8 characters.
        
        PinNumCount = 0
        
        For currentPin = 1 To Len(RandomPinString)
            tmpPin = Mid$(RandomPinString, currentPin, 1)
            If IsNumeric(tmpPin) Then PinNumCount = PinNumCount + 1
        Next
        
        If PinNumCount = (PinNum * 2) Then GoTo GenerateRndPinString
        ' Since this RandomPinString will be used as a ListItem key, we need to ensure that the
        ' result RadnomPinString is not a whole numeric value. For example, a RandomPinString
        ' may take te value 24982751 which is an invalid ListItem's key.
       
End Function

Public Function ObtainFilePassword(uAction As String, File As String) As String
    
    Load frmKeyword
    frmKeyword.txtKeyword.Text = ""
    frmKeyword.cmdOK.Enabled = False
    
    If uAction = "Export" Then
        frmKeyword.lblAction.Caption = "lock"
        frmKeyword.Caption = "Export Records to File"
        frmKeyword.img.Picture = frmMain.imgExport.Picture
    ElseIf uAction = "Import" Then
        frmKeyword.lblAction.Caption = "unlock"
        frmKeyword.Caption = "Import Records from File"
        frmKeyword.img.Picture = frmMain.imgImport.Picture
    End If
    
    frmKeyword.lblFile.Caption = LCase$(File)
    frmKeyword.img.Height = frmKeyword.img.Height - 20
    frmKeyword.img.Stretch = True
    frmKeyword.Show 1
    ObtainFilePassword = Trim(frmKeyword.txtKeyword.Text)
    Unload frmKeyword
    
End Function
Public Sub UpdateProgress(PositionSetBack As Long, StringLength As Long)
    
' This function is used to update the progress bar while an operation is in progress.
    
    Static position
    Dim txt As String, r As Long, estTotal As Long
    estTotal = frmMain.picProgress.Tag
    
    If PositionSetBack = 1 Then position = 0

    position = position + CSng((StringLength / estTotal) * 100)
    If position > 100 Then
        position = 100
    End If
    
    txt$ = Format$(CLng(position)) + "%"
    
    frmMain.picProgress.Line (0, 0)-((position * (frmMain.picProgress.ScaleWidth / 100)), frmMain.picProgress.ScaleHeight), Navy, BF
    frmMain.picProgress.CurrentX = (frmMain.picProgress.ScaleWidth - frmMain.picProgress.TextWidth(txt$)) \ 2
    frmMain.picProgress.CurrentY = (frmMain.picProgress.ScaleHeight - frmMain.picProgress.TextHeight(txt$)) \ 2
    r = BitBlt(frmMain.picProgress.hDC, 0, 0, frmMain.picProgress.ScaleWidth, frmMain.picProgress.ScaleHeight, frmMain.picProgress.hDC, 0, 0, SRCCOPY)

End Sub
Public Function ReadRecord(regKey As String) As DataRecord
    
    Dim memberCount As Long
    Dim memberLen As Long
    Dim LastPos As Long

    regString = GetSetting(MainTitle, UserRegSection, regKey)
    tmpString = decrypt(regString, UserKeyword & Mid$(regKey, 3, 1) & Mid$(regKey, 5, 1))
    memberCount = 0
    memberLen = 0
    LastPos = 0
    
    For currentChr = 1 To Len(tmpString)
        tmpString2 = Mid$(tmpString, currentChr, 1)
        memberLen = memberLen + 1
        
        If tmpString2 = sDivide Then
            memberCount = memberCount + 1
            
            Select Case memberCount
                Case 1
                    ReadRecord.Description = Left$(tmpString, currentChr - 1)
                Case 2
                    ReadRecord.Server = Mid$(tmpString, LastPos, memberLen - 1)
                Case 3
                    ReadRecord.UserName = Mid$(tmpString, LastPos, memberLen - 1)
                Case 4
                    ReadRecord.Password = Mid$(tmpString, LastPos, memberLen - 1)
                Case 5
                    ReadRecord.Notes = Mid$(tmpString, LastPos, memberLen - 1)
            End Select
            
            memberLen = 0
            LastPos = currentChr + 1
        End If
    Next
    
End Function
