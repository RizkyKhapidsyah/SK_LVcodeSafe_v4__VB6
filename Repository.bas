Attribute VB_Name = "LibFunctions"
Option Explicit
Option Compare Text

Private Type UserDefaults               ' configure a Type variable to be used within this program
    GradientPopup As Integer
    GradientMain As Integer
    LblColorMain As Long
    LblColorPopup As Long
    WindowColor As String
    PopupColor As String
    ButtonSize As Integer
    Language As Long
    DBname As String
    FontSize As String
    Font As String
    Align As Integer
    KeyWd1 As Long
    KeyWd2 As Long
    KeyWd3 As Long
    KeyWd4 As Long
End Type
Public DBrecID As Variant               ' record id passed from main window to child forms
Public mainDB As DAO.Database    ' shared database for all forms
Public mainRS As DAO.Recordset   ' share recordset for main form & edit/add record form
Private Const ProgramDefaultColor   As String = "8421376"     ' default gradient color
Public Const WinBlahColor               As String = "-12345"       ' code to prevent gradient coloring
Public mainFilterIndex As String        ' Filter string for filtering mainRS
Public LastCatUpdate As Date            ' Date/time Category table updated via this program
Public LastLangUpdate As Date           ' Date/time Language table updated via this program
Public bAppClose As Boolean             ' flag to prevent program shutdown if forms have unsaved changes
Public GP As Variant                    ' General purpose variable used throughout
Public rsAttachment As DAO.Recordset    ' shared recordset when modifying/viewing attachments
Public MyDefaults As UserDefaults   ' Public variable of user-defined type above
' Function which returns information on a window's text box
Public Declare Function apiSendMessageS Lib "user32" _
                         Alias "SendMessageA" _
                        (ByVal hWnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         lparam As String) As Long
Public Declare Function apiSendMessage Lib "user32" _
                         Alias "SendMessageA" _
                        (ByVal hWnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         lparam As Long) As Long

Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Private Const EM_REPLACESEL = &HC2

' Local variables used in program
Private Const RmksColor = &H8000&              ' default Remarks color

Sub Main()
'=======================================================================
' Main Sub - initializes connection to database, initially defines the mainRS recordset and default settings
'   this routine is also called when a user changes the master database from the Options window
'=======================================================================
' Inserted by LaVolpe
On Error GoTo Sub_Main_General_ErrTrap_by_LaVolpe
    Screen.MousePointer = vbHourglass
    ' get user defaults for window colors, providing a default if not previously defined
    MyDefaults.WindowColor = GetSetting("LaVolpeCodeSafe", "Settings", "MainColor", "8421376")
    MyDefaults.PopupColor = GetSetting("LaVolpeCodeSafe", "Settings", "SecondaryColor", "65280")
    MyDefaults.GradientMain = CInt(GetSetting("LaVolpeCodeSafe", "Settings", "GradientMain", "0"))
    MyDefaults.GradientPopup = CInt(GetSetting("LaVolpeCodeSafe", "Settings", "GradientPopUp", "0"))
    MyDefaults.LblColorMain = CLng(GetSetting("LaVolpeCodeSafe", "Settings", "LblColorMain", "&H00FFFFFF"))
    MyDefaults.LblColorPopup = CLng(GetSetting("LaVolpeCodeSafe", "Settings", "LblColorPopUp", "&H00FFFFFF"))
    MyDefaults.Font = GetSetting("LaVolpeCodeSafe", "Settings", "FontType", "Times New Roman")
    MyDefaults.FontSize = GetSetting("LaVolpeCodeSafe", "Settings", "FontSize", "11")
    MyDefaults.Align = CInt(GetSetting("LaVolpeCodeSafe", "Settings", "Alignment", "3"))
    MyDefaults.KeyWd1 = CLng(GetSetting("LaVolpeCodeSafe", "Settings", "CompilerColor", "10485760"))
    MyDefaults.KeyWd2 = CLng(GetSetting("LaVolpeCodeSafe", "Settings", "FunctionsColor", "8404992"))
    MyDefaults.KeyWd3 = CLng(GetSetting("LaVolpeCodeSafe", "Settings", "MiscColor", "8421440"))
    MyDefaults.KeyWd4 = CLng(GetSetting("LaVolpeCodeSafe", "Settings", "PropertiesColor", "0"))
    On Error Resume Next
    GP = Null                       ' call the db connection form and wait until it is finished before continuing
    Load frmDBprompt
    Do Until IsNull(GP) = False
        DoEvents                    ' keep waiting
    Loop
    ' ok, db connection form closed & now let's see if a database was selected
    If mainDB Is Nothing Then   ' nope, gotta shut down
        End
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    frmLibrary.Show                                         ' and show it
    frmLibrary.StatusBar1.Panels(2) = MyDefaults.DBname
    If FileLen(MyDefaults.DBname) / 1024 > 1000 Then
        frmLibrary.StatusBar1.Panels(4) = Format(FileLen(MyDefaults.DBname) / 1024000#, "#.00 mb")
    Else
        frmLibrary.StatusBar1.Panels(4) = Format(FileLen(MyDefaults.DBname) / 1024#, "#.00 kb")
    End If
    mainFilterIndex = "[All Categories]|0|[With & Without]|0|[All Languages]|0|"    ' default filter
    RefreshCategories                                       ' load categories to main form
    RefreshLanguages                                       ' load languages to main form
    frmLibrary.cboFilter(1).ListIndex = 0           ' set attachment option in filter to default
    frmLibrary.cboFilter(0).ListIndex = 0
    FilterRecordset                                            ' load and filter recordset
    RefreshOtherApps
    frmLibrary.cmdFilterDefault(2).Tag = ""
    Screen.MousePointer = vbDefault                ' done
Exit Sub

Sub_Main_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Main]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub RefreshCodeList()
'=======================================================================
'   Function is called whenever by FilterRecordset function and simply populates the main form's code list
'=======================================================================
Dim intI As Integer, rsAttach As DAO.Recordset
Dim strSQL As String, itmX As ListItem, iIcon As Integer

On Error GoTo CloseRecordset
 frmLibrary.lvCode.ListItems.Clear      ' clear the listing & construct query string to return those code entries with attachments
    strSQL = "SELECT DISTINCT tblSourceCode.IDnr FROM tblAttachments " & _
    "INNER JOIN tblSourceCode ON tblAttachments.RecIDRef = tblSourceCode.IDnr;"
Set rsAttach = mainDB.OpenRecordset(strSQL, dbOpenDynaset)
DBrecID = Null
With frmLibrary.lvCode.ListItems          ' add each record in the mainRS recordset & include an icon if it has an attachment
    If mainRS.RecordCount > 0 Then      ' loop thru each item in the code list (via the mainRS recordset)
        mainRS.MoveFirst
        Do While mainRS.EOF = False     ' and see if an attachment exists for that code
            iIcon = 0                                   ' initially set icon flag to 0 (no icon)
            If rsAttach.RecordCount Then   ' now loop thru each record in the attachment table to see if there's a match
                rsAttach.FindFirst "[IDnr] = " & mainRS.Fields("IDnr")  ' gotta match?
                If rsAttach.NoMatch = False Then iIcon = 13                ' if so set the icon flag to attachment icon reference
            End If
            ' Add an entry into the code listing
            '   All entries include the code name, db record ID as "RecID:###"
            Set itmX = .Add(, "RecID:" & CStr(mainRS.Fields("IDnr")), mainRS.Fields("CodeName"), , iIcon)
            If IsNull(DBrecID) Then DBrecID = mainRS.Fields("IDnr") ' track first record to select later
            mainRS.MoveNext
        Loop
    End If
    rsAttach.Close                      ' finished loading display, so close attachment recordset
    Set rsAttach = Nothing
    Set itmX = Nothing              ' reset variable
    If .Count = 0 Then               ' do we have any items in the code listing? if not, display a warning
        If mainFilterIndex <> "[All Categories]|[With & Without]|[All Languages]|" Then
            MsgBox "The filter you applied now excludes any code in your database or your database is blank.", vbInformation + vbOKOnly
        End If
    End If
End With
CloseRecordset:
If Err.Number Then
    If Err.Number = 35602 Then
        Err.Clear
        Resume Next
    End If
    On Error Resume Next
    mainRS.Close
    Set mainRS = Nothing
    MsgBox "Choose another database..." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End If
End Sub

Function ConvertToRGB(HexLng As String) As String
'=======================================================================
'   'This will convert Hexidecimal color coding to RGB color coding
'       variables passed must either be numeric or in the format of &H########
'=======================================================================
' Inserted by LaVolpe
On Error GoTo Function_ConvertToRGB_General_ErrTrap_by_LaVolpe
If IsNumeric(Mid(HexLng, 2)) Then
    If Val(HexLng) < 0 Then HexLng = 255
    HexLng = BigDecToHex(HexLng)
End If
'For Convert Hexidecimal to RGB:  Converts Hexidecimal to RGB
On Error GoTo errorsub
Dim Tmp$
Dim lo1 As Integer, lo2 As Integer
Dim hi1 As Long, hi2 As Long
Const Hx = "&H"
Const BigShift = 65536
Const LilShift = 256, Two = 2
Tmp = HexLng
If UCase(Left$(HexLng, 2)) = "&H" Then Tmp = Mid$(HexLng, 3)
Tmp = Right$("0000000" & Tmp, 8)
If IsNumeric(Hx & Tmp) Then
lo1 = CInt(Hx & Right$(Tmp, Two))       ' Red
hi1 = CLng(Hx & Mid$(Tmp, 5, Two))   ' Green
lo2 = CInt(Hx & Mid$(Tmp, 3, Two))     ' blue
hi2 = CLng(Hx & Left$(Tmp, Two))
'ConvertToRGB = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
ConvertToRGB = Format(lo1, "000") & Format(hi1, "000") & Format(lo2, "000")
End If
Exit Function

errorsub:  MsgBox Err.Description, vbExclamation, "Error"
Exit Function

Function_ConvertToRGB_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function ConvertToRGB]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Private Function BigDecToHex(ByVal DecNum) As String
'=======================================================================
'   Used to convert any decimal value to hex equivalent
'=======================================================================
    ' This function is 100% accurate untill
    '     15,000,000,000,000,000 (1.5E+16)
    Dim NextHexDigit As Double
    Dim HexNum As String

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo BigDecToHex_General_ErrTrap

HexNum = ""
While DecNum <> 0
    NextHexDigit = DecNum - (Int(DecNum / 16) * 16)
    If NextHexDigit < 10 Then
        HexNum = Chr(Asc(NextHexDigit)) & HexNum
    Else
        HexNum = Chr(Asc("A") + NextHexDigit - 10) & HexNum
    End If
    DecNum = Int(DecNum / 16)
Wend

If HexNum = "" Then HexNum = "0"
BigDecToHex = HexNum
Exit Function

' Inserted by LaVolpe OnError Insertion Program.
BigDecToHex_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: BigDecToHex" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function

Public Function DoGradient(ObjName As Object, levelID As Integer, _
    Optional bDontBypass As Boolean = False) As Boolean
'=======================================================================
'   'This will color any object with a AutoRedraw property in gradient colors
'   Required variables
'   - ObjName: this is the object to repaint, generally a form or picture box
'   - LevelID: this determines which color gets applied (main window level, or popup form level)
'   - bDontBypass: optional variable to force value in ProgramDefaultColor to be applied
'   - The color value of -12345 means no coloring
'=======================================================================
On Error Resume Next
Dim sColor As String, bSolidsOnly As Boolean
If levelID = 1 Then             ' Main window vs popup
    If MyDefaults.WindowColor = WinBlahColor Then   ' if no color is being requested, then
        If bDontBypass = False Then Exit Function            ' if object wants above request bypassed, bypass it
        sColor = ProgramDefaultColor                              ' otherwise exit
    Else
        sColor = MyDefaults.WindowColor                       ' a color value is passed, so use it
        bSolidsOnly = CBool(MyDefaults.GradientMain)
    End If
Else
    If MyDefaults.PopupColor = WinBlahColor Then        ' same logic as above but used with popup color options
        If bDontBypass = False Then Exit Function
        sColor = ProgramDefaultColor
    Else
        sColor = MyDefaults.PopupColor
        bSolidsOnly = CBool(MyDefaults.GradientPopup)
    End If
End If
If bDontBypass = True Then bSolidsOnly = False
    
    Dim I As Integer, y As Integer, x As Integer, Z As Integer
    Dim R As Integer, B As Integer, G As Integer
    Dim Red As Integer, Blue As Integer, Green As Integer
    ObjName.BackColor = CLng(sColor)
    If bSolidsOnly = False Then
        ObjName.AutoRedraw = True                                       ' set this variable & if it returns an error,
        If Err.Number > 0 Then                                                  ' then we can't color it this way
            Err.Clear
            Exit Function
        End If
    Else
        DoGradient = True
        Exit Function
    End If
    sColor = ConvertToRGB(sColor)         ' Send hex or decimal number to be converted to an RGB string
    Red = Val(Left(sColor, 3))                  ' Set the red color (0-255)
    Green = Val(Mid(sColor, 4, 3))           ' set the green color
    Blue = Val(Right(sColor, 3))                ' set the blue color
    x = ObjName.ScaleHeight                  ' keep track of original scaleheight & scalemode values
    Z = ObjName.ScaleMode
    ObjName.DrawStyle = 6                       ' Now set these drawing variables
    ObjName.DrawMode = 13
    ObjName.DrawWidth = 13
    ObjName.ScaleMode = 3
    ObjName.ScaleHeight = 256
    ' loop thru each color value & color the object
    For I = 0 To 255
        If I > Red Then R = Red Else R = I              ' if red value exceeded, use red value
        If I > Green Then G = Green Else G = I       ' if green value exceeded, use green value
        If I > Blue Then B = Blue Else B = I             ' if blue value exceeded, use blue value
        ObjName.Line (0, y)-(ObjName.Width, y + 1), RGB(R, G, B), BF    ' paint line Y
        y = y + 1                                                   ' increment line counter
    Next I
    ObjName.ScaleHeight = x                             ' reset original scaleheight & scalemode properties
    ObjName.ScaleMode = Z                               ' this is important if other objects use its size for references

DoGradient = True           ' return value of true
End Function

Public Function ExtractData(sInChar As String, sSep As String, iListNr As Integer) As String
'=======================================================================
'   This little function extracts a specific occurrence of a character delimited string
'   Required passed variables:
'   - sInChar: this is the full character delimited string
'   - sSep: this is the character being used as a delimiter
'   - iListNr: this is the occurrence you want returned (1st, 2nd, etc)
'=======================================================================
Dim intI As Integer, intJ As Integer, intCount As Integer
' if no string is sent, exit function
' Inserted by LaVolpe
On Error GoTo Function_ExtractData_General_ErrTrap_by_LaVolpe
If Len(sInChar) = 0 Then Exit Function
If Right(sInChar, 1) <> sSep Then sInChar = sInChar & sSep  ' add the delimited character at end of string if needed

intJ = InStr(intI + 1, sInChar, sSep)   ' find the first occurrence of the delimiter
If intJ > 0 Then intCount = 1              ' start a count

Do Until intCount = iListNr Or intJ = 0     ' continue thru the string until the count=iListNr or no more delimiters
    intI = intJ                                             ' reset the starting point to the last delimiter found
    intJ = InStr(intJ + 1, sInChar, sSep)     ' try to find another delimiter
     If intJ > 0 Then intCount = intCount + 1   ' increment the count if one is found
Loop
' if the correct number of items were found, then intJ will always be > 0
If intJ > 0 Then ExtractData = Mid(sInChar, intI + 1, intJ - intI - 1)  ' return the text between the delimiters
Exit Function

Function_ExtractData_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function ExtractData]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Public Sub RefreshCategories()
'=======================================================================
'   This sub repopulates the main window with all categories from the database and selects the one matching
'       the current filter setting and can be called by the form for updating categories
'=======================================================================
Dim rsCats As DAO.Recordset, sLastFilter As String, intI As Integer
' Inserted by LaVolpe
On Error GoTo Sub_RefreshCategories_General_ErrTrap_by_LaVolpe
sLastFilter = ExtractData(mainFilterIndex, "|", 1)      ' extract current filter setting
frmLibrary.lstFilter.Clear                                   ' clear combo box & set recordset for referencing
On Error GoTo AddFinalChoice
Set rsCats = mainDB.OpenRecordset("tblCategories", dbOpenDynaset)
If rsCats.RecordCount > 0 Then                              ' loop thru each category & add it to the combo box
    With rsCats
        .MoveFirst
        Do While .EOF = False
            frmLibrary.lstFilter.AddItem .Fields("Category")     ' category name & track it's db record ID
            frmLibrary.lstFilter.ItemData(frmLibrary.lstFilter.NewIndex) = .Fields("ID")
            .MoveNext
        Loop
    End With
End If
rsCats.Close                            ' close the recordset
Set rsCats = Nothing

AddFinalChoice:
With frmLibrary.lstFilter    ' now add an entry of All Categories & look for the current filter setting
    .AddItem "[ALL Categories]", 0
    .ItemData(0) = -1
    For intI = 0 To .ListCount
         If frmLibrary.lstFilter.List(intI) = sLastFilter Then Exit For
    Next
    .ListIndex = intI                   ' if found select it, otherwise select the first entry (intI=0 unless match found)
End With
Exit Sub

Sub_RefreshCategories_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub RefreshCategories]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Public Sub RefreshOtherApps()
Dim rsApps As Recordset, strSQL As String, intI As Integer
On Error GoTo NoAppsLoaded
For intI = frmLibrary.mnuOtherApps.Count - 1 To 1 Step -1
    Unload frmLibrary.mnuOtherApps(intI)
Next
strSQL = "Select * From tblApplications Order by tblApplications.AppName;"
Set rsApps = mainDB.OpenRecordset(strSQL, dbOpenDynaset)
With frmLibrary.mnuOtherApps
    If rsApps.RecordCount > 0 Then
        rsApps.MoveFirst
            Load .Item(1): intI = 2
            .Item(1).Caption = "-"
        Do While rsApps.EOF = False
            Load .Item(intI)
            .Item(intI).Caption = rsApps.Fields("AppName")
            .Item(intI).Tag = rsApps.Fields("AppExe")
            intI = intI + 1
            rsApps.MoveNext
        Loop
    End If
End With
NoAppsLoaded:
Exit Sub
End Sub

Public Sub RefreshLanguages()
'=======================================================================
'   This sub repopulates the main window with all categories from the database and selects the one matching
'       the current filter setting and can be called by the form updating languages
'=======================================================================
Dim rsLang As DAO.Recordset, strSQL As String, sLastFilter As String, intI As Integer, bDefaultExists As Boolean
' Inserted by LaVolpe
On Error GoTo Sub_RefreshLanguages_General_ErrTrap_by_LaVolpe
MyDefaults.Language = CLng(GetSetting("LaVolpeCodeSafe", "Settings", "Language", "0"))

sLastFilter = ExtractData(mainFilterIndex, "|", 5)      ' extract current filter setting
frmLibrary.cboFilter(2).Clear                                   ' clear combo box & set recordset for referencing
On Error GoTo AddFinalChoice
Set rsLang = mainDB.OpenRecordset("tblLanguage", dbOpenDynaset)
If rsLang.RecordCount > 0 Then                              ' loop thru each language & add it to the combo box
    With rsLang
        .MoveFirst
        Do While .EOF = False
            frmLibrary.cboFilter(2).AddItem .Fields("Language") ' language name & track it's db record ID
            frmLibrary.cboFilter(2).ItemData(frmLibrary.cboFilter(2).NewIndex) = .Fields("ID")
            ' see if the language just added happens to be the user's default language & if so, track it
            If .Fields("ID") = MyDefaults.Language Then bDefaultExists = True
            .MoveNext
        Loop
    End With
End If
rsLang.Close                            ' close the recordset
Set rsLang = Nothing

AddFinalChoice:
With frmLibrary.cboFilter(2)    ' now add an entry of All Languages & look for current filter setting
    .AddItem "[ALL Languages]", 0
    .ItemData(0) = -1
    For intI = 0 To .ListCount
         If frmLibrary.cboFilter(2).List(intI) = sLastFilter Then Exit For
    Next
    .ListIndex = intI               ' if found select it, otherwise select the first entry (intI=0 unless match found)
    ' if the default did not exist in listing, then set global variable to the last item in the list
    If bDefaultExists = False Then MyDefaults.Language = .ItemData(.ListCount - 1)
End With
Exit Sub

Sub_RefreshLanguages_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub RefreshLanguages]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Public Sub FilterRecordset()
'=======================================================================
'   Function filters the mainRS using user-defined filter
'=======================================================================
Dim sFilters(0 To 2) As String, bDefaults(0 To 2) As Boolean
Dim sSelectClause As String, sWhereClause As String, intI As Integer, sOrderByClause As String
'   [All Categories]    - default category
'   [With & Without]  - default attachments
'   [All Languages]     - default language
' Inserted by LaVolpe
On Error GoTo Sub_FilterRecordset_General_ErrTrap_by_LaVolpe
For intI = 0 To 2       ' check each filter option
    sFilters(intI) = ExtractData(mainFilterIndex, "|", intI * 2 + 1)    ' extract the current filter for the one being checked
    ' if the current filter isn't the default filter then we need to build a query string
    If sFilters(intI) <> Choose(intI + 1, "[All Categories]", "[With & Without]", "[All Languages]") Then
        bDefaults(intI) = False     ' flag indicating filter being used
        Select Case intI
            Case 0:         ' Filtering categories and the query string needed
                sSelectClause = "INNER JOIN tblCodeCatXref ON tblSourceCode.IDnr = tblCodeCatXref.CodeID"
            Case 1:         ' Filtering attachments and the query string needed
                If Len(sSelectClause) Then sSelectClause = "(" & sSelectClause & ") "   ' format with required parentheses
                ' depending on whether or not we want records with or without attachments, the type of join needs
                '   to be modified to either INNER or LEFT
                If sFilters(1) = "Those Without" Then sSelectClause = sSelectClause & "LEFT " Else sSelectClause = sSelectClause & "INNER "
                sSelectClause = sSelectClause & "JOIN tblAttachments ON tblSourceCode.IDnr = tblAttachments.RecIDRef"
            Case 2:         ' Filtering languages and the query string needed
                If Len(sSelectClause) Then sSelectClause = "(" & sSelectClause & ") "   ' format with required parentheses
                sSelectClause = sSelectClause & "INNER JOIN tblCodeLangXref ON tblSourceCode.IDnr = tblCodeLangXref.CodeID"
        End Select
    Else
        bDefaults(intI) = True  ' no filter used
    End If
Next
If Len(sSelectClause) Then  ' if a filter string was constructed, we need to construct the appropriate SELECT clause
    intI = 0
    Do Until Mid(sSelectClause, intI + 1, 1) <> "("     ' find the first position after a left parenthesis
        intI = intI + 1
    Loop
    If intI = 0 Then    ' if not found, then this is the proper format
        sSelectClause = "SELECT * FROM tblSourceCode " & sSelectClause
    Else                    ' if not found, then this isthe proper format
        sSelectClause = "SELECT * FROM " & Left(sSelectClause, intI) & "tblSourceCode " & Mid(sSelectClause, intI + 1)
    End If
Else                        ' if no filter string constructed (using all defaults) then this is the proper select clause
    sSelectClause = "SELECT * FROM tblSourceCode"
End If
' Last of 3 steps -- build the Where clause
For intI = 0 To 2
    If bDefaults(intI) = False Then ' if a filter was used then a Where clause is needed
        Select Case intI
            Case 0:     ' for Categories
                sWhereClause = "((tblCodeCatXref.CatID)=" & ExtractData(mainFilterIndex, "|", intI * 2 + 2) & ")"
            Case 1:     ' for filter where looking for Without Attachments
                If InStr(sSelectClause, "LEFT JOIN") Then
                    If Len(sWhereClause) Then sWhereClause = sWhereClause & "AND "
                    sWhereClause = sWhereClause & "((tblAttachments.ID) Is Null) "
                End If
            Case 2:     ' for Languages
                If Len(sWhereClause) Then sWhereClause = sWhereClause & "AND "
                sWhereClause = sWhereClause & "((tblCodeLangXref.LangID)=" & ExtractData(mainFilterIndex, "|", intI * 2 + 2) & ")"
        End Select
    End If
Next
' Wow! now finish the where clause
If Len(sWhereClause) Then sWhereClause = " Where (" & sWhereClause & ")"
' Now the Order By clause
Select Case frmLibrary.cboFilter(0).ListIndex
Case 1, 2
    sOrderByClause = "tblSourceCode.OrigDate"
    If frmLibrary.cboFilter(0).ListIndex = 2 Then sOrderByClause = sOrderByClause & " !@!"
Case 3, 4
    sOrderByClause = "tblSourceCode.UpdateDate"
    If frmLibrary.cboFilter(0).ListIndex = 4 Then sOrderByClause = sOrderByClause & " !@!"
Case Else
    sOrderByClause = "tblSourceCode.CodeName"
End Select
If InStr(sOrderByClause, "tblSourceCode.CodeName") = 0 Then sOrderByClause = sOrderByClause & ", tblSourceCode.CodeName;"
sOrderByClause = " ORDER BY " & Replace(sOrderByClause, "!@!", "Desc")
' reset the current recordset, if active & open a new recordset using the filter string
Set mainRS = Nothing
Set mainRS = mainDB.OpenRecordset(sSelectClause & sWhereClause & sOrderByClause, , dbReadOnly)
RefreshCodeList     ' call function to repopulate the main window code listing
Exit Sub

Sub_FilterRecordset_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub FilterRecordset]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Function ExtractAttachment(RSfield As DAO.Field, strFileName As String, ActualSize As Long) As Boolean
'=======================================================================
'   This function will pull the contents of an open recordset and save it to a filename
'   Required variables:
'   -RSfield: the field from an open recordset, not the field name but the field
'   -strFileName: the full path and name of the file to store the data
'=======================================================================

 Dim FileNum As Integer
 Dim Buffer() As Byte
 Dim bytesNeeded As Long
 Dim Buffers As Long
 Dim Remainder As Long
 Dim Offset As Long
 Dim R As Integer
 Dim I As Long
 Dim ChunkSize As Long
On Error GoTo ExtractError
 
 ChunkSize = 65536          ' size of data to read/write at a time (lower if needed)
 bytesNeeded = ActualSize  ' actual size of the data field
 If Len(Dir(strFileName)) Then Kill strFileName ' delete the filename if it exists
' Calculate the number of buffers needed to copy
 Buffers = bytesNeeded \ ChunkSize
 Remainder = bytesNeeded Mod ChunkSize
 ' Copy the file to the temporary file chunk by chunk:
 FileNum = FreeFile
 Open strFileName For Binary As #FileNum
 For I = 0 To Buffers - 1
    ReDim Buffer(ChunkSize)
    Buffer = RSfield.GetChunk(Offset, ChunkSize)
    Put #FileNum, , Buffer()
    Offset = Offset + ChunkSize
 Next        ' Copy the remaining chunk of the bitmap to the file:
 ReDim Buffer(Remainder)
 Buffer = RSfield.GetChunk(Offset, Remainder)
 Put #FileNum, , Buffer()
ExtractAttachment = True

ReleaseFile:
On Error Resume Next
Close #FileNum
Exit Function

ExtractError:
MsgBox "Following error preventing requested action." & vbCrLf & Err.Description, vbExclamation + vbOKOnly
Resume ReleaseFile
End Function

Public Function LoadAttach(RSfield As DAO.Field, FileName As String) As Long
'=======================================================================
'   This will load a file into an OLE data field within an open recordset
'   Required Variables:
'   -RSfield: actual field to store data to, not the field name but the field itself
'   -FileName: full path & name of the file to load
'=======================================================================

 Dim ChunkSize As Long
 Dim FileNum As Integer
 Dim Buffer()  As Byte
 Dim bytesNeeded As Long
 Dim Buffers As Long
 Dim Remainder As Long
 Dim I As Long

' Inserted by LaVolpe
On Error GoTo Function_LoadAttach_General_ErrTrap_by_LaVolpe
LoadAttach = -1
If Dir(FileName) = "" Then          ' ensure the file can be found
    MsgBox "Failed to load attachment. File not found", vbExclamation + vbOKOnly
    Exit Function
End If
On Error GoTo LoadError
ChunkSize = 65536                    ' size of file to read/write at a time (lower if needed)
FileNum = FreeFile
Open FileName For Binary As #FileNum
bytesNeeded = LOF(FileNum)
Buffers = bytesNeeded \ ChunkSize
Remainder = bytesNeeded Mod ChunkSize
For I = 0 To Buffers - 1
    ReDim Buffer(ChunkSize)
    Get #FileNum, , Buffer
    RSfield.AppendChunk Buffer
Next
ReDim Buffer(Remainder)
Get #FileNum, , Buffer
RSfield.AppendChunk Buffer
LoadAttach = bytesNeeded

ReleaseFile:
On Error Resume Next
Close #FileNum
Exit Function

LoadError:
MsgBox "Following error prevent requested action." & vbCrLf & Err.Description, vbExclamation + vbOKOnly
Resume ReleaseFile
Exit Function

Function_LoadAttach_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function LoadAttach]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Public Function GetUniqueFileName(sExt As String, sDestPath As String) As String
'=======================================================================
'   Returns a unique filename for a specified extension in a specified path
'   Required Variables:
'   -sExt: extension of a filename, with or without the dot
'   -sDestPath is the full path name to generate a unique filename from
'=======================================================================
Dim I As Integer, LastNr As Long, bOk As Boolean
' Inserted by LaVolpe
On Error GoTo Function_GetUniqueFileName_General_ErrTrap_by_LaVolpe
If Right(sDestPath, 1) <> "\" Then sDestPath = sDestPath & "\"      ' include a trailing backslash if needed
If Left(sExt, 1) = "." Then sExt = Mid(sExt, 2)                                 ' remove the dot if needed

With frmAttachments.fileAll     ' using a file listbox
    .Path = sDestPath                       ' set the path
    Do While bOk = False                ' loop until a unique name is found (usually only takes one pass)
      LastNr = CLng(Timer)              ' use the timer function as a filename
        GetUniqueFileName = "~Atch" & LastNr & "." & sExt   ' build the filename
        .Pattern = GetUniqueFileName     '  reset the pattern to only this file
        If .ListCount = 0 Then bOk = True ' no match, it's unique
    Loop
    GetUniqueFileName = sDestPath & GetUniqueFileName   ' build the complete filename & exit
End With
Exit Function

Function_GetUniqueFileName_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function GetUniqueFileName]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Public Function CloseAllWindows(bMainWindowToo As Boolean, Optional ExceptionTag As String = "EveryOne") As Integer
'============================================================
' Closes all windows in preparation for shutdown or closes all child windows
'============================================================
' Inserted by LaVolpe
On Error GoTo Function_CloseAllWindows_General_ErrTrap_by_LaVolpe
bAppClose = False               ' set global variable to false

Dim I As Integer, iLastWindow As Integer

' if the main window is not to be closed set variable to prevent it from happening
If bMainWindowToo = False Then iLastWindow = 1

For I = Forms.Count - 1 To iLastWindow Step -1  ' Loop thru each open window & close it (the main window will always be #0)
    If Not IsMissing(ExceptionTag) Then
        If Forms(I).Tag = ExceptionTag Then GoTo CheckNextForm
    End If
    Unload Forms(I)                         ' Unload it
    If bAppClose Then                       ' But if the form reports changed data, then abort here
        CloseAllWindows = 1
        Exit Function
    End If
CheckNextForm:
Next
Exit Function

Function_CloseAllWindows_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function CloseAllWindows]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Private Function IsKeyWord(sText As String, ColorCodes() As Long) As Long
'===================================================================
' Simple function really. This sub holds the keyword bank for VB. Words can easily be added
'   or deleted here.  The words are organized by Letter of the alphabet & up to 3 levels (0-3)
' the 4th level (properties/method keywords) is handled in the calling function
'===================================================================

' The levels related directly to the color codes on the form.
' The only exception are remarks which are hard coded in the declarations section of this module
Dim sKeyWords As String, sText2Check As String, iStart As Integer, levelID As Integer
' Inserted by LaVolpe
On Error GoTo Function_IsKeyWord_General_ErrTrap_by_LaVolpe
sText2Check = Trim(sText)
sText2Check = Replace(sText2Check, "$", "") ' remove $ from functions, but add it back later
BeginKeyWordSearch:                                     ' start search but ignore if keyword color is black
If ColorCodes(levelID) <> 0 Or levelID = 0 Then
    IsKeyWord = ColorCodes(levelID)
    sKeyWords = ""
    Select Case UCase(Left(sText2Check, 1))     ' filter by 1st letter of the word
    Case "'"    ' Remarks
        If levelID = 0 Then
            sKeyWords = Trim(sText)               ' the entire line is a keyword
            IsKeyWord = RmksColor       ' set the color
        End If
    Case "#"                                        ' used in declarations section only
        Select Case levelID
        Case 0: sKeyWords = "#Const|#Else|#End|#If"
        End Select
    Case Chr(vbKeyA)
        Select Case levelID
        Case 0: sKeyWords = "As|Alias|And|Any|Array"
        Case 1: sKeyWords = "Abs|Asc|Atn"
        Case 2: sKeyWords = "App|AppActivate"
        End Select
    Case Chr(vbKeyB)
        Select Case levelID
        Case 0: sKeyWords = "Base|Boolean|Byte|ByVal"
        Case 2: sKeyWords = "Beep"
        End Select
    Case Chr(vbKeyC)
        Select Case levelID
        Case 0: sKeyWords = "Call|Case|Choose|Clear|Close|Compare|Const|Currency"
        Case 1: sKeyWords = "CBool|CByte|CCur|CDate|CDbl|CDec|Chr|CInt|CLng|Cos|CSng|CStr|CVar|CVErr"
        Case 2: sKeyWords = "Command|CreateObject"
        End Select
    Case Chr(vbKeyD)
        Select Case levelID
        Case 0:
            sKeyWords = "Date|Declare|Deftype|Dim|Do|Double|" & _
                "DefBool|DefByte|DefInt|DefLng|DefCur|DefSng|DefDbl|DefDec|DefDate|DefStr|DefObj|DefVar|"
        Case 1: sKeyWords = "DateAdd|DateDiff|DatePart|DateSerial|DateValue|Day"
        Case 2: sKeyWords = "DDB|Dir|DoEvents"
        End Select
    Case Chr(vbKeyE)
        Select Case levelID
        Case 0: sKeyWords = "Each|Else|End|Enum|Err|Error|Exit|Explicit"
        Case 1: sKeyWords = "EOF|Erase|Eqv|Exp"
        Case 2: sKeyWords = "Environ"
        End Select
    Case Chr(vbKeyF)
        Select Case levelID
        Case 0: sKeyWords = "For|Function|False"
        Case 1: sKeyWords = "FileAttr|FileDateTime|FileLen|Fix|Format|FV"
        Case 2: sKeyWords = "FileCopy|FreeFile"
        End Select
    Case Chr(vbKeyG)
        Select Case levelID
        Case 0: sKeyWords = "Get|Global|GoSub|GoTo"
        Case 1: sKeyWords = "GetAttr"
        Case 2: sKeyWords = "GetObject|GetSetting"
        End Select
    Case Chr(vbKeyH)
        Select Case levelID
        Case 1: sKeyWords = "Hex|Hour"
        End Select
    Case Chr(vbKeyI)
        Select Case levelID
        Case 0: sKeyWords = "If|Integer|Is|Imp"
        Case 1: sKeyWords = "InStr|InStrRev|Int|IRR|IsArray|IsDate|IsEmpty|IsError|IsMissing|IsNull|IsNumeric|IsObject"
        Case 2: sKeyWords = "IPmt|Input|InputBox"
        End Select
    Case Chr(vbKeyK)
        Select Case levelID
        Case 2: sKeyWords = "Kill"
        End Select
    Case Chr(vbKeyL)
        Select Case levelID
        Case 0: sKeyWords = "Let|Lib|Like|Loc|Lock|Log|Long|Loop"
        Case 1: sKeyWords = "LBound|LCase|Left|Len|LOF|LSet|LTrim"
        Case 2: sKeyWords = "Line|Load|LoadPicture|LoadResPicture|LoadResString|LoadResData"
        End Select
    Case Chr(vbKeyM)
        Select Case levelID
        Case 0: sKeyWords = "Mod|Module"
        Case 1: sKeyWords = "Mid|Minute|MIRR|Month"
        Case 2: sKeyWords = "Me|MsgBox"
        End Select
    Case Chr(vbKeyN)
        Select Case levelID
        Case 0: sKeyWords = "New|Next|Not|Now"
        Case 1: sKeyWords = "NPer|NPV"
        Case 2: sKeyWords = "Name"
        End Select
    Case Chr(vbKeyO)
        Select Case levelID
        Case 0: sKeyWords = "Object|On|Option|Optional|Or|"
        Case 1: sKeyWords = "Oct"
        Case 2: sKeyWords = "Open"
        End Select
    Case Chr(vbKeyP)
        Select Case levelID
        Case 0: sKeyWords = "Preserve|Private|Property|Public"
        Case 1: sKeyWords = "Pmt|PV"
        Case 2: sKeyWords = "PPmt|Print|Put"
        End Select
    Case Chr(vbKeyQ)
        Select Case levelID
        Case 2: sKeyWords = "QBColor"
        End Select
    Case Chr(vbKeyR)
        Select Case levelID
        Case 0: sKeyWords = "Randomize|ReDim|Reset|Resume|Return"
        Case 1: sKeyWords = "Raise|Rate|Replace|Right|Rnd|RSet|RTrim"
        Case 2: sKeyWords = "RGB"
        End Select
    Case Chr(vbKeyS)
        Select Case levelID
        Case 0: sKeyWords = "Select|Set|Single|Static|Step|Stop|String|Sub|Switch"
        Case 1: sKeyWords = "Second|SendKeys|SetAttr|Sgn|Sin|SLIN|Space|Spc|Sqr|Str|StrComp|StrConv|SYD"
        Case 2: sKeyWords = "SaveSetting|Seek|Shell"
        End Select
    Case Chr(vbKeyT)
        Select Case levelID
        Case 0: sKeyWords = "Text|Then|Time|Timer|To|True|Type"
        Case 1: sKeyWords = "Tab|Tan|TimeSeral|TimeValue|Trim|TypeName"
        End Select
    Case Chr(vbKeyU)
        Select Case levelID
        Case 0: sKeyWords = "Unlock|Until"
        Case 1: sKeyWords = "UBound|UCase"
        Case 2: sKeyWords = "Unload"
        End Select
    Case Chr(vbKeyV)
        Select Case levelID
        Case 0: sKeyWords = "Variant"
        Case 1: sKeyWords = "Val|VarType"
        End Select
    Case Chr(vbKeyW)
        Select Case levelID
        Case 0: sKeyWords = "Wend|While|With"
        Case 1: sKeyWords = "Weekday"
        Case 2: sKeyWords = "Width|Write"
        End Select
    Case Chr(vbKeyX)
        Select Case levelID
        Case 2: sKeyWords = "Xor"
        End Select
    Case Chr(vbKeyY)
        Select Case levelID
        Case 1: sKeyWords = "Year"
        End Select
    Case Else
        IsKeyWord = 0
        Exit Function
    End Select
    sKeyWords = "|" & sKeyWords & "|"       ' add leading/trailing bars as needed
    iStart = InStr(sKeyWords, "|" & sText2Check & "|")  ' see if text word is a keyword
End If
If iStart > 0 Then      ' if so, we extract the keyword from above & replace the text word
    ' this procedure corrects capitalization (i.e.,    end if   becomes End If  )
    sText = Replace(sText, sText2Check, Mid(sKeyWords, iStart + 1, InStr(iStart + 1, sKeyWords, "|") - iStart - 1))
Else    ' oops, word not a keyword, but is it in the other levels?
    If levelID < UBound(ColorCodes) Then    ' check thru each level until done
        levelID = levelID + 1
        GoTo BeginKeyWordSearch
    End If                                                      ' if got this far, word is not a keyword
    IsKeyWord = 0                                       ' so set the text color to black
End If
Exit Function

Exit Function

Function_IsKeyWord_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function IsKeyWord]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Public Function CheckLine4KeyWords(txtObj As Control, tempRTF As Object, Optional TextRange As Variant = Null, _
    Optional bNoSwap As Boolean = False, Optional ProgBar As ProgressBar) As Boolean
'===================================================================
' Primary function to compare text against known key words & color the words appropriately
'===================================================================

Dim iWord As Integer, iBegin As Long, iEnd As Integer, sText As String, iPos As Integer, bNextWordProperty As Boolean
Dim lWordColor As Long, sLineOfText As String, iOffset(0 To 2) As Long, CRoffset As Integer
Dim Looper As Long, LoopStart As Long, LoopStop As Long, sMask As String
Dim NonKeyWords As String, ColorCodes(0 To 3) As Long

On Error Resume Next
ColorCodes(0) = MyDefaults.KeyWd1
ColorCodes(1) = MyDefaults.KeyWd2
ColorCodes(2) = MyDefaults.KeyWd3
ColorCodes(3) = MyDefaults.KeyWd4
With tempRTF
    tempRTF.Font.Name = MyDefaults.Font
    tempRTF.Font.Size = MyDefaults.FontSize
    If IsNull(TextRange) Then           ' coloring an entire file, not just a changed lne of text
        If bNoSwap = False Then .TextRTF = txtObj
    Else                                            ' coloring a changed line of text
        .Text = Mid(txtObj.Text, TextRange(0), TextRange(1))
    End If
    LoopStart = 0: iOffset(0) = 1     ' set startup variables & get nr of lines being checked
    LoopStop = apiSendMessage(.hWnd, EM_GETLINECOUNT, 0, 0) - 1
    If Not ProgBar Is Nothing Then ProgBar.Max = LoopStop
    For Looper = LoopStart To LoopStop  ' loop thru each line in the text box
        If Not ProgBar Is Nothing Then
            If Not IsNull(GP) Then GoTo ShowResults
            ProgBar = Looper
            DoEvents
        End If
        Do Until InStr(Mid(.Text, iOffset(0), 1), vbCr) = 0 And InStr(Mid(.Text, iOffset(0), 1), vbLf) = 0
            iOffset(0) = iOffset(0) + 1
        Loop
        iOffset(2) = apiSendMessage(.hWnd, EM_LINELENGTH, iOffset(0), 0)
        sLineOfText = Mid(.Text, iOffset(0), iOffset(2))    ' calc visible line of text
        iBegin = iOffset(0)     ' set starting character position of the line
        iEnd = iOffset(2)        ' and ending character position of the line
        lWordColor = 0
        GoSub RepaintLine    ' color entire line black
        iPos = Len(sLineOfText) - Len(LTrim(sLineOfText))   ' note any spaces at end of line
        iBegin = InStr(sLineOfText, vbLf)                   ' if a line feed/return near beginning of line,
        If iBegin Then iBegin = 1                           ' adjust offset
        If InStr(sLineOfText, vbCr) > 0 And InStr(sLineOfText, vbCr) < 2 Then iBegin = iBegin + 1
        iPos = iPos + iBegin                    ' now calc the new starting position of the line
        sMask = Replace(sLineOfText, ")", " ")  ' we are going to build a mask that will be a space
        sMask = Replace(sMask, "(", " ")        ' delimited word list. Since keywords can end with a space,
        sMask = Replace(sMask, ",", " ")        ' comma, period, parenthesis (left or right), let's mask these
        sMask = Replace(sMask, ".", " ")
        sMask = Replace(sMask, ":", " ")
        sMask = Replace(sMask, ";", " ")
        ' this check is made at start of line & after each word. Basically, if the word begins with a period
        '   it is assumed to be a property or method & is colored approrpriately. The trick is to check for
        '   a period when the periods have been masked out
        bNextWordProperty = (Mid(sLineOfText, iPos + 1, 1) = ".")
        Do Until iPos >= Len(sMask)     ' check each word in the mask
               iBegin = iPos + 1                 ' increment the starting position for new words
               ' another exception, gotta check for word that begins with a quote, cause the characters
               ' within the quotes are not colored & where there's a starting quote, there should be an ending one too
                If Left(Trim(Mid(sMask, iBegin)), 1) = Chr(34) Then
                    iEnd = InStr(InStr(iBegin, sMask, Chr(34)) + 1, sMask, Chr(34))
                    If iEnd = 0 Then iEnd = Len(sMask) + 1  ' if no ending quote found, go to end of line
                Else                                        ' not quoted, so find the end of the word
                    iEnd = InStr(iBegin + 1, sMask, " ")
                End If
                If iEnd = 0 Then iEnd = Len(sMask) + 1  ' if end of word not found, go to end of line
                iPos = iEnd                             ' keep end of word & parse the word
                sText = Mid(sMask, iBegin, iEnd - iBegin)
                If Len(Trim(sText)) > 0 Then    ' can't be a zero-length string
                    ' it's not, but if it is a quoted word or it started with a period, handle that now
                    If Left(Trim(sText), 1) = Chr(34) Or (bNextWordProperty = True And IsNumeric(sText) = False) Then
                        ' quoted words are colored black & words starting with a period are colored as user defined
                        If bNextWordProperty = True Then lWordColor = ColorCodes(3) Else lWordColor = 0
                    Else    ' a real word -- now check it against library of keywords
                        lWordColor = IsKeyWord(sText, ColorCodes)
                    End If
                    ' build offset values in order to drop the colored word into the string
                    iOffset(1) = Len(sText) - Len(LTrim(sText))  ' left offset
                    If lWordColor = RmksColor Then               ' a remark line
                        iOffset(2) = 0                           ' right offset
                        ' set the length of the remark line & the starting point
                        iEnd = Len(sLineOfText) - iBegin + 1 - iOffset(1) + 2
                        iBegin = iBegin + iOffset(1) + iOffset(0) - 2
                        GoSub RepaintLine   ' color the line
                        Exit Do
                    Else
                        If lWordColor <> 0 Then         ' a keyword since the color is not black
                             iOffset(2) = Len(sText) - Len(RTrim(sText))    ' right offset
                            iEnd = iEnd - iBegin - iOffset(2) - iOffset(1)  ' length of word
                            iBegin = iBegin + iOffset(1) + iOffset(0) - 2   ' starting point of word
                            sText = Trim(sText)                             ' trim the word
                            GoSub PaintKeyWord                              ' and color it
                        End If
                    End If
                End If
                bNextWordProperty = (Mid(sLineOfText, iPos, 1) = ".")   ' see if next word is a property
                If bNextWordProperty = False Then bNextWordProperty = (Mid(sLineOfText, iPos + 1, 1) = ".")
        Loop
        iOffset(0) = iOffset(0) + Len(sLineOfText) + 1
    Next Looper
ShowResults:
    If IsNull(TextRange) Then           ' when is null, coloring a file vs a line of text
        tempRTF.SelStart = 0            ' lines of text are done directly, files are done with a hidden
        tempRTF.SelLength = Len(tempRTF.Text)   ' text box to prevent flickering & actually
        tempRTF.SelFontName = MyDefaults.Font   ' speeds up the process
        tempRTF.SelFontSize = MyDefaults.FontSize
        txtObj = ""
        txtObj.Font.Name = MyDefaults.Font
        txtObj.Font.Size = MyDefaults.FontSize
        txtObj = tempRTF
    End If
.TextRTF = ""
End With
ResetProgressBar:
If Not ProgBar Is Nothing Then ProgBar = 0
Exit Function
' function tested with 19,000+ characters & averages coloring 2,700 characters per second

PaintKeyWord:
If IsNull(TextRange) Then
    tempRTF.SelStart = iBegin
    tempRTF.SelLength = iEnd
    tempRTF.SelColor = lWordColor
    tempRTF.SelText = sText
    tempRTF.SelLength = 0
Else
    txtObj.SelStart = iBegin + TextRange(0) - 1
    txtObj.SelLength = iEnd
    txtObj.SelColor = lWordColor
    txtObj.SelText = sText
    txtObj.SelLength = 0
End If
Return
RepaintLine:
If IsNull(TextRange) Then
    tempRTF.SelStart = iBegin - 1
    tempRTF.SelLength = iEnd
    tempRTF.SelColor = lWordColor
    tempRTF.SelLength = 0
Else
    txtObj.SelStart = iBegin + TextRange(0) - 1
    txtObj.SelLength = iEnd + 1
    txtObj.SelColor = lWordColor
    txtObj.SelLength = 0
End If
Return
End Function
