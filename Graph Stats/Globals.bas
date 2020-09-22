Attribute VB_Name = "Globals"
Option Explicit

Public Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Public Declare Function SendMessageStr Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long

Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const ES_NUMBER = &H2000&
Public Const GWL_STYLE = (-16)
Public Const LB_FINDSTRING As Long = &H18F
Public Const LB_FINDSTRINGEXACT As Long = &H1A2
Public Const CB_ERR As Long = (-1)
Public Const LB_ERR As Long = (-1)
Public Const WM_USER As Long = &H400
Public Const CB_FINDSTRING As Long = &H14C
Public Const CB_SHOWDROPDOWN As Long = &H14F

Public Counter As Integer
Public i As Integer
Public DB As Database
Public Tbl As TableDef
Public Fld As Field
Public RS As Recordset
Public CurrentRecord As String
Public CurrentCategory As String
Public Find As TableDef
Public Sorted As Boolean
Public SortedArray() As Double
Public TheArray() As Double
Public Form_Loaded As Boolean

Public Function MonthName(ByVal Month As Integer) As String
Select Case Month
Case 1
MonthName = "Jan"
Case 2
MonthName = "Feb"
Case 3
MonthName = "Mar"
Case 4
MonthName = "Apr"
Case 5
MonthName = "May"
Case 6
MonthName = "June"
Case 7
MonthName = "July"
Case 8
MonthName = "Aug"
Case 9
MonthName = "Sept"
Case 10
MonthName = "Oct"
Case 11
MonthName = "Nov"
Case 12
MonthName = "Dec"
End Select
End Function

Public Function CurrentTable() As String
Dim NewTable As String
If CurrentCategory = "" Then
NewTable = Day(Now) & "_" & MonthName(Month(Now)) & "_" & Year(Now)
    If CheckForTable(NewTable) = False Then
    CreateTable NewTable
    End If
CurrentTable = NewTable
CurrentCategory = NewTable
Else
CurrentTable = CurrentCategory
End If
End Function
Public Sub CreateTable(TableName As String)

Set Tbl = DB.CreateTableDef(TableName)
Set Fld = Tbl.CreateField("Description", dbText, 50)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Tag", dbText, 50)
Tbl.Fields.Append Fld

For i = 1 To 26
Set Fld = Tbl.CreateField(CStr(i), dbSingle, 8)
Tbl.Fields.Append Fld
Next

DB.TableDefs.Append Tbl

End Sub
Public Sub DeleteTable(TableName As String)
Dim Action As VbMsgBoxResult
Action = MsgBox("Delete the category " & """" & TableName & """" & " and all the data it contains?", vbOKCancel, "DELETE TABLE")
If Action = vbYes Then
DB.TableDefs.Delete TableName
End If
End Sub
Public Sub BeginRecord(TableName As String, RecName As String)

If TableName = "" Then Exit Sub
If RecName = "" Then Exit Sub

Set RS = DB.OpenRecordset(TableName)
With RS
.AddNew
.Fields("Description") = RecName
.Update
.Close
End With
CurrentCategory = TableName
CurrentRecord = RecName
End Sub


Public Function CheckForTable(TableName As String) As Boolean

If TableName = "" Then Exit Function
Set Find = New TableDef
    For Each Find In DB.TableDefs
        If Find.Name = TableName Then
        CheckForTable = True
        Set Find = Nothing
        Exit Function
        End If
    Next
CheckForTable = False
End Function





Public Sub OpenDB()
On Error GoTo Out:
If IsFile(App.Path & "\DataFile.MDB") Then GoTo SkipCreation
Set DB = DBEngine.Workspaces(0).CreateDatabase(App.Path + "\DataFile.MDB", dbLangGeneral)
SkipCreation:
Set DB = OpenDatabase(App.Path & "\DataFile.MDB")
Exit Sub
Out:
MsgBox "Unable to open or create the database."
End Sub

Public Function IsFile(FileString As String) As Boolean
Dim FileNumber As Integer
On Error Resume Next
FileNumber = FreeFile()
Open FileString For Input As #FileNumber
If Err Then
IsFile = False
Exit Function
End If
IsFile = True
Close #FileNumber
End Function
