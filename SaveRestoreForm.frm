VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveRestoreForm 
   Caption         =   "Backup_GUI"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   OleObjectBlob   =   "SaveRestoreForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveRestoreForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSave_Click()
    Dim filePath As String
    Dim cellList As String
    Dim cellArray() As String
    Dim i As Integer
    Dim cellAddress As String
    Dim cellValue As Variant
    Dim fileContent As String
    Dim cell As Range
    
    ' Open Save File Dialog
    filePath = Application.GetSaveAsFilename("mydata.csv", "CSV Files (*.csv), *.csv")
    If filePath = "False" Then Exit Sub
    
    cellList = Me.TextBox1.Value
    cellArray = Split(cellList, ",")
    
    For i = LBound(cellArray) To UBound(cellArray)
        cellAddress = Trim(cellArray(i))
        On Error Resume Next
        Set cell = Range(cellAddress)
        On Error GoTo 0
        If Not cell Is Nothing Then
            cellValue = cell.Value
            fileContent = fileContent & cellAddress & "," & cellValue & vbCrLf
        End If
    Next i
    
    ' Write to CSV file
    Open filePath For Output As #1
    Print #1, fileContent
    Close #1
    
    MsgBox "Data saved successfully!", vbInformation
End Sub

Private Sub btnRestore_Click()
    Dim filePath As String
    Dim fileContent As String
    Dim cellArray() As String
    Dim i As Integer
    Dim cellData() As String
    Dim cellAddress As String
    Dim cellValue As Variant
    
    ' Open Open File Dialog
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
    If filePath = "False" Then Exit Sub
    
    ' Read the content of CSV file
    Open filePath For Input As #1
    fileContent = Input$(LOF(1), 1)
    Close #1
    
    cellArray = Split(fileContent, vbCrLf)
    
    For i = LBound(cellArray) To UBound(cellArray)
        If Len(Trim(cellArray(i))) > 0 Then
            cellData = Split(cellArray(i), ",")
            cellAddress = Trim(cellData(0))
            cellValue = cellData(1)
            On Error Resume Next
            Range(cellAddress).Value = cellValue
            On Error GoTo 0
        End If
    Next i
    
    MsgBox "Data restored successfully!", vbInformation
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Save/Restore Cell Values"
    Me.Label1.Caption = "Enter cell addresses (comma-separated):"
    Me.TextBox1.Value = "C2, D5, D6, D7, D8, D10, D11, G5, G6, G7, G9, D10, D11, D13, D14, D15, D16, D17, D20, D21, D22, D23, D24, L3, L4, L5, L6, L7, L8, L20"
    Me.btnSave.Caption = "Save to CSV"
    Me.btnRestore.Caption = "Restore from CSV"
End Sub

