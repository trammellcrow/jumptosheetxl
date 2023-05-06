VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Activate Worksheet"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Jump to Sheet UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim selectedWorksheetName As String

    If ListBox1.ListIndex <> -1 Then
        selectedWorksheetName = ListBox1.Value

        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name = selectedWorksheetName Then
                ws.Activate
                Exit For
            End If
        Next ws
    End If

    Unload Me
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ListBox1.AddItem ws.Name
    Next ws
End Sub
