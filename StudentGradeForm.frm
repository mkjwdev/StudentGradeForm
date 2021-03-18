VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StudentGradeForm 
   Caption         =   "Register Grades"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   OleObjectBlob   =   "StudentGradeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StudentGradeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRegister_Click()
    If txtStudentName = "" Then
        MsgBox "Type the student’s name!"
        Exit Sub
    Else
        If txtExam1 = "" Or txtExam2 = "" Then
            MsgBox "Type the student’s grades!"
            Exit Sub
        End If
        Sheets(ActiveSheet.Name).Range("A1") = "Student Name"
        Sheets(ActiveSheet.Name).Range("B1") = "Final Grade"
        Sheets(ActiveSheet.Name).Range("C1") = "Result"
        Sheets(ActiveSheet.Name).Range("A2").Select
        Do Until IsEmpty(ActiveCell)
            ActiveCell.Offset(1, 0).Select
        Loop
        ActiveCell.Value = txtStudentName
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = Format(CDbl(lblAverage), "#,#0.0")
        ActiveCell.Offset(0, 1).Select
    End If
    If chkIsApproved = True Then
        ActiveCell.Value = "Approved"
    Else
        ActiveCell.Value = "Disapproved"
    End If
    txtStudent = ""
    txtExam1 = ""
    txtExam2 = ""
    lblAverage = ""
    chkIsApproved = False
    lblAverage.ForeColor = vbRed
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub txtExam1_Change()
    If txtExam1 = "" Or txtExam2 = "" Then
        lblAverage = "0.0"
        chkIsApproved = False
        lblAverage.ForeColor = vbRed
    Else
        lblAverage = Format((CDbl(txtExam1) + CDbl(txtExam2)) / 2, "#,#0.0")
    End If
    If CDbl(lblAverage) >= 5 Then
        chkIsApproved = True
        lblAverage.ForeColor = vbBlue
    Else
        chkIsApproved = False
        lblAverage.ForeColor = vbRed
    End If
End Sub

Private Sub txtExam2_Change()
    If txtExam1 = "" Or txtExam2 = "" Then
        lblAverage = "0.0"
        chkIsApproved = False
        lblAverage.ForeColor = vbRed
    Else
        lblAverage = Format((CDbl(txtExam1) + CDbl(txtExam2)) / 2, "#,#0.0")
    End If
    If CDbl(lblAverage) >= 5 Then
        chkIsApproved = True
        lblAverage.ForeColor = vbBlue
    Else
        chkIsApproved = False
        lblAverage.ForeColor = vbRed
    End If
End Sub

Private Sub UserForm_Click()

End Sub
