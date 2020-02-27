VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12255
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim slides(5) As Boolean
Dim filename As String


Private Sub CheckBox2_Click()
    If UserForm2.CheckBox2.Value = True Then
       slides(1) = True
    Else
       slides(1) = False
    End If
End Sub

Private Sub CheckBox3_Click()
    If UserForm2.CheckBox3.Value = True Then
       slides(2) = True
    Else
       slides(2) = False
    End If
End Sub

Private Sub CheckBox4_Click()
    If UserForm2.CheckBox4.Value = True Then
       slides(3) = True
    Else
       slides(3) = False
    End If
End Sub

Private Sub CheckBox5_Click()
    If UserForm2.CheckBox5.Value = True Then
       slides(4) = True
    Else
       slides(4) = False
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim p As Integer
    Dim objPresentation As Presentation
    
    On Error GoTo TestMe_Error
    
    If filename = vbNullString Then
        MsgBox "Please give a file"
    Else
        Set objPresentation = Presentations.Open(ActivePresentation.Path & "\" & filename)
        For p = 0 To 5
            If slides(p) = True Then
                objPresentation.slides.Item(p + 1).Copy
                 Presentations.Item(1).slides.Paste
                 Presentations.Item(1).slides.Item(Presentations.Item(1).slides.Count).Design = _
                    objPresentation.slides.Item(p + 1).Design
            End If
        Next p
        objPresentation.Close
        UserForm2.CheckBox1.Value = False
        UserForm2.CheckBox2.Value = False
        UserForm2.CheckBox3.Value = False
        UserForm2.CheckBox4.Value = False
        UserForm2.CheckBox5.Value = False
    End If
    
TestMe_Error:
    Select Case Err.Number
    Case -2147467259
        MsgBox "Please insert a valid filename! It must be stored in the same folder as this file. Ex: test.pptx"
    Case 0
        Debug.Print "Worked"
    Case Else
        MsgBox Err.Number & " " & Err.Source & " - " & Err.Description, vbCritical, "Error"
    End Select
End Sub

Private Sub CommandButton2_Click()
Unload UserForm2
End Sub

Private Sub Frame1_Click()
End Sub

Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()
    filename = UserForm2.TextBox1.Text
End Sub

Private Sub UserForm_Click()
    Call FillArray
End Sub


Private Sub CheckBox1_Click()
    If UserForm2.CheckBox1.Value = True Then
       slides(0) = True
    Else
       slides(0) = False
    End If
End Sub

Sub FillArray()
    Dim inti As Integer
    For inti = 0 To 5
        slides(inti) = False
    Next inti
End Sub
