Attribute VB_Name = "FormControl"
Private List() As Control
Private curr_obj As Object
Private iHeight As Integer
Private iWidth As Integer
Private x_size As Double
Private y_size As Double
'*****************************************************************************************
'                           LICENSE INFORMATION
'*****************************************************************************************
'   FormControl Version 2.0
'   Code module for resizing a form based on screen size, then resizing the
'   controls based on the forms size
'
'   Copyright (C) 2015
'   Vishnu Sivan
'   Email: codemaker2014@gmail.com
'   Created: AUG2
'
' see <http://www.bpccollege.ac.in>.
'*****************************************************************************************

Private Type Control
    Index As Integer
    Name As String
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
End Type

Public Sub ResizeControls(frm As Form)
Dim i As Integer
'   Get ratio of initial form size to current form size
x_size = frm.height / iHeight
y_size = frm.width / iWidth

'Loop though all the objects on the form
'Based on the upper bound of the # of controls
For i = 0 To UBound(List)
    'Grad each control individually
    For Each curr_obj In frm
        'Check to make sure its the right control
        If curr_obj.TabIndex = List(i).Index Then
            'Then resize the control
             With curr_obj
                .Left = List(i).Left * y_size
                .width = List(i).width * y_size
                .height = List(i).height * x_size
                .Top = List(i).Top * x_size
             End With
        End If
    'Get the next control
    Next curr_obj
Next i
End Sub

Public Function SetFontSize() As Integer
    'Make sure x_size is greater than 0
    If Int(x_size) > 0 Then
    'Set the font size
        SetFontSize = Int(x_size * 8)
    End If
End Function

Public Sub GetLocation(frm As Form)
Dim i As Integer
'   Load the current positions of each object into a user defined type array.
'   This information will be used to rescale them in the Resize function.

'Loop through each control
For Each curr_obj In frm
'Resize the Array by 1, and preserve
'the original objects in the array
    ReDim Preserve List(i)
    With List(i)
        .Name = curr_obj
        .Index = curr_obj.TabIndex
        .Left = curr_obj.Left
        .Top = curr_obj.Top
        .width = curr_obj.width
        .height = curr_obj.height
    End With
    i = i + 1
Next curr_obj
    
'   This is what the object sizes will be compared to on rescaling.
    iHeight = frm.height
    iWidth = frm.width
End Sub

Public Sub CenterForm(frm As Form)
    frm.Move (Screen.width - frm.width) \ 2, (Screen.height - frm.height) \ 2
End Sub

Public Sub ResizeForm(frm As Form)
    'Set the forms height
    frm.height = Screen.height / 2
    'Set the forms width
    frm.width = Screen.width / 2
    'Resize all of the controls
    'based on the forms new size
    ResizeControls frm
End Sub


