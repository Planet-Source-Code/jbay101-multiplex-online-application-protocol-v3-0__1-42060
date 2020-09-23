Attribute VB_Name = "modBuildDialog"
Option Explicit
Dim data As String

Function BuildDialog(frm As Form, Target As String, Path As String)
Dim defs As String
Dim xtype As String
Dim usepic As Boolean
'If MsgBox("Use PictureBox???", vbYesNo) = vbYes Then
usepic = True
'Else
'usepic = False
'End If


defs = "Dim " & frm.Name & "ID, " & frm.Name

'start with the main form
add "&fid = CForm.CreateForm()"
add "Set &frm = CForm.frmobj(&fid)"


add ""

Dim isarray As Boolean
Dim XName As String


Dim XControl As Control
For Each XControl In frm.Controls
    
    On Error Resume Next
    If TypeOf XControl Is PictureBox Then If usepic = False Then GoTo skip
    
    Err.Clear
    XControl.Index = XControl.Index
    
    If Err.Number <> 343 Then
        'MsgBox "343"
        defs = defs & ", " & XControl.Name & "_" & XControl.Index
    Else
        Err.Clear
        defs = defs & ", " & XControl.Name
    End If
    If TypeOf XControl Is CommandButton Then xtype = "CommandButton"
    If TypeOf XControl Is Timer Then xtype = "Timer"
    If TypeOf XControl Is TextBox Then xtype = "TextBox"
    If TypeOf XControl Is PictureBox Then xtype = "PictureBox"
    If TypeOf XControl Is Image Then xtype = "Image"
    If TypeOf XControl Is OptionButton Then xtype = "OptionButton"
    If TypeOf XControl Is CheckBox Then xtype = "CheckBox"
    If TypeOf XControl Is Label Then xtype = "Label"
    If TypeOf XControl Is Shape Then xtype = "Shape"
    If TypeOf XControl Is ListBox Then xtype = "ListBox"
    If TypeOf XControl Is ComboBox Then xtype = "ComboBox"
    If TypeOf XControl Is Frame Then xtype = "Image"
    
    'MsgBox xtype
    
    On Error Resume Next
    Err.Clear
    XControl.Index = XControl.Index
    
    If Err.Number <> 343 Then
        'MsgBox "343"
        Err.Clear
        add "CForm.CreateControl &fid, '" & XControl.Name & "_" & XControl.Index & "', '" & xtype & "'"
        add "Set " & XControl.Name & "_" & XControl.Index & " = CForm.obj(&fid, '" & XControl.Name & "_" & XControl.Index & "')"
        add ""
    Else
        add "CForm.CreateControl &fid, '" & XControl.Name & "', '" & xtype & "'"
        add "Set " & XControl.Name & " = CForm.obj(&fid, '" & XControl.Name & "')"
        add ""
    End If
skip:
Next
 add ""
 
add "&frm.Caption = '" & frm.Caption & "'"
add "&frm.Height = " & frm.Height
add "&frm.Width = " & frm.Width
add "&frm.Top = " & frm.Top
add "&frm.Left = " & frm.Left
add "&frm.AutoRedraw = " & frm.AutoRedraw
add "&frm.Enabled = " & frm.Enabled
add "&frm.BackColor = " & frm.BackColor
add "&frm.ForeColor = " & frm.ForeColor
add "&frm.ScaleMode = " & frm.ScaleMode
'add "&frm.ShowInTaskbar = " & frm.ShowInTaskbar
'add "&frm.StartUpPosition = " & frm.StartUpPosition
add "&frm.Visible = " & frm.Visible
        If frm.Picture <> 0 Then
            SavePicture frm.Picture, Path & "\" & frm.Name & ".bmp"
            add "Set " & frm.Name & "." & "Picture" & " = CForm.LoadImage('%path%\" & frm.Name & ".bmp')"
        End If
    

        If frm.Icon <> 0 Then
            SavePicture frm.Icon, Path & "\" & frm.Name & ".ico"
            add "Set " & frm.Name & "." & "Icon" & " = CForm.LoadImage('%path%\" & frm.Name & ".ico')"
        End If
    


For Each XControl In frm.Controls
    If TypeOf XControl Is PictureBox Then If usepic = False Then GoTo skip2
    On Error Resume Next
    'MsgBox XControl.Parent
    
    Err.Clear
    XControl.Index = XControl.Index
    
    If Err.Number <> 343 Then
        isarray = True
        XName = XControl.Name & "_" & XControl.Index
    Else
        isarray = False
        XName = XControl.Name
    End If

    If TypeOf XControl Is Timer Then
    ElseIf TypeOf XControl Is Line Then
    Else
        'MsgBox XName
        add XName & "." & "Left" & " = " & XControl.Left
        add XName & "." & "Top" & " = " & XControl.Top
        add XName & "." & "Width" & " = " & XControl.Width
        add XName & "." & "Height" & " = " & XControl.Height
    End If
    
    
    If TypeOf XControl Is CommandButton Then
        add XName & "." & "Caption" & " = " & FormatText(XControl.Caption)
        add XName & "." & "Font.Bold" & " = " & XControl.Font.Bold
        add XName & "." & "Font.Size" & " = " & XControl.Font.Size
        add XName & "." & "Font.Italic" & " = " & XControl.Font.Italic
        add XName & "." & "Font.Name" & " = '" & XControl.Font.Name & "'"
        add XName & "." & "Visible" & " = " & XControl.Visible
        
        add XName & "." & "BackColor" & " = " & XControl.BackColor
        add XName & "." & "Cancel" & " = " & XControl.Cancel
        add XName & "." & "Enabled" & " = " & XControl.Enabled
        add XName & "." & "Default" & " = " & XControl.Default
        add XName & "." & "TabIndex" & " = " & XControl.TabIndex
    End If
    
    If TypeOf XControl Is Timer Then
        add XName & "." & "Enabled" & " = " & XControl.Enabled
        add XName & "." & "Interval" & " = " & XControl.Interval
    End If
    
    If TypeOf XControl Is TextBox Then
        add XName & "." & "Text" & " = " & FormatText(XControl.text)
        add XName & "." & "Font.Bold" & " = " & XControl.Font.Bold
        add XName & "." & "Font.Size" & " = " & XControl.Font.Size
        add XName & "." & "Font.Italic" & " = " & XControl.Font.Italic
        add XName & "." & "Font.Name" & " = '" & XControl.Font.Name & "'"
        add XName & "." & "Visible" & " = " & XControl.Visible
        
'        add XName & "." & "MultiLine" & " = " & XControl.MultiLine
        
        add XName & "." & "BackColor" & " = " & XControl.BackColor
        add XName & "." & "ForeColor" & " = " & XControl.ForeColor
        add XName & "." & "PasswordChar" & " = '" & XControl.PasswordChar & "'"
        add XName & "." & "Enabled" & " = " & XControl.Enabled
        
        add XName & "." & "TabIndex" & " = " & XControl.TabIndex
        
    End If
    
    If TypeOf XControl Is PictureBox Then
        add XName & "." & "ScaleMode" & " = " & XControl.ScaleMode
    
        add XName & "." & "BackColor" & " = " & XControl.BackColor
        add XName & "." & "BorderStyle" & " = " & XControl.BorderStyle
        add XName & "." & "AutoRedraw" & " = " & XControl.AutoRedraw
        add XName & "." & "Visible" & " = " & XControl.Visible
        
        add XName & "." & "FillColor" & " = " & XControl.FillColor
        add XName & "." & "FillStyle" & " = " & XControl.FillStyle
        
        add XName & "." & "DrawMode" & " = " & XControl.DrawMode
        
        add XName & "." & "Font.Bold" & " = " & XControl.Font.Bold
        add XName & "." & "Font.Size" & " = " & XControl.Font.Size
        add XName & "." & "Font.Italic" & " = " & XControl.Font.Italic
        add XName & "." & "Font.Name" & " = '" & XControl.Font.Name & "'"
        
        If XControl.Picture <> 0 Then
            SavePicture XControl.Picture, Path & "\" & XName & ".bmp"
            add "Set " & XName & "." & "Picture" & " = CForm.LoadImage('%path%\" & XName & ".bmp')"
        End If
    
        
    End If
    
    If TypeOf XControl Is Image Then
        add XName & "." & "Visible" & " = " & XControl.Visible
        
        If XControl.Picture <> 0 Then
            SavePicture XControl.Picture, Path & "\" & XName & ".bmp"
            add "Set " & XName & "." & "Picture" & " = CForm.LoadImage('%path%\" & XName & ".bmp')"
        End If
    
    End If
    
    If TypeOf XControl Is OptionButton Then
        add XName & "." & "Caption" & " = '" & XControl.Caption & "'"
        add XName & "." & "Font.Bold" & " = " & XControl.Font.Bold
        add XName & "." & "Font.Size" & " = " & XControl.Font.Size
        add XName & "." & "Font.Italic" & " = " & XControl.Font.Italic
        add XName & "." & "Font.Name" & " = '" & XControl.Font.Name & "'"
        add XName & "." & "Visible" & " = " & XControl.Visible
        
        add XName & "." & "BackColor" & " = " & XControl.BackColor
'        add XName & "." & "Cancel" & " = " & XControl.Cancel
        add XName & "." & "Enabled" & " = " & XControl.Enabled
'        add XName & "." & "Default" & " = " & XControl.Default
        add XName & "." & "TabIndex" & " = " & XControl.TabIndex
        add XName & "." & "Value" & " = " & XControl.Value
    
    End If
    
    If TypeOf XControl Is CheckBox Then
        add XName & "." & "Caption" & " = '" & XControl.Caption & "'"
        add XName & "." & "Font.Bold" & " = " & XControl.Font.Bold
        add XName & "." & "Font.Size" & " = " & XControl.Font.Size
        add XName & "." & "Font.Italic" & " = " & XControl.Font.Italic
        add XName & "." & "Font.Name" & " = '" & XControl.Font.Name & "'"
        add XName & "." & "Visible" & " = " & XControl.Visible
        
        add XName & "." & "BackColor" & " = " & XControl.BackColor
'        add XName & "." & "Cancel" & " = " & XControl.Cancel
        add XName & "." & "Enabled" & " = " & XControl.Enabled
'        add XName & "." & "Default" & " = " & XControl.Default
        add XName & "." & "TabIndex" & " = " & XControl.TabIndex
        add XName & "." & "Value" & " = " & XControl.Value
    End If
    
    If TypeOf XControl Is Label Then
        add XName & "." & "Caption" & " = '" & Replace(XControl.Caption, vbCrLf, "' & vbCrLf & '") & "'"
        add XName & "." & "Font.Bold" & " = " & XControl.Font.Bold
        add XName & "." & "Font.Size" & " = " & XControl.Font.Size
        add XName & "." & "Font.Italic" & " = " & XControl.Font.Italic
        add XName & "." & "Font.Name" & " = '" & XControl.Font.Name & "'"
        add XName & "." & "Visible" & " = " & XControl.Visible
        
'        add XName & "." & "MultiLine" & " = " & XControl.MultiLine
        
        add XName & "." & "BackColor" & " = " & XControl.BackColor
        add XName & "." & "ForeColor" & " = " & XControl.ForeColor
        add XName & "." & "Enabled" & " = " & XControl.Enabled
        
    End If
    
    If TypeOf XControl Is ListBox Then
        add XName & "." & "Font.Bold" & " = " & XControl.Font.Bold
        add XName & "." & "Font.Size" & " = " & XControl.Font.Size
        add XName & "." & "Font.Italic" & " = " & XControl.Font.Italic
        add XName & "." & "Font.Name" & " = '" & XControl.Font.Name & "'"
        add XName & "." & "Visible" & " = " & XControl.Visible
        
'        add XName & "." & "MultiLine" & " = " & XControl.MultiLine
        
        add XName & "." & "BackColor" & " = " & XControl.BackColor
        add XName & "." & "ForeColor" & " = " & XControl.ForeColor
        add XName & "." & "Enabled" & " = " & XControl.Enabled
        
    End If
    If TypeOf XControl Is ComboBox Then
    
    add XName & "." & "Font.Bold" & " = " & XControl.Font.Bold
        add XName & "." & "Font.Size" & " = " & XControl.Font.Size
        add XName & "." & "Font.Italic" & " = " & XControl.Font.Italic
        add XName & "." & "Font.Name" & " = '" & XControl.Font.Name & "'"
        add XName & "." & "Visible" & " = " & XControl.Visible
        
        add XName & "." & "Text" & " = '" & XControl.text & "'"
        
        add XName & "." & "BackColor" & " = " & XControl.BackColor
        add XName & "." & "ForeColor" & " = " & XControl.ForeColor
        add XName & "." & "Enabled" & " = " & XControl.Enabled
        add XName & "." & "Locked" & " = " & XControl.Locked
    End If
    
    If TypeOf XControl Is Shape Then
        add XName & "." & "Shape" & " = " & XControl.Shape
        add XName & "." & "BackColor" & " = " & XControl.BackColor
        add XName & "." & "FillColor" & " = " & XControl.FillColor
        add XName & "." & "FillStyle" & " = " & XControl.FillStyle
        
        add XName & "." & "BackStyle" & " = " & XControl.BackStyle
        add XName & "." & "BorderStyle" & " = " & XControl.BorderStyle
        add XName & "." & "BorderWidth" & " = " & XControl.BorderWidth
        add XName & "." & "DrawMode" & " = " & XControl.DrawMode
        add XName & "." & "Visible" & " = " & XControl.Visible
        
    End If
    
    
    add ""
skip2:
Next






add ""

Open Target For Output As #1

data = Replace(data, "'", Chr(34))
'data = Replace(data, Chr(34) & Chr(34), "'")
data = Replace(data, "&fid", frm.Name & "ID")
data = Replace(data, "&frm", frm.Name)
Print #1, defs
Print #1, data

Close #1
End
End Function


Function add(what As String)
data = data & vbCrLf & what
End Function

Function FormatText(text As String) As String
Dim buf As String
buf = text
buf = Replace(buf, vbCrLf, "' & vbCrLf & '")
buf = Replace(buf, vbCr, "' & vbCr & '")
buf = Replace(buf, vbLf, "' & vbLf & '")
buf = Replace(buf, Chr(34), "''")

FormatText = "'" & buf & "'"
End Function
