<div align="center">

## Create Random Control at Runtime


</div>

### Description

Create a label, commandbutton, frame, textbox, hscrollbar, listbox, picturbox, shape, dirlistbox, filelistbox, drivelistbox, vscrollbar, optionbutton, line, checkbox, image or combobox randomly with random height, width, top & left properties
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Max \- Demian Net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/max-demian-net.md)
**Level**          |Advanced
**User Rating**    |4.5 (36 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/max-demian-net-create-random-control-at-runtime__1-8881/archive/master.zip)





### Source Code

```
'Make a command1, try to make it smal & it the bottom right hand corner for best results.
 Private WithEvents txtDynamic As TextBox
 Private Sub Command1_Click()
 On Error Resume Next
 Dim RandomControl(1 To 18) As String
 Dim i As Integer
 Randomize
 RandomControl(1) = "VB.TextBox"
 RandomControl(2) = "VB.CommandButton"
 RandomControl(3) = "VB.Shape"
 RandomControl(4) = "VB.Label"
 RandomControl(5) = "VB.ListBox"
 RandomControl(6) = "VB.PictureBox"
 RandomControl(7) = "VB.Frame"
 RandomControl(8) = "VB.HScrollBar"
 RandomControl(9) = "VB.VScrollBar"
 RandomControl(10) = "VB.Image"
 RandomControl(11) = "VB.Line"
 RandomControl(12) = "VB.DirListBox"
 RandomControl(13) = "VB.DriveListBox"
 RandomControl(14) = "VB.FileListBox"
 RandomControl(15) = "VB.Timer"
 RandomControl(16) = "VB.ComboBox"
 RandomControl(17) = "VB.OptionButton"
 RandomControl(18) = "VB.CheckBox"
 i = Int((18 * Rnd) + 1)
 RandomTop = Int(Rnd * Me.Height)
 RandomLeft = Int(Rnd * Me.Width)
 RandomWidth = Int(Rnd * Me.Height)
 RandomText = Int(Rnd * 3200)
 Set RandDynamic = Controls.Add(RandomControl(i), "Rand" & RandomText)
   With RandDynamic
     .Visible = True
     .Text = "Demian Net"
     .Caption = "Demian Net"
     .BackColor = vbRed
     .Width = RandomWidth
     .Top = RandomTop
     .Left = RandomLeft
   End With
 End Sub
```

