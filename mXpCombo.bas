Attribute VB_Name = "mXpCombo"
Option Explicit
Private cb() As cXPCombo

Public Sub DrawTheXPCombos(frm As Form, _
                            Optional FrameClr As OLE_COLOR = -1, _
                            Optional FrameClrHot As OLE_COLOR = -1, _
                            Optional FrameClrDisabled As OLE_COLOR = -1, _
                            Optional ArrowEnColor As OLE_COLOR = -1, _
                            Optional ArrowDisColor As OLE_COLOR = -1, _
                            Optional NormalThumb As StdPicture, _
                            Optional HotThumb As StdPicture, _
                            Optional DropedThumb As StdPicture, _
                            Optional DissThumb As StdPicture, _
                            Optional DrawThumb As Boolean = True)
Dim CTL As Control
Dim i As Long
    For Each CTL In frm.Controls
        If TypeOf CTL Is ComboBox Then
            ReDim Preserve cb(i)
            Set cb(i) = New cXPCombo
            cb(i).MakeXP CTL
            cb(i).CboBackColor = CTL.BackColor
            cb(i).FrameColorHot = FrameClrHot
            cb(i).FrameColorNormal = FrameClr
            cb(i).FrameColorDiss = FrameClrDisabled
            cb(i).ArrowEnabledColor = ArrowEnColor
            cb(i).ArrowDisabledColor = ArrowDisColor
            cb(i).ShowArrow = DrawThumb
            Set cb(i).PictureNormalThumb = NormalThumb
            Set cb(i).PictureHotThumb = HotThumb
            Set cb(i).PictureDropedThumb = DropedThumb
            Set cb(i).PictureDissThumb = DissThumb
            i = i + 1
        End If
    Next CTL
End Sub
