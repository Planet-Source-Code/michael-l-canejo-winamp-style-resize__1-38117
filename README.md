<div align="center">

## Winamp Style Resize


</div>

### Description

Winamp 2.80 and below had a resizing feature on it's playlist that resized the form in blocks of pixels instead of per pixel. The reason for it was because they didn't know how to stretch the playlist skins (i hope not), so instead they kept resizing it a certain amount of pixels and then painting the image block after it's last one. This does get you better quality on your skinning cause stretching would degrade the quality unless it wasn't made to be stretched. I am using this in an XP Control i'm working on right now because the vb Image control flickers when the stretch property is set to true and you want to resize it with a form. I figured resizing it per chunck of pixels rather than per pixel, it would reduce flickering and save resources on old computers. I havn't testing my theory, but this method still looks appealing to me. I think this would work great on borderless forms or unsizable forms for customizing. It's kinda sloppy but efficient ;)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael L\. Canejo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-l-canejo.md)
**Level**          |Beginner
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-l-canejo-winamp-style-resize__1-38117/archive/master.zip)





### Source Code

```
'Just paste it into a brand new form
'or a old one and it will add cursor
'resizers based on coordinates which
'is really nice i think. There was no
'point in zipping such a simple feature.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iTwipX, iTwipY
Dim iScale As Integer
 On Error Resume Next
 iScale = Me.ScaleMode
 iTwipX = Screen.TwipsPerPixelX
 iTwipY = Screen.TwipsPerPixelY
 Me.ScaleMode = 3
 If Button = 1 Then
 If MousePointer = 0 Then GoTo Fix_Scale
 If MousePointer = 7 Then GoTo NS_Resize
 If X / iTwipX > ScaleWidth + iTwipX Then
  Width = Width + iTwipX ^ 2 * 2
 ElseIf X / iTwipX < ScaleWidth - iTwipX Then
  Width = Width - iTwipX ^ 2 * 2
 End If
 If MousePointer = 9 Then GoTo Fix_Scale
NS_Resize:
 If Y / iTwipY > ScaleHeight + iTwipY Then
  Height = Height + iTwipY ^ 2 * 2
 ElseIf Y / iTwipY < ScaleHeight - iTwipY Then
  Height = Height - iTwipY ^ 2 * 2
 End If
 Else
 If X / iTwipX > ScaleWidth - 12 _
 And Y / iTwipY > ScaleHeight - 10 Then
  MousePointer = 8
 ElseIf X / iTwipX > ScaleWidth - 12 _
 And Y / iTwipY < ScaleHeight - 10 Then
  MousePointer = 9
 ElseIf X / iTwipX < ScaleWidth - 12 _
 And Y / iTwipY > ScaleHeight - 10 Then
  MousePointer = 7
 Else
  MousePointer = 0
 End If
 End If
Fix_Scale:
 Me.ScaleMode = iScale%
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MousePointer = 0
End Sub
```

