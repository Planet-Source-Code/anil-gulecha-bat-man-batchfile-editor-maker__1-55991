Attribute VB_Name = "ModMsgBox"
' M=middle
' T=top
' L=left
' D=down
' R=right
Public TL$, TR$, LL$, LR$, MT$, MR$, ML$, MD$, BoxWidth As Integer, BoxTitle As String

'This function is the star.It returns the various lines
'for the box command
'You'll have to understand it your self !!

Function Lline(Optional st As String = "") As String
' if st="`" return first line
If st = "`" Then
Dim BetweenStr As String, FreeInStr As Byte
If Trim(BoxTitle) = "" Then Lline = TL & String(BoxWidth - 2, MT) & TR: Exit Function

FreeInStr = BoxWidth - 5 - Len(BoxTitle)
BetweenStr = MT & " " & BoxTitle & " " & String(FreeInStr, MT)
Lline = TL & BetweenStr & TR

Exit Function
End If


'if st="" then return plain line
If st = "" Then
Lline = ML & String(BoxWidth - 2, " ") & MR
Exit Function
End If

'if st="~" return last line
If st = "~" Then
Lline = LL & String(BoxWidth - 2, MD) & LR
Exit Function
End If

'else return line with text
Dim FreeSpace As Byte, Prt1 As Byte, Prt2 As Byte
FreeSpace = (BoxWidth - 2) - Len(st)
Prt1 = FreeSpace \ 2
Prt2 = FreeSpace - Prt1
Lline = ML & String(Prt1, " ") & st & String(Prt2, " ") & MR

End Function

'Sets the msgbox Asci  code and title
Function SetBoxStuff(TopLeft As Byte, MidTop As Byte, TopRight As Byte, MidRight As Byte, _
bottomright As Byte, MidBottom As Byte, BottomLeft As Byte, Midleft As Byte, Tit As String) As Boolean
' TL$, TR$, LL$, LR$, MT$, MR$, ML$, MD$, BoxWidth As Integer
On Error GoTo sbs_end:

TL = Chr(TopLeft)
TR = Chr(TopRight)
LL = Chr(BottomLeft)
LR = Chr(bottomright)
MT = Chr(MidTop)
MR = Chr(MidRight)
ML = Chr(Midleft)
MD = Chr(MidBottom)
SetBoxStuff = True
BoxTitle = Tit

Exit Function

sbs_end:
SetBoxStuff = False

End Function

