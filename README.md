<div align="center">

## Split String for Word Wrapping


</div>

### Description

<p>Breaks up a string so that it can be effectively printed - word wrapped - using the <code>Print</code> statement.</p>

<p><code>SplitLines</code> is a function in Visual Basic to return an array of strings from a long string such that the each array element has its <code>P.TextWidth(Lines(i)) &lt; W</code>. The function uses the current font settings of the object <code>P</code> which could be a Form, a PictureBox or the Printer object.</p>
 
### More Info
 
<pre>

Txt     -> is the String that is to be split

P      -> a Form, a PictureBox or the Printer object. The font settings

of this will be used to determine the TextWidth

W      -> the maximum width of the string array

</pre>

Example Usage

----

<pre>

Dim Ltxt() as string

Dim OriStr as string

OriStr = "This contains the string that is to be split...... ..."

Ltxt = SplitLines(OriStr, Form1, 1500)

For i = 1 to UBound(Ltxt)

Form1.Print Ltxt(i)

Next i

</pre>

An array of strings.

The function counts on the fact that font characteristics for the object P has been set.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gajendra S\. Dhir](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gajendra-s-dhir.md)
**Level**          |Intermediate
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gajendra-s-dhir-split-string-for-word-wrapping__1-28744/archive/master.zip)





### Source Code

```
Public Function SplitLines(Txt As String, P As Object, W As Single) As String()
Dim Lines() As String, CurrW As Single, CurrWord As String
Dim L As Integer, i As Integer, WCnt As Integer
CurrW = 0
L = Len(Txt)
If (P.TextWidth(Txt) > W) Or (InStr(Txt, vbCr) > 0) Then
	i = 1
	WCnt = 1
	ReDim Lines(WCnt) As String
	Do Until i > L
		CurrWord = ""
		Do Until i > L Or Mid(Txt, i, 1) <= " "
			CurrWord = CurrWord & Mid(Txt, i, 1)
			i = i + 1
		Loop
		If CurrW + P.TextWidth(CurrWord) > W Then
			WCnt = WCnt + 1
			ReDim Preserve Lines(WCnt) As String
			CurrW = 0
		End If
		Lines(WCnt) = Lines(WCnt) + CurrWord
		CurrW = P.TextWidth(Lines(WCnt))
		Do Until i > L Or Mid(Txt, i, 1) > " "
			Select Case Mid(Txt, i, 1)
			Case " "
				Lines(WCnt) = Lines(WCnt) + " "
				CurrW = P.TextWidth(Lines(WCnt))
			Case vbLf
			Case vbCr
				WCnt = WCnt + 1
				ReDim Preserve Lines(WCnt) As String
				CurrW = 0
			Case Chr(9)
				Lines(WCnt) = Lines(WCnt) + " "
				CurrW = P.TextWidth(Lines(WCnt))
			End Select
			i = i + 1
		Loop
	Loop
Else
	ReDim Lines(1) As String
	Lines(1) = Txt
End If
For i = 1 To WCnt
  Lines(i) = LTrim(RTrim(Lines(i)))
Next i
SplitLines = Lines
End Function
```

