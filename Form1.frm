VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTextRotate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotator: Multiple Line Captions with Accelerator Keys"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFont 
      Caption         =   "Change Font (Select True-Type Only)"
      Height          =   375
      Left            =   255
      TabIndex        =   14
      Top             =   4920
      Width           =   5550
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5430
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   2085
      TabIndex        =   8
      Top             =   2985
      Width           =   3585
      Begin VB.OptionButton optAlignV 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bottom Aligned"
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAlignV 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Center Aligned"
         Height          =   315
         Index           =   2
         Left            =   1620
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton optAlignV 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Top Aligned"
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   9
         Tag             =   "2"
         Top             =   -15
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Alignment (Vertical)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   1
         Left            =   1470
         TabIndex        =   12
         Top             =   30
         Width           =   1980
      End
   End
   Begin VB.CheckBox chkDegrees 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vertical @ 270 degrees"
      Height          =   345
      Index           =   1
      Left            =   3705
      TabIndex        =   7
      Top             =   3615
      Width           =   2055
   End
   Begin VB.CheckBox chkDegrees 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vertical @ 90"
      Height          =   345
      Index           =   0
      Left            =   2130
      TabIndex        =   6
      Top             =   3615
      Value           =   1  'Checked
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   540
      Left            =   2055
      TabIndex        =   1
      Top             =   2325
      Width           =   3585
      Begin VB.OptionButton optAlignment 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Right Aligned"
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   225
         Width           =   1350
      End
      Begin VB.OptionButton optAlignment 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Center Aligned"
         Height          =   315
         Index           =   2
         Left            =   1650
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton optAlignment 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Left Aligned"
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Tag             =   "2"
         Top             =   -30
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Alignment (Horizonal)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   0
         Left            =   1455
         TabIndex        =   5
         Top             =   -15
         Width           =   1980
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   4035
      Width           =   5550
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Think of these rectangles as very large buttons..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   315
      TabIndex        =   13
      Top             =   90
      Width           =   5280
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderStyle     =   5  'Dash-Dot-Dot
      Height          =   1725
      Left            =   2085
      Top             =   450
      Width           =   3525
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderStyle     =   5  'Dash-Dot-Dot
      Height          =   3525
      Left            =   255
      Top             =   450
      Width           =   1725
   End
End
Attribute VB_Name = "frmTextRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project required for new vertical buttons to be posted soon
' 27 Feb 2004

' Printing rotated text is difficult. Functions like TextOut & DrawText are
' very limited, especially if you want to use hotkeys (i.e., &File)
' Neither API will print the underscore below the hotkey in rotated text.

' Now, if I'm wrong, please send me proof so I can reduce 100's of lines
' of code down to only a few & I thank you in advance

' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/fontext_0odw.asp
' Quoted comment from above source follows:
' "Remarks: The DrawText function supports only fonts whose escapement
' and orientation are both zero." << this means no angles other than 0 degrees

' As you will see, MSDN isn't completely truthful. Getting DrawText to work
' with non-zero escapement/orientation is doable, but not easy.


' In order to get the hotkey underlined & print text in every conceivable
' alignment combination, I built a mini-wordprocessor (BreakWords).
' That wordprocessor breaks captions/strings into individual words & then
' reassembles them to determine where the line breaks would occur.
' The APIs don't word-break vertical strings properly.

' Then the lines of the captions/strings are tested (SplitLines) to determine
' where a clipping action would occur if the number of lines were too many for
' the area to print into/onto.
' The APIs don't do this either for vertical strings.

' Last but not least, tracking the exact position of the hotkey is a bit
' difficult since it can literally be in any X,Y coordinate or it may even
' be clipped which would prevent it from being printed.

' The remarks aren't too bad in this project, but following code for
' vertical string manipulation isn't easy for anyone.

' I designed the core functions to be portable to other applications should
' you decide to use the code. I would suggest tweaking the code to not allow
' routines to clip text. Rather you could add a clipping region to
' whatever DC you are printing to and let it control the clipping.
' Personal preferences.

' Last but not least... This project was only designed to track hotkeys for
' rotated text that is either 0, 90 or 270 degrees. Any other angles of
' rotation would require some serious Trigonometry & I just don't feel like it :)


' Main function provided:  RotateText
' Supporting Functions:  FormatForHotKey, BreakWords, SplitLines

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' Font APIs
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
' Rectangle Calculations APIs
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
' Text Drawing API
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_BOTTOM As Long = &H8
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_NOCLIP As Long = &H100
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_RIGHT As Long = &H2
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_TOP As Long = &H0
Private Const DT_VCENTER As Long = &H4
Private Const DT_WORDBREAK As Long = &H10


Private Sub RotateText(ByVal lDC As Long, ByVal fontAngle As Long, _
            ByVal hAlign As Long, ByVal vAlign As Long, _
            ByVal sText As String, X1 As Long, Y1 As Long, _
            X2 As Long, Y2 As Long)
            
' Core function. this function only handles fonts with orientations of
' 270, 90 & 0 degrees.

' [IN]
' lDC=the device context to print to. It must already have the proper font selected
' fontAngle=the font orientation, either 0, 90 or 270 only
' hAlign=0,1 or 2 for left, right or center justification respectively
' vAlign=0,1,or 2 for top, bottom or center alignment respectively
' sText=the caption or string to be processed
' X1=left boundar of rectangle to print text to
' Y1=top boundary of rectangle to print text to
' X2=right boundary of rectangle to print text to
' Y2=bottom boundary of rectangle to print text to
            
Dim tRect As RECT, cRect As RECT, hRect As RECT
Dim Xoffset As Long, yOffset As Long
Dim sLines() As String, sCurLine As String
Dim hKeyIndex As Integer, hKeyOffset As Long
Dim maxWidth As Long, maxHeight As Long, sCaption As String
Dim singleLine As Long, lFlags As Long

'sanity checks
If Len(sText) = 0 Then Exit Sub
If fontAngle > 0 And fontAngle < 180 Then fontAngle = 90
If fontAngle > 179 Then fontAngle = 270

If vAlign < 0 Then vAlign = 0
If vAlign > 2 Then vAlign = 2

If hAlign < 0 Then hAlign = 0
If hAlign > 2 Then hAlign = 2

' calculate height of a single character. Used for positioning rectangles
DrawText lDC, "X1", 1, cRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
singleLine = cRect.Bottom

If fontAngle = 0 Then
    ' horizontal fonts are fairly simple. But when multiple lines are present,
    ' the API doesn't allow use of DT_Top, DT_Bottom, DT_VCenter & a few more.
    ' Therefore, we need to position the rectangle vertically, but API will
    ' handle the horizontal portion for us and even do the clipping
    SetRect cRect, X1, Y1, X2, Y2
    lFlags = DT_WORDBREAK
    DrawText lDC, sText, Len(sText), cRect, lFlags Or DT_CALCRECT
    If cRect.Bottom > singleLine Then
        ' we have multiple lines, so we have to move our rectangle before we print
        If cRect.Bottom > Y2 Then
            cRect.Bottom = Y2       ' too much text, some will get clipped
        Else
            If vAlign = 2 Then      ' center alignment
                OffsetRect cRect, 0, -cRect.Top + ((Y2 - Y1) - (cRect.Bottom - cRect.Top)) \ 2 + Y1
            Else                    ' bottom alignment
                If vAlign = 1 Then OffsetRect cRect, 0, -cRect.Top + Y2 - (cRect.Bottom - cRect.Top)
            End If
        End If
        'for horizontal alighment, we simply reset the left/right to min/max
        SetRect cRect, X1, cRect.Top, X2, cRect.Bottom
    Else
        ' easy one liner, send on through after we abuse the flags
        lFlags = lFlags Or Choose(vAlign + 1, DT_TOP, DT_BOTTOM, DT_VCENTER)
        lFlags = lFlags Or DT_SINGLELINE ' DT_SingleLine overrides DT_WordBreak
    End If
    lFlags = lFlags Or Choose(hAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
    DrawText lDC, sText, Len(sText), cRect, lFlags

Else

    maxWidth = X2 - X1
    maxHeight = Y2 - Y1
    ' find the accelerator key & remove/replace unneeded ampersands
    sCaption = FormatForHotKey(sText)
    
    ' break caption into vbCRLF delimited words
    BreakWords lDC, sCaption, maxHeight
    ' determine overall rectangle size & build multiple lines, if needed
    SplitLines lDC, sCaption, sLines(), maxHeight, maxWidth
    
    If vAlign = 2 Then   ' centered 90/270 degree
        Xoffset = ((X2 - X1) - maxWidth) \ 2 - singleLine + X1
    End If
    If fontAngle = 270 Then
        yOffset = Y1  ' 270 degree top offset
        If vAlign = 1 Then Xoffset = (X2 - X1) - maxWidth
        If Xoffset < 0 Then Xoffset = 0 ' remove if using a clipping region
        Xoffset = X2 - Xoffset
    Else
        yOffset = Y2 ' 90 degree top offset
        If vAlign = 1 Then Xoffset = X2 - maxWidth - X1
        If Xoffset < 0 Then Xoffset = 0 ' remove if using a clipping region
        Xoffset = Xoffset + X1
    End If
    SetRect tRect, Xoffset, yOffset, Xoffset + 1, yOffset + 1
    
    Dim I As Integer
    For I = 0 To UBound(sLines)
        ' when I ran caption through mini-word processor, I removed all
        ' ampersands (&) and replaced them with chr$(8). This way the only
        ' ampersand is the hotkey for the caption
        sCurLine = sLines(I)
        hKeyIndex = InStr(sCurLine, "&")
        ' see if hotkey exists & replace the ampersand if needed
        If hKeyIndex Then sCurLine = Replace$(sCurLine, "&", "")
        ' now add back any other ampersands... so they can be measured/printed
        sCurLine = Replace$(sCurLine, Chr$(8), "&")
        
        ' measure the line of text
        DrawText lDC, sCurLine, Len(sCurLine), cRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
        
        Select Case hAlign
        Case 1 ' right aligned
            If fontAngle = 90 Then
                OffsetRect tRect, 0, -(maxHeight - cRect.Right)
            Else
                OffsetRect tRect, 0, (maxHeight - cRect.Right)
            End If
        Case 2 ' centered between top & bottom of rectangle
            If fontAngle = 90 Then
                OffsetRect tRect, 0, -(maxHeight - cRect.Right) \ 2
            Else
                OffsetRect tRect, 0, (maxHeight - cRect.Right) \ 2
            End If
        Case Else ' left aligned - no extra offsets
        End Select
        tRect.Bottom = tRect.Top + cRect.Bottom
        
        If hKeyIndex > 0 And hKeyIndex < Len(sCurLine) + 1 Then
            ' here we track relative position of the character that is the hotkey
            ' first we measure everything up to that character
            DrawText lDC, Left$(sCurLine, hKeyIndex), Len(Left$(sCurLine, hKeyIndex)), hRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
            hKeyOffset = hRect.Right
            ' now we measure just that character
            DrawText lDC, Mid$(sCurLine, hKeyIndex, 1), 1, hRect, DT_CALCRECT Or DT_NOCLIP
            ' depending on text angles, we track it differently
            If fontAngle = 270 Then
                hKeyOffset = hKeyOffset - hRect.Right
                OffsetRect hRect, 0, tRect.Left - cRect.Bottom * 2
            Else          ' 90
                hKeyOffset = hRect.Right - hKeyOffset
                OffsetRect hRect, 0, tRect.Left
            End If
            hRect.Top = hRect.Bottom - 1
        End If
        
        ' now we can draw our rotated text.
        DrawText lDC, sCurLine, Len(sCurLine), tRect, DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP
        
        ' finalize tracking the hotkey & offseting the next line to be printed
        ' Note: the 1 at end of offSetRect calls is to separate the underscore
        ' a little further from the character. I feel VB puts it too close.
        ' Change the 1 to 0 in the two lines below if you want to.
        If fontAngle = 90 Then
            If hKeyIndex Then OffsetRect hRect, tRect.Top + hKeyOffset - hRect.Right, 1
            OffsetRect tRect, cRect.Bottom, -tRect.Top + yOffset
        Else          '270
            If hKeyIndex Then OffsetRect hRect, tRect.Top + hKeyOffset, 1
            OffsetRect tRect, -cRect.Bottom, -tRect.Top + yOffset
        End If
        
    Next
    ' after all lines are printed, go back & print the underscore for the hotkey
    ' Note: by using a rectangle, we don't need to create/destroy an underlined
    ' font & try to exactly overwrite the exiting hotkey.
    If hRect.Right Or hRect.Bottom Then
        Rectangle lDC, hRect.Top, hRect.Left, hRect.Bottom, hRect.Right
    End If
    
    ' End Note:
    ' If this routine was really used to draw captions for buttons, menus, etc,
    ' then you probably wouldn't want to send it through the 2 helper functions
    ' (FormatForHotKey, BreakWords) every time it is printed. I would suggest
    ' processing the caption thru the 2 helper functions outside of this function
    ' only when the caption/font changes or the orientation changes. Then pass the
    ' sLines() array to this function.  Unless other variables are also passed to
    ' this routine, the caption will still need to be processed by SplitLines
    ' within this function. If you chose to go that way, then you wouldn't want
    ' to clear the sLines() array.
    Erase sLines()
End If

End Sub

Private Sub BreakWords(ByVal lDC As Long, sCaption As String, maxWidth As Long)
' function will break a string into words based off of occurrences of either a
' space or carriage return. This routine seems to exactly replicate the
' DrawText API's function with the DT_WordBreak flag active

Dim sWords() As String, iBreaks(0 To 2) As Integer, I As Integer
Dim sLine As String, sBreakChar As String, J As Integer, nrWords As Integer
Dim firstBreak As Integer, tRect As RECT, curWidth As Long
Dim chrOffset As Integer

' cache the caption, we will overwrite it later
sLine = sCaption
' keep track of last word break position
chrOffset = 1
Do While Len(sLine)
    ' reset these variables to defaults
    iBreaks(0) = Len(sLine) + 1
    firstBreak = 0
    
    ' Note to self: increment the array UBound in Declare if other linebreaks are to be tested
    For J = 1 To UBound(iBreaks)
        ' loop thru each line-break character(s)
        sBreakChar = Choose(J, " ", vbCrLf)
        ' determine if any linebreaks are in the current string
        iBreaks(J) = InStr(chrOffset, sLine, sBreakChar)
        ' if so, keep track of the one occuring first in the string
        If iBreaks(J) < iBreaks(firstBreak) And iBreaks(J) > 0 Then firstBreak = J
    Next
    ' increment our word array
    ReDim Preserve sWords(0 To nrWords)
    If firstBreak Then
        ' a linebreak found; add the preceding word to the array
        sWords(nrWords) = Left$(sLine, iBreaks(firstBreak) - 1)
        ' remove the word from the current string
        sLine = Mid$(sLine, Len(sWords(nrWords)) + 1)
        ' update the character position marker
        chrOffset = firstBreak + 1
        ' increment the word count
        nrWords = nrWords + 1
    Else
        ' no more linebreaks found; add last word to the array
        sWords(nrWords) = sLine
        Exit Do
    End If
Loop
' reset return variable & temp variable
sCaption = ""
sLine = ""

' now loop thru the words, measuring them & adding them to individual lines
' first remove any trailing vbCRLF's
For I = nrWords To 1 Step -1
    If sWords(I) <> vbCrLf Then Exit For
Next
If I < nrWords Then ReDim Preserve sWords(0 To I)
nrWords = I
For I = 0 To nrWords
    ' reset the Right value; used for word width
    tRect.Right = 0
    ' test for carriage returns; handled separately
    If InStr(sWords(I), vbCrLf) Then
        ' append the return variable
        sCaption = sCaption & sWords(I)
        ' start a new line with the word after the carriage return
        sLine = Mid$(sWords(I), InStr(sWords(I), vbCrLf) + 2)
    Else
        ' measure the word & test it against the max width
        DrawText lDC, (sLine & sWords(I)), Len(sLine & sWords(I)), tRect, DT_CALCRECT
        If tRect.Right < maxWidth + 1 Then
            ' word added to the line is still < than the max width, so simply add the word to the line
            sLine = sLine & sWords(I)
            ' update the return variable
            sCaption = sCaption & sWords(I)
        Else
            ' the word added to the line > than the max width
            ' update the reutrn variable & start a new blank line
            sCaption = sCaption & vbCrLf & LTrim$(sWords(I))
            sLine = LTrim$(sWords(I))
        End If
    End If
Next
Erase sWords
Erase iBreaks
' exit the routine; returning the new wordwrapped caption
End Sub

Private Sub SplitLines(ByVal lDC As Long, ByVal sCaption As String, sLines() As String, maxWidth As Long, maxHeight As Long)
' Function splits a single string into multiple strings based on carriage returns
' This function also aids clipping decisions by tallying up how "tall" the
' entire caption is when formatted against the passed rectangle dimensions.

Dim I As Integer, xRect As RECT, totalHT As Long
Dim hotKey As Byte, hotKeyPos As Byte, lineLen As Long
Dim J As Integer, K As Integer

' initialize the array
sLines = Split(sCaption, vbCrLf)

For I = 0 To UBound(sLines)
    
    ' now calculate the width/height of the current line
    DrawText lDC, sLines(I), Len(sLines(I)), xRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP
    
    ' keep tally of total height
    totalHT = totalHT + xRect.Bottom
    
    ' see if this line exceeds rectangle dimensions
    If totalHT > maxHeight Then
            
        ReDim Preserve sLines(0 To I - 1)
        Exit For
        
    End If

Next
' return the total height of the caption (less any lines that were clipped)
If totalHT Then maxHeight = totalHT
End Sub


Private Function FormatForHotKey(ByVal sCaption As String) As String
' Formats a caption for display. The following formatting takes place.
' 1. Replace all double ampersands (&&) with chr$(8) for replacement later
' 2. Find the true hotkey/accelerator key & remove all other ampersands
' 3. Return a copy of the caption with the above format
' The above is done ahead of time to ensure proper measurements when
' aligning text to be printed. Characters removed will be added back when needed

Dim I As Integer, J As Integer, lastAmp As Integer
' By accepted rules, if more than one ampersand is in caption, then
' only the last instance of a single ampersand identifies the hot key
I = InStr(sCaption, "&")
Do Until I = 0
    If Mid$(sCaption, I, 2) = "&&" Then
        ' we will use the "backspace" chr$(8) for a placeholder
        sCaption = Left$(sCaption, I - 1) & Chr$(8) & Mid$(sCaption, I + 2)
    Else
        If lastAmp Then
            ' already had a hotkey, need to remove it & set the new one
            sCaption = Left$(sCaption, lastAmp - 1) & Mid$(sCaption, lastAmp + 1)
            I = I - 1
        End If
        ' set the hotkey
        lastAmp = I
    End If
    ' check for another instance
    I = InStr(I + 1, sCaption, "&")
Loop
If lastAmp Then
    ' set accesskey for usercontrols
End If
FormatForHotKey = sCaption
End Function


'/////////// FOLLOWING CODE ONLY USED FOR THE EXAMPLE \\\\\\\\\\\\\\\\

Private Sub chkDegrees_Click(Index As Integer)
chkDegrees(Abs(Index - 1)).Value = Abs(chkDegrees(Index).Value - 1)
If chkDegrees(Index) Then Call DoTheWork
End Sub

Private Function LoadFont(Angle As Long) As Long

If Angle < 0 Then Exit Function

    Dim newFont As LOGFONT
    newFont.lfCharSet = 1
    newFont.lfWeight = Abs(Text1.FontBold) * 300 + 400
    newFont.lfItalic = Abs(Text1.FontItalic)
    newFont.lfUnderline = Abs(Text1.FontUnderline)
    newFont.lfFaceName = Text1.FontName & Chr$(0)
    ' Font Size?...
    ' whatever your preference... seems to be the same results with either function
        'newFont.lfHeight = (Me.FontSize * -20) / Screen.TwipsPerPixelY
    ' or
        'newFont.lfHeight = -MulDiv((Me.FontSize), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
    ' I prefer this one since it doesn't require 2 additional API calls...
    newFont.lfHeight = (Text1.FontSize * -20) / Screen.TwipsPerPixelY
    newFont.lfEscapement = Angle * 10
    newFont.lfOrientation = Angle * 10
    ' create a memory font
    LoadFont = CreateFontIndirect(newFont)
End Function

Private Sub cmdFont_Click()
With CommonDialog1
    .FontBold = Text1.FontBold
    .FontItalic = Text1.FontItalic
    .FontUnderline = Text1.FontUnderline
    .FontName = Text1.FontName
    .FontSize = Text1.FontSize
    .FontStrikethru = Text1.FontStrikethru
    .Flags = cdlCFScalableOnly Or cdlCFTTOnly Or cdlCFWYSIWYG Or cdlCFBoth
    .CancelError = True
End With
On Error GoTo ExitRoutine
CommonDialog1.ShowFont
With CommonDialog1
    Text1.FontBold = .FontBold
    Text1.FontItalic = .FontItalic
    Text1.FontName = .FontName
    Text1.FontSize = .FontSize
    Text1.FontStrikethru = .FontStrikethru
    Text1.FontUnderline = .FontUnderline
End With
Call DoTheWork
ExitRoutine:
End Sub

Private Sub Form_Load()
Show
Call DoTheWork
End Sub

Private Sub optAlignment_Click(Index As Integer)
If optAlignment(Index) Then optAlignment(0).Tag = Index
Call DoTheWork
End Sub

Private Sub optAlignV_Click(Index As Integer)
If optAlignV(Index) Then optAlignV(0).Tag = Index
Call DoTheWork
End Sub

Private Sub DoTheWork()

Cls
Do Until Right$(Text1.Text, 2) <> vbCrLf
    Text1.Text = Left$(Text1.Text, Len(Text1.Text) - 2)
Loop

Dim hFont As Long, fontAngle As Long
' in the function calls below, I am subtracting a 2 pixel border
' from the destination rectangle so the text doesn't print on the
' rectangle edges. The beauty of the function is the ability to
' specifically place rotated text anywhere & have it clipped too.

fontAngle = Abs(CBool(chkDegrees(0) = 0)) * 180 + 90
hFont = SelectObject(Me.hdc, LoadFont(fontAngle))

' let's do the vertical fonts first
RotateText Me.hdc, fontAngle, _
    Val(optAlignment(0).Tag), Val(optAlignV(0).Tag), _
    Text1.Text, Shape1.Left + 2, Shape1.Top + 2, _
    Shape1.Width - 4 + Shape1.Left, Shape1.Height - 4 + Shape1.Top

DeleteObject SelectObject(Me.hdc, hFont)
fontAngle = 0
hFont = SelectObject(Me.hdc, LoadFont(fontAngle))

' now we'll do the horizontal font
RotateText Me.hdc, fontAngle, _
    Val(optAlignment(0).Tag), Val(optAlignV(0).Tag), _
    Text1.Text, Shape2.Left + 2, Shape2.Top + 2, _
    Shape2.Width - 4 + Shape2.Left, Shape2.Height - 4 + Shape2.Top

DeleteObject SelectObject(Me.hdc, hFont)

Refresh

End Sub
