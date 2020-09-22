VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brain Twister!"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   8760
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   0
      Picture         =   "frmMain.frx":3282
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
      Begin VB.Timer Timer3 
         Interval        =   500
         Left            =   7680
         Top             =   6240
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   8280
         Top             =   6240
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   8880
         Top             =   6240
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   0
      Width           =   9600
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   0
      Picture         =   "frmMain.frx":4471B
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Begin VB.Menu mnuShowDebug 
         Caption         =   "Show &Debug data"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Add Solution"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove Solution"
      End
      Begin VB.Menu mnuUpdateSol 
         Caption         =   "&Update Solution"
      End
   End
   Begin VB.Menu mnuPuzzle 
      Caption         =   "&Puzzle"
      Begin VB.Menu mnuBack 
         Caption         =   "&Back"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim deltay As Single
Dim deltax As Single
Dim deltaa As Single

Dim gotPiece As Boolean
Dim gamestart As Boolean

Dim pieces As Integer
Dim lastpiece As Integer


Dim Twister() As PuzzlePiece

Dim testsol As String
Dim solution(100) As String
Dim curpuz As Integer
Dim lastpuz As Integer
Dim puzzlepack(40) As Puzzle
Dim dragscreen As Boolean

Dim bViewVertexNumbers As Boolean
Dim bViewVertexes As Boolean
Dim bShowDebug As Boolean
Dim hCursors(12) As Long
Dim zList() As Integer

Private Type BOUNDSBOX
    X1 As Single
    Y1 As Single
    X2 As Single
    Y2 As Single
    X3 As Single
    Y3 As Single
    X4 As Single
    Y4 As Single
End Type

Dim PuzzleTimer As Long
Dim StartTimer As Long

Dim NormalBrush As Long
Dim FlippedBrush As Long
Dim SelectBrush As Long

Dim curBounds As BOUNDSBOX
Dim snap As Single

Dim nda(3) As Integer
Dim ndx(3) As Integer
Dim ndy(3) As Integer

'Function square(ByVal value As Single) As Single
'    square = value * value
'End Function

' converts plotted points into points relative to center of the puzzle piece.
' this allows us to be able to rotate the points easily
Private Sub GenerateInfo(pinfo() As PuzzlePoint, points() As POINTAPI, center As POINTAPI)
    For i = LBound(pinfo) To UBound(pinfo)
        With points(i)
            If .Y = center.Y Then
                ' set to 90 degrees
                pinfo(i).angle = 0.5 * 3.141592654
            Else
                ' get the angle
                pinfo(i).angle = Atn((.X - center.X) / (.Y - center.Y))
            End If

            If pinfo(i).angle < 0 Then pinfo(i).angle = 6.283185308 + pinfo(i).angle

            pinfo(i).dist = Sqr((.X - center.X) ^ 2 + (.Y - center.Y) ^ 2)
            If (.Y - center.Y) > 0 Then pinfo(i).dist = -pinfo(i).dist
        End With
    Next
End Sub

' reads puzzle piece data from a text file
Private Sub ReadPiecesTXT()
    Dim ender As String
    Dim n As Integer
    Open App.Path & "\puz.txt" For Input As 1
    Input #1, ender
    If ender <> "PUZ" Then Stop
    Input #1, pieces

    ReDim Twister(pieces - 1)
    ReDim zList(pieces - 1)

    For n = 0 To pieces - 1
        With Twister(n)
            .angle = 0
            Input #1, .center.X
            Input #1, .center.Y
            Input #1, .pointcount

            ReDim .points(.pointcount - 1)
            ReDim .pinfo(.pointcount - 1)

            For i = 0 To .pointcount - 1
                With .points(i)
                    Input #1, .X
                    Input #1, .Y
                End With
            Next

            GenerateInfo .pinfo, .points, .center

            .flipped = 1

            Randomize Timer
            .centerx = .center.X
            .centery = .center.Y

            '.centerx = Int(Rnd * 150) - 300 + (Picture2.Width / Screen.TwipsPerPixelX) / 2
            '.centery = Int(Rnd * 150) - 300 + (Picture2.Height / Screen.TwipsPerPixelY) / 2
            '.centerx = Int(10 * Rnd - 5) * 50 + (Picture2.Width / Screen.TwipsPerPixelX) / 2
            '.centery = Int(6 * Rnd - 3) * 50 + (Picture2.Height / Screen.TwipsPerPixelY) / 2
            zList(n) = n
        End With

        DrawPuzzlePiece n, 0, 0, 0
    Next

    Close 1
End Sub


' reads puzzle piece data from a binary file
Private Sub ReadPiecesBIN()
    Dim str As String * 4
    Dim n As Integer
    Open App.Path & "\pieces.dat" For Binary As 1
    Get #1, , str
    If str <> "PUZ " Then Stop
    Get #1, , pieces

    ReDim Twister(pieces - 1)
    ReDim zList(pieces - 1)

    

    For n = 0 To pieces - 1
        With Twister(n)
            .angle = 0
            Get #1, , .center.X
            Get #1, , .center.Y
            Get #1, , .pointcount

            ReDim .points(.pointcount - 1)
            ReDim .pinfo(.pointcount - 1)

            For i = 0 To .pointcount - 1
                With .points(i)
                    Get #1, , .X
                    Get #1, , .Y
                End With
            Next

            GenerateInfo .pinfo, .points, .center

            .flipped = 1
            Randomize Timer

            '.centerx = Int(Rnd * 150) - 300 + (Picture2.Width / Screen.TwipsPerPixelX) / 2
            '.centery = Int(Rnd * 150) - 300 + (Picture2.Height / Screen.TwipsPerPixelY) / 2
            .centerx = 100 + n * 150
            .centery = (Picture2.Height / Screen.TwipsPerPixelY) / 2
            zList(n) = n
        End With

        DrawPuzzlePiece n, 0, 0, 0
    Next

    Close 1
End Sub

' used to convert text puzzle piece data into binary
Private Sub WritePiecesBIN()
    Open App.Path & "\pieces.dat" For Binary As 1
    Put #1, , "PUZ "
    Put #1, , pieces

    For n = 0 To pieces - 1
        With Twister(n)
            Put #1, , .center.X
            Put #1, , .center.Y
            Put #1, , .pointcount

            For i = 0 To .pointcount - 1
                With .points(i)
                    Put #1, , .X
                    Put #1, , .Y
                End With
            Next

        End With
    Next

    Close 1
End Sub

Private Sub Form_Load()
    Dim ender As String
    Dim n As Integer


    Me.Show

    Randomize Timer

    'ReadPiecesTXT
    'WritePiecesBIN

    ReadPiecesBIN
    n = 0
    ' read puzzle picture and solution data
    Open App.Path & "\original.pak" For Input As 1
    While Not EOF(1)
        With puzzlepack(n)
            Input #1, .picture
            Input #1, .solutions
            ReDim .solution(.solutions - 1)
            For i = 0 To .solutions - 1
                Input #1, .solution(i)
            Next
        End With
        n = n + 1
    Wend
    lastpuz = n - 1
    Close 1

    ' read user data
    If Dir(App.Path & "\original.dat") <> "" Then
        n = 0
        Dim test As Integer
        Open App.Path & "\original.dat" For Input As 1
        While Not EOF(1)
            With puzzlepack(n)
                Input #1, test
                If test = 1 Then .solved = True
            End With
            n = n + 1
        Wend
        Close 1

    End If

    curpuz = 0

    ' load the first puzzle picture
    frmObjective.picture = LoadPicture(App.Path & "\puzzles\" & puzzlepack(curpuz).picture)
    frmObjective.Show , Me

    ' load our cursors
    For i = 1 To 12
        hCursors(i) = LoadResourceCursor(i)
    Next

    NormalBrush = CreateSolidBrush(RGB(255, 255, 0))
    FlippedBrush = CreateSolidBrush(RGB(220, 255, 0))
    SelectBrush = CreateSolidBrush(RGB(255, 220, 0))


    StartTimer = GetTickCount

    Picture2.FontSize = 20
    Picture2.ForeColor = RGB(255, 255, 255)

    Dim j As Integer
    Dim r As String
    
    While DoEvents
        BitBlt Picture1.hdc, 0, 0, 640, 480, Picture4.hdc, 0, 0, SRCCOPY

        ' redraw pieces
        For i = 0 To pieces - 1
            j = zList(i)
            DrawPuzzlePiece j, 0, 0, 0
        Next

        BitBlt Picture2.hdc, 0, 0, Picture2.Width / Screen.TwipsPerPixelX, Picture2.Height / Screen.TwipsPerPixelY, Picture1.hdc, 0, 0, SRCCOPY

'        PuzzleTimer = GetTickCount - StartTimer
'
'        eq = PuzzleTimer / 1000
'        r = eq \ 60 & ":" & Format(eq Mod 60, "00")
'        TextOut Picture2.hdc, 550, 10, r, Len(r)

        Picture2.Refresh
    Wend

    'BitBlt Picture2.hdc, 0, 0, Picture2.Width / Screen.TwipsPerPixelX, Picture2.Height / Screen.TwipsPerPixelY, Picture1.hdc, 0, 0, SRCCOPY
    'Picture2.Refresh

End Sub

' draws a translated and rotated puzzle piece,
' but doesn't commit the rotation and translation to the puzzle
Private Sub DrawPuzzlePiece(piece As Integer, offx As Single, offy As Single, rotation As Single)
    Dim hBrush As Long

    ' convert angle to radians
    rotation = rotation * 3.141592654 / 180

    With Twister(piece)

        For i = 0 To .pointcount - 1
            .points(i).X = offx + Int(.centerx) - Sin(.pinfo(i).angle + .angle + rotation) * .pinfo(i).dist * .flipped
            .points(i).Y = offy + Int(.centery) - Cos(.pinfo(i).angle + .angle + rotation) * .pinfo(i).dist
        Next

        ' re-create the region that will be used for hit-testing and drawing
        If .hrgn <> 0 Then DeleteObject (.hrgn)
        .hrgn = CreatePolygonRgn(.points(0), .pointcount, WINDING)

        ' select the proper brush
        If .current Then
            hBrush = SelectBrush
        Else
            If .flipped = -1 Then
                hBrush = FlippedBrush
            Else
                hBrush = NormalBrush
            End If
        End If

        ' draw the region
        FillRgn Picture1.hdc, .hrgn, hBrush

        ' debug stuff...
        If bViewVertexNumbers Or bViewVertexes Then
            For i = 0 To .pointcount - 1

                If .current And bViewVertexNumbers Then
                    TextOut Picture1.hdc, .points(i).X - 5, .points(i).Y - 5, CStr(i), Len(CStr(i))
                End If

                If bViewVertexes Then
                    Picture1.Line (.points(i).X - 2, .points(i).Y)-(.points(i).X + 2, .points(i).Y)
                    Picture1.Line (.points(i).X, .points(i).Y - 2)-(.points(i).X, .points(i).Y + 2)
                End If
            Next

            TextOut Picture1.hdc, .centerx + offx, .centery + offy, CStr(piece + 1), Len(CStr(piece + 1))

        End If

        Polygon Picture1.hdc, .points(0), .pointcount

    End With

    'Picture1.Refresh
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If Me.WindowState = vbMinimized Then Exit Sub

    Picture1.Cls

    If Me.Height > 650 Then

        Picture1.Height = Me.Height - 650
        Picture1.Width = Me.Width - 50
        Picture2.Height = Picture1.Height
        Picture2.Width = Picture1.Width

        For i = 0 To pieces - 1
            DrawPuzzlePiece i, 0, 0, 0
        Next

        BitBlt Picture2.hdc, 0, 0, Picture2.Width / Screen.TwipsPerPixelX, Picture2.Height / Screen.TwipsPerPixelY, Picture1.hdc, 0, 0, SRCCOPY
        Picture2.Refresh

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    DeleteObject SelectBrush
    DeleteObject FlippedBrush
    DeleteObject NormalBrush

    For i = 0 To pieces - 1
        With Twister(i)
            If .hrgn <> 0 Then DeleteObject (.hrgn)
            .hrgn = 0
        End With
    Next

    Open App.Path & "\original.dat" For Output As 1
    For n = 0 To lastpuz
        Print #1, IIf(puzzlepack(n).solved, "1", "0")
    Next
    Close 1

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuBack_Click()
    Dim i As Integer

    curpuz = curpuz - 1
    If curpuz < 0 Then curpuz = 0

    frmObjective.picture = LoadPicture(App.Path & "\puzzles\" & puzzlepack(curpuz).picture)

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNext_Click()
    Dim i As Integer

    curpuz = curpuz + 1
    If curpuz > lastpuz Then curpuz = lastpuz

    frmObjective.picture = LoadPicture(App.Path & "\puzzles\" & puzzlepack(curpuz).picture)

End Sub

Private Sub mnuRemove_Click()
Dim solFound As Boolean
Dim thisSol As Integer
    With puzzlepack(curpuz)
        For i = 0 To .solutions - 1
            If testsol = .solution(i) Then
                solFound = True
                thisSol = i
                Exit For
            End If
        Next
    
        If solFound Then
            If MsgBox("Remove this solution?", vbQuestion + vbYesNo + vbDefaultButton2, "Brain Twister") = vbYes Then
                
                For i = thisSol To .solutions - 1
                    .solution(i) = .solution(i + 1)
                Next
                
                .solutions = .solutions - 1
                ReDim Preserve .solution(.solutions - 1)
            
                MsgBox "Solution removed.", vbInformation
            End If
        Else
            MsgBox "Solution not found!", vbInformation
        End If
    
    End With
End Sub

Private Sub mnuReset_Click()
    If MsgBox("This will set all the puzzle states to ""unsolved"". Continue?", vbYesNo + vbQuestion, "Brain Twister") = vbYes Then
        For n = 0 To lastpuz
            puzzlepack(n).solved = False
        Next
    End If
End Sub

Private Sub mnuSave_Click()
Dim solFound As Boolean
    With puzzlepack(curpuz)
        For i = 0 To .solutions - 1
            If testsol = .solution(i) Then solFound = True
        Next
    
        If solFound Then
            MsgBox "Solution exists!", vbInformation
        Else
            MsgBox "Solution added!", vbInformation
            .solutions = .solutions + 1
            ReDim Preserve .solution(.solutions - 1)
            .solution(.solutions - 1) = testsol
        End If
    End With
    
End Sub

Private Sub mnuUpdateSol_Click()
Dim j As Integer, k As Integer
Dim outfile As Long
Dim outstr As String

    outfile = FreeFile
    Open App.Path & "\original.pak" For Output As outfile
    For j = 0 To lastpuz
        With puzzlepack(j)
            outstr = .picture & "," & .solutions
            For k = 0 To .solutions - 1
                outstr = outstr & "," & .solution(k)
            Next
            Print #outfile, outstr
        End With
    Next
    Close outfile

    MsgBox "Solutions updated!", vbInformation

End Sub

Private Sub mnuShowDebug_Click()
    Dim i As Integer

    bShowDebug = Not bShowDebug
    bViewVertexes = Not bViewVertexes
    bViewVertexNumbers = Not bViewVertexNumbers
    mnuShowDebug.Checked = Not mnuShowDebug.Checked
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeySpace
        dragscreen = True
    Case vbKeyPageDown
        mnuNext_Click
    Case vbKeyPageUp
        mnuBack_Click
    Case vbKeyC
        If Shift = vbCtrlMask Then mnuSave_Click
    End Select
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then dragscreen = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    ' get the reference mouse position
    ' will be used as our reference for determining difference in position when moving pieces
    deltax = X
    deltay = Y

    ' Switch from intro screen to game mode on first click
    If Not gamestart Then gamestart = True

    ' gotpiece is a boolean to check if a piece is selected
    gotPiece = False
    Twister(lastpiece).selected = False

    ' iterate through all the puzzle pieces
    For i = 0 To pieces - 1
        j = i
        'j = zList(i)
        ' hit test our mouse coords within each piece
        If PtInRegion(Twister(j).hrgn, deltax, deltay) Then
            gotPiece = True
'
'            zList(j) = zList(0)
'            zList(0) = j

            Select Case Button
            Case 1
                ' Left MB was clicked, so select the piece
                Twister(j).selected = True
                Twister(j).current = True
                ' deselect the last piece
                If j <> lastpiece Then Twister(lastpiece).current = False
                ' make this piece the last piece
                lastpiece = j
            Case 2
                ' Right MB was clicked, so flip the piece
                Twister(j).flipped = -Twister(j).flipped

            End Select
            Exit For
        End If
    Next

    For i = 0 To pieces - 1
        With Twister(i)
            .offsety = .centery
            .offsetx = .centerx
            .offseta = .angle
        End With
    Next

    If Not gotPiece And Twister(lastpiece).current Then
        If X > Twister(lastpiece).centerx And Y > Twister(lastpiece).centery Then
            SetCursorType CUR_ROT_BOTTOMRIGHT
        End If

        If X < Twister(lastpiece).centerx And Y < Twister(lastpiece).centery Then
            SetCursorType CUR_ROT_TOPLEFT
        End If

        If X < Twister(lastpiece).centerx And Y > Twister(lastpiece).centery Then
            SetCursorType CUR_ROT_BOTTOMLEFT
        End If

        If X > Twister(lastpiece).centerx And Y < Twister(lastpiece).centery Then
            SetCursorType CUR_ROT_TOPRIGHT
        End If
    Else
        SetCursorType CUR_HAND_OPEN
    End If

    ' get the bounds of the current piece

'    minx = 9000
'    miny = 9000
'    maxx = -9000
'    maxy = -9000
'
'    For i = 0 To pieces - 1
'        With Twister(i)
'            If .current Then
'                For j = 0 To .pointcount - 1
'                    If .points(j).X > maxx Then maxx = .points(j).X
'                    If .points(j).Y > maxy Then maxy = .points(j).Y
'                    If .points(j).X < minx Then minx = .points(j).X
'                    If .points(j).Y < miny Then miny = .points(j).Y
'                Next
'                Exit For
'            End If
'        End With
'    Next
'
'    With curBounds
'        .X1 = minx - 10
'        .Y1 = miny - 10
'
'        .X2 = maxx + 10
'        .Y2 = miny - 10
'
'        .X3 = maxx + 10
'        .Y3 = maxy + 10
'
'        .X4 = minx - 10
'        .Y4 = maxy + 10
'    End With

    ' get the reference angle. set to 90 degrees (in radians) if deltay = Twister(lastpiece).centery
    ' or else it will result in a divide by zero overflow
    If deltay = Twister(lastpiece).centery Then
        deltaa = 1.5707963267949
    Else
        deltaa = Atn((deltax - Twister(lastpiece).centerx) / (deltay - Twister(lastpiece).centery))
    End If

    ' adjust the rotation angle if we move above the center of the piece
    ' this has to do with angles and the result of arctangent
    If deltay < Twister(lastpiece).centery Then deltaa = deltaa + 3.141592654

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Single

    Dim i As Integer, j As Integer

    ' Clicking the Left mouse button...
    If Button = 1 Then

        ' Should we drag the objects around?
        If dragscreen Then

            For i = 0 To pieces - 1
                With Twister(i)
                    .centerx = .offsetx + (X - deltax)
                    If .centerx < 0 Then .centerx = 0
                    If .centerx > 640 Then .centerx = 640
                    
                    .centery = .offsety + (Y - deltay)
                    If .centery < 0 Then .centery = 0
                    If .centery > 480 Then .centery = 480
                    
                    .selected = False
                End With
            Next

            Exit Sub
        End If

        With Twister(lastpiece)
            If Not gotPiece And .current Then
                If X > .centerx And Y > .centery Then
                    SetCursorType CUR_ROT_BOTTOMRIGHT
                End If

                If X < .centerx And Y < .centery Then
                    SetCursorType CUR_ROT_TOPLEFT
                End If

                If X < .centerx And Y > .centery Then
                    SetCursorType CUR_ROT_BOTTOMLEFT
                End If

                If X > .centerx And Y < .centery Then
                    SetCursorType CUR_ROT_TOPRIGHT
                End If
            Else
                SetCursorType CUR_HAND_OPEN
            End If


            ' same as in Picture1_MouseDown
            ' but get the CURRENT rotation angle

            If Y = .centery Then
                a = 0.5 * 3.141592654
            Else
                a = Atn((X - .centerx) / (Y - .centery))
            End If
            If Y < .centery Then a = a + 3.141592654
        End With

        ' limit translation coords to window size
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0

        If X > Picture2.Width / Screen.TwipsPerPixelX Then X = Picture2.Width / Screen.TwipsPerPixelX
        If Y > Picture2.Height / Screen.TwipsPerPixelY Then Y = Picture2.Height / Screen.TwipsPerPixelY

        ' Redraw background
        BitBlt Picture1.hdc, 0, 0, 640, 480, Picture4.hdc, 0, 0, SRCCOPY

        If Shift = vbShiftMask Then snap = 45 Else snap = 5

        If dragscreen Then
            For i = 0 To pieces - 1
                With Twister(i)
                .centerx = .centerx + (X - deltax)
                .centery = .centery + (Y - deltay)
                .selected = False
                End With
            Next
            ' user was dragging the pieces, so don't perform an alignment check
        Else

            For i = 0 To pieces - 1
                With Twister(i)
                If .current Then
                    If .selected Then
                        ' translation occured
                        .centerx = .offsetx + (X - deltax)
                        .centery = .offsety + (Y - deltay)
                    Else
                        ' rotation occured
                        If Y = Twister(lastpiece).centery Then
                            a = 0.5 * 3.141592654
                        Else
                            a = Atn((X - Twister(lastpiece).centerx) / (Y - Twister(lastpiece).centery))
                        End If
                        If Y < Twister(lastpiece).centery Then a = a + 3.141592654
                        If Shift = vbShiftMask Then snap = 45 Else snap = 5
                        .angle = .offseta + Int((a - deltaa) * 180 / 3.1415927 / snap) * snap * Twister(i).flipped / 180 * 3.141592654
                    
                    End If
                End If
                End With
            Next
        
        End If
    Else

        If Not gamestart Then Exit Sub

        ' iterate through all the puzzle pieces
        Dim bHit As Boolean

        For i = 0 To pieces - 1
            j = zList(i)
            ' hit test our mouse coords within each piece
            If PtInRegion(Twister(j).hrgn, X, Y) Then
                If Twister(j).current Then
                    SetCursorType CUR_HAND_OPEN
                Else
                    SetCursorType CUR_HAND_POINT
                End If
                bHit = True
            End If
        Next

        If Not bHit Then
            If gotPiece Then
                If X > Twister(lastpiece).centerx And Y > Twister(lastpiece).centery Then
                    SetCursorType CUR_ROT_BOTTOMRIGHT
                End If

                If X < Twister(lastpiece).centerx And Y < Twister(lastpiece).centery Then
                    SetCursorType CUR_ROT_TOPLEFT
                End If

                If X < Twister(lastpiece).centerx And Y > Twister(lastpiece).centery Then
                    SetCursorType CUR_ROT_BOTTOMLEFT
                End If

                If X > Twister(lastpiece).centerx And Y < Twister(lastpiece).centery Then
                    SetCursorType CUR_ROT_TOPRIGHT
                End If
            End If
        End If

'        minx = 9000
'        miny = 9000
'        maxx = -9000
'        maxy = -9000
'
'        For i = 0 To pieces - 1
'            j = zList(i)
'
'            With Twister(j)
'                If .current Then
'                    For k = 0 To .pointcount - 1
'                        If .points(k).X > maxx Then maxx = .points(k).X
'                        If .points(k).Y > maxy Then maxy = .points(k).Y
'                        If .points(k).X < minx Then minx = .points(k).X
'                        If .points(k).Y < miny Then miny = .points(k).Y
'                    Next
'
'                    Picture2.DrawStyle = vbDot
'                    Picture2.Line (minx - 10, miny - 10)-(maxx + 10, maxy + 10), RGB(255, 0, 0), B
'                    Picture2.Refresh
'
'                    Exit For
'                End If
'            End With
'        Next
'

        If puzzlepack(curpuz).solved Then
            frmObjective.FontSize = 8
            frmObjective.FontBold = True
            r = "You have already solved this puzzle!"
            TextOut frmObjective.hdc, frmObjective.Width / Screen.TwipsPerPixelX / 2 - frmObjective.TextWidth(r) / Screen.TwipsPerPixelX / 2, 210, r, Len(r)

            frmObjective.Refresh
        End If

    End If


End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, j As Integer
    Dim match As Integer

    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    If X > Picture2.Width / Screen.TwipsPerPixelX Then X = Picture2.Width / Screen.TwipsPerPixelX
    If Y > Picture2.Height / Screen.TwipsPerPixelY Then Y = Picture2.Height / Screen.TwipsPerPixelY

    If Not gotPiece Then
        If X > Twister(lastpiece).centerx And Y > Twister(lastpiece).centery Then
            SetCursorType CUR_ROT_BOTTOMRIGHT
        End If

        If X < Twister(lastpiece).centerx And Y < Twister(lastpiece).centery Then
            SetCursorType CUR_ROT_TOPLEFT
        End If

        If X < Twister(lastpiece).centerx And Y > Twister(lastpiece).centery Then
            SetCursorType CUR_ROT_BOTTOMLEFT
        End If

        If X > Twister(lastpiece).centerx And Y < Twister(lastpiece).centery Then
            SetCursorType CUR_ROT_TOPRIGHT
        End If
    Else
        SetCursorType CUR_HAND_OPEN
    End If

    'Picture1.Cls
    BitBlt Picture1.hdc, 0, 0, 640, 480, Picture4.hdc, 0, 0, SRCCOPY

    If (Y = deltay And X = deltax) And gotPiece = False Then
        ' user didn't click anything, didn't rotate, so deselect
        For i = 0 To pieces - 1
            With Twister(i)
                .current = False
                .selected = False
            End With
        Next

        Exit Sub
    End If

    If Button = 1 Then

        testsol = ""

        For i = 0 To pieces - 1
            With Twister(i)
                For j = 0 To .pointcount - 1
                    For k = i To pieces - 1
                        If i <> k Then
                            For l = 0 To Twister(k).pointcount - 1
                                If Abs(Twister(k).points(l).X - .points(j).X) < 8 And _
                                   Abs(Twister(k).points(l).Y - .points(j).Y) < 8 Then

                                    If bShowDebug Then

                                        Picture1.ForeColor = RGB(0, 255, 0)
                                        Picture1.Circle (.points(j).X, .points(j).Y), 8
                                        Picture1.ForeColor = RGB(0, 0, 0)

                                        Dim r As String
                                        r = i & "(" & j & ")-" & k & "(" & l & ")"
                                        TextOut Picture1.hdc, 0, match * 14, r, Len(r)
                                    End If
                                    match = match + 1
                                    'testsol = testsol & i & "." & j & "-" & k & "." & l & ","
                                    testsol = testsol & i & j & k & l
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                Next
            End With
            'DrawPuzzlePiece i, 0, 0, 0
        Next

        Dim lastd As Double
        lastd = 8
        For i = 0 To pieces - 1
            With Twister(i)
                ldx = 0
                ldy = 0
                For j = 0 To .pointcount - 1
                    For k = 0 To pieces - 1
                        If i <> k Then
                            For l = 0 To Twister(k).pointcount - 1
                                dx = Twister(k).points(l).X - .points(j).X
                                dy = Twister(k).points(l).Y - .points(j).Y
                                If Abs(dx) < 8 And Abs(dy) < 8 Then

                                    d = Sqr(dx ^ 2 + dy ^ 2)

                                    If .current And (d < lastd) Then
                                        ldx = dx
                                        ldy = dy
                                        lastd = Sqr(dx ^ 2 + dy ^ 2)
                                    End If

                                    'Exit For
                                End If
                            Next
                        End If
                    Next
                Next
                .centerx = .centerx + ldx
                .centery = .centery + ldy
            End With
        Next

        deltax = X
        deltay = Y


        For i = 0 To puzzlepack(curpuz).solutions - 1

            'This is what we'll use in the final version, so the msgbox doesn't keep popping up
            'If testsol = puzzlepack(curpuz).solution(i) And Not puzzlepack(curpuz).solved Then

            'for now, use this so we can test all possible combinations
            If testsol = puzzlepack(curpuz).solution(i) And Not puzzlepack(curpuz).solved Then

                MsgBox "Congratulations! You solved this puzzle!", vbInformation

                puzzlepack(curpuz).solved = True

                'curpuz = curpuz + 1
                'If curpuz > lastpuz Then curpuz = lastpuz
                'Picture3.picture = LoadPicture(App.Path & "\puzzles\" & puzzlepack(curpuz).picture)
                Exit For
            End If
        Next
    End If

End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
    Picture1_KeyDown KeyCode, Shift
End Sub

Private Sub Picture2_KeyUp(KeyCode As Integer, Shift As Integer)
    Picture1_KeyUp KeyCode, Shift
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1_MouseDown Button, Shift, X, Y
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1_MouseUp Button, Shift, X, Y
End Sub

Private Sub SetCursorType(cursor As CURSOR_TYPES)
    If hCursors(cursor) > 0 Then SetCursor hCursors(cursor)
End Sub

' Load a cursor from the resource
Function LoadResourceCursor(ByVal nid As Variant) As Long
    Dim hRes As Long
    Dim dwSize As Long
    Dim hGlobal As Long
    Dim lpBytes As Long
    Dim pBytes() As Byte
    Dim hMod As Long

    Dim sID As String
    sID = "#" & CStr(nid)
    hMod = App.hInstance
    hRes = FindResource(hMod, ByVal sID, ByVal RT_CURSOR)
    If hRes <> 0 Then
        dwSize = SizeofResource(hMod, hRes)
        hGlob = LoadResource(hMod, hRes)
        lpBytes = LockResource(hGlob)
        ReDim pBytes(dwSize - 1)
        CopyMemory pBytes(0), ByVal lpBytes, dwSize
        LoadResourceCursor = CreateIconFromResource(pBytes(0), dwSize, False, &H30000)
    Else
        LoadResourceCursor = 0
    End If

End Function

Private Sub Timer1_Timer()
    Dim n As Integer

    ' Redraw background
    BitBlt Picture1.hdc, 0, 0, 640, 480, Picture4.hdc, 0, 0, SRCCOPY

    For n = 0 To pieces - 1
        With Twister(n)
            .centerx = .centerx + ndx(n)
            If .centerx > 640 Or .centerx < 0 Then
                ndx(n) = -ndx(n)
                nda(n) = Int(Rnd * 10) - 5
            End If

            .centery = .centery + ndy(n)
            If .centery > 480 Or .centery < 0 Then
                ndy(n) = -ndy(n)
                nda(n) = Int(Rnd * 10) - 5
            End If
            .angle = .angle + 3.141592654 * Int(180 * nda(n) / 5) * 5
        End With
    Next

    For n = 0 To pieces - 1
        DrawPuzzlePiece n, 0, 0, 0
    Next

    BitBlt Picture2.hdc, 0, 0, Picture2.Width / Screen.TwipsPerPixelX, Picture2.Height / Screen.TwipsPerPixelY, Picture1.hdc, 0, 0, SRCCOPY
    Picture2.Refresh

End Sub

Private Sub Timer2_Timer()
    Timer1.Enabled = True
    Timer2.Enabled = False
    Randomize Timer
    For n = 0 To 3
        ndx(n) = Int(Rnd * 4) - 2
        ndy(n) = Int(Rnd * 4) - 2
        nda(n) = Int(Rnd * 2) - 1
    Next

End Sub

