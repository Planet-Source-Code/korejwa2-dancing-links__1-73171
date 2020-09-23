VERSION 5.00
Begin VB.Form fDLinks 
   Caption         =   "Dancing Links"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   10680
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   25
      Top             =   3960
      Width           =   735
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Binary Matrix"
      Height          =   6855
      Index           =   3
      Left            =   3600
      TabIndex        =   20
      Top             =   960
      Width           =   4215
      Begin VB.TextBox txtMatrixBinary 
         Height          =   3495
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton cmdMatrixBinaryRnd 
         Caption         =   "Random"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdMatrixBinaryDLX 
         Caption         =   "Search"
         Height          =   375
         Left            =   2760
         TabIndex        =   21
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Index           =   3
         Left            =   240
         TabIndex        =   24
         Top             =   4200
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sudoku"
      Height          =   6855
      Index           =   2
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Width           =   4215
      Begin VB.CommandButton cmdSudokuDLX4x4 
         Caption         =   "4x4"
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   6120
         Width           =   615
      End
      Begin VB.TextBox txtSudoku 
         Height          =   3495
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton cmdSudokuDLX 
         Caption         =   "Search"
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSudokuClr 
         Caption         =   "Default"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   4200
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "N Queens"
      Height          =   6855
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox cboNQueens 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton cmdNQueensDLX 
         Caption         =   "Search"
         Height          =   371
         Left            =   2760
         TabIndex        =   10
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label lblChess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Q"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1560
         Width           =   255
      End
      Begin VB.Shape shpChess 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   240
         Shape           =   1  'Square
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "N="
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   6150
         Width           =   375
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   4200
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pentominoes"
      Height          =   6855
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   1
         TabIndex        =   26
         Top             =   6240
         Value           =   1
         Width           =   1695
      End
      Begin VB.CommandButton cmdPentomino 
         Caption         =   "Search"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Shape shpPentominoDemo 
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblPentSpeed 
         Caption         =   "Label2"
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   6240
         Width           =   495
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   4200
         Width           =   3735
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Binary"
      Height          =   375
      Index           =   3
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sudoku"
      Height          =   375
      Index           =   2
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   10680
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Queens"
      Height          =   375
      Index           =   1
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pent"
      Height          =   375
      Index           =   0
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3360
      Width           =   735
   End
End
Attribute VB_Name = "fDLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 0

'FileName:     DLinks.frm
'Author:       John Korejwa <korejwa.net>
'Date:         18 / November / 2010
'Description:  DLX is a depth-first-search backtracking algorithm that finds all solutions to the
'              exact cover problem.  Solve sudoku and N queens puzzles.  See Knuth's Dancing Links
'              paper for algorithm details.
'
'              http://en.wikipedia.org/wiki/Dancing_Links
'              http://www-cs-faculty.stanford.edu/~uno/papers/dancing-color.ps.gz

Private Type RECT
    Left       As Long
    top        As Long
    Right      As Long
    Bottom     As Long
End Type

Private Const WM_VSCROLL As Long = &H115
Private Const SB_BOTTOM  As Long = 7

Private m_ReserveSpace As RECT

'The DLX pentomino search is broken up into multiple procedures, allowing the search to be suspended
'and then resumed with the current state saved in the following global variables.  In most real cases
'this is not necessary or desired.  The other 3 DLX searches are pretty much contained in one function.
'With the exception of "visit exact cover solution", the recursive part of the algorithm is exactly the
'same in all 3 cases, and can be cut-and-paste for any DLX search.  If only one solution is needed, you
'can simply "Exit Do" after visiting the first solution found.  Column header construction can also be
'cut-and-paste usually.

Private m_Running As Long  'non-zero when pentomino search is active
Private m_Delay   As Long  'delay between pentomino graphic moves
Private m_psize   As Long  'size of pentomino graphic data object
Private l_b       As Long  'solution count
Private l_k       As Long  'depth in search tree
Private l_p       As Long  'current row
Private l_q       As Long  'current column
Private l_l()     As Long  'data/column object left
Private l_r()     As Long  'data/column object right
Private l_u()     As Long  'data/column object up
Private l_d()     As Long  'data/column object down
Private l_c()     As Long  'data object column id / column object size
Private l_o()     As Long  'solution vector

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long 'The left and top members are zero. The right and bottom members contain the width and height of the window.
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'add Item to Text1 and scroll to bottom; trim Text1.Text if necessary to comply with TextBox assign size limitation of 65535 chars
Private Sub AddItemText(ByRef Item As String)
    Dim r As Long, s As Long, t As Long

    With Text1
        s = Len(.Text)
        t = Len(Item)
        If (t > 65533) Then
            r = InStr(t - 65533, Item, vbCrLf)
            If (r = 0) Then
                r = 65533
            Else
                r = t - (r + 1)
            End If
            .Text = Right$(Item, r) & vbCrLf
        ElseIf (s + t > 65533) Then
            r = InStr(s + t - 65533, .Text, vbCrLf)
            If (r = 0) Then
                r = 65533 - t
            Else
                r = s - (r + 1)
            End If
            .Text = Right$(.Text, r) & Item & vbCrLf
        Else
            .Text = .Text & Item & vbCrLf
        End If
        SendMessage .hwnd, WM_VSCROLL, SB_BOTTOM, 0
    End With
End Sub

'returns a multiline String containing matrix a(0...m-1, 0...n-1)
'or, if n=0, column vector a(0...m-1)
'or, if m=0, row    vector a(0..n-1)
Public Function MatrixFormat(ByRef a As Variant, ByVal m As Long, ByVal n As Long, ByVal fmt As String) As String
    Dim h As Long, i As Long, j As Long, k As Long
    Dim s As String, t As String ''"$#,##0;;\Z\e\r\o"  '"+0.00;0.00;0"

    If ((m > 0) And (n > 0)) Then 'matrix a(0...m-1, 0...n-1)
        h = 0 'determine maximum string length h
        For i = 0 To m - 1
            For j = 0 To n - 1
                s = Format$(a(i, j), fmt): If InStrRev(s, ".") = Len(s) Then s = Left$(s, Len(s) - 1)
                k = Len(s)
                If (h < k) Then h = k
            Next j
        Next i
        h = h + 1
        For i = 0 To m - 1
            For j = 0 To n - 1
                s = Format$(a(i, j), fmt): If InStrRev(s, ".") = Len(s) Then s = Left$(s, Len(s) - 1)
                k = Len(s)
                t = t & Space$(h - k) & s
            Next j
            t = t & vbCrLf
        Next i
    ElseIf ((m = 0) And (n > 0)) Then 'row vector a(0...n-1)
        h = 0
        For j = 0 To n - 1
            s = Format$(a(j), fmt)
            k = Len(s)
            If (h < k) Then h = k
        Next j
        h = h + 1
        For j = 0 To n - 1
            s = Format$(a(j), fmt)
            k = Len(s)
            t = t & Space$(h - k) & s
        Next j
        t = t & vbCrLf
    ElseIf ((m > 0) And (n = 0)) Then 'column vector a(0...m-1)
        h = 0
        For i = 0 To m - 1
            k = Len(Format$(a(i), fmt))
            If (h < k) Then h = k
        Next i
        h = h + 1
        For i = 0 To m - 1
            s = Format$(a(i), fmt)
            k = Len(s)
            t = t & Space$(h - k) & s & vbCrLf
        Next i
    End If
    MatrixFormat = t
End Function

'extract a numerical array from a text string.
'multiline string results in 2 dimensional array a(0..m-1, 0..n-1)
'otherwise a one dimensional array is returned.  a(0...n-1)
'returns number of dimensions (0, 1, or 2) or -1 on error
Public Function MatrixFormatInv(s As String, ByRef a As Variant, ByRef m As Long, ByRef n As Long) As Long
    Dim c As Long, i As Long, j As Long, k As Long, p As Long
    Dim g() As Long, h() As Long
On Error GoTo ErrHandler

    ReDim g(15)      'start of numerical string
    ReDim h(15)      'length of numerical string
    i = 1            'char index
    k = 0            'flag - "used decimal point seen in current number"
    m = 0            'array height
    n = 0            'array width
    p = Len(s)       'length of string
    Do While (i <= p)
        c = Asc(Mid$(s, i, 1)) 'current char
        i = i + 1
        Select Case c 'look for #, -#, .#, -.#
        Case 48 To 57 '"#"   (0...9)
            g(n) = i - 1
            k = 0
        Case 46       '"."
            If (i <= p) Then
                c = Asc(Mid$(s, i, 1))
                i = i + 1
                If ((c >= 48) And (c <= 57)) Then '".#"
                    k = 1
                    g(n) = i - 2
                End If
            End If
        Case 45 '"-"
            If (i <= p) Then
                c = Asc(Mid$(s, i, 1))
                i = i + 1
                Select Case c
                Case 48 To 57 '"-#"
                    g(n) = i - 2
                    k = 0
                Case 46 '"-."
                    If (i <= p) Then
                        c = Asc(Mid$(s, i, 1))
                        i = i + 1
                        If ((c >= 48) And (c <= 57)) Then '"-.#"
                            k = 1
                            g(n) = i - 3
                        End If
                    End If
                End Select
            End If
        Case 10, 13 'crlf
            If ((m = 0) And (i <= p)) Then m = n
        End Select
        If (g(n) <> 0) Then
            Do
                If (i <= p) Then
                    c = Asc(Mid$(s, i, 1))
                Else
                    c = 0 'exit loop
                End If
                i = i + 1
                If (c = 46) Then '"."
                    If (k = 0) Then
                        k = 1
                        c = 48 'dummy to stay in loop
                    End If
                End If
            Loop While ((c >= 48) And (c <= 57))
            h(n) = i - 1 - g(n) 'length of numerical string, excluding possible exponent
            Select Case c
            Case 68, 69, 100, 101 'exponent char {e,d,E,D}
                If (i <= p) Then
                    c = Asc(Mid$(s, i, 1))
                    i = i + 1
                    If ((i <= p) And ((c = 45) Or (c = 43))) Then '"-" "+"
                        c = Asc(Mid$(s, i, 1))
                        i = i + 1
                    End If
                    If ((c >= 48) And (c <= 57)) Then 'digit follows exponent
                        Do
                            If (i <= p) Then
                                c = Asc(Mid$(s, i, 1))
                            Else
                                c = 0
                            End If
                            i = i + 1
                        Loop While ((c >= 48) And (c <= 57))
                        h(n) = i - 1 - g(n) 'length of numerical string, including exponent
                    End If
                End If
            End Select
            n = n + 1
            k = 0
            If ((n And 15) = 0) Then
                ReDim Preserve g(n + 15)
                ReDim Preserve h(n + 15)
            End If
        End If
    Loop
    If (m <> 0) Then 'multi dimensional array?
        If ((n Mod m) = 0) Then
            p = m
            m = n \ m
            n = p
        Else
            m = 0
        End If
    End If
    If (m = 0) Then
        If (n = 0) Then
            MatrixFormatInv = 0
        Else
            ReDim a(n - 1)
            For i = 0 To n - 1
                a(i) = Mid$(s, g(i), h(i))
            Next i
            MatrixFormatInv = 1
        End If
    Else
        ReDim a(m - 1, n - 1)
        k = 0
        For i = 0 To m - 1
            For j = 0 To n - 1
                a(i, j) = Mid$(s, g(k), h(k))
                k = k + 1
            Next j
        Next i
        MatrixFormatInv = 2
    End If
Exit Function
ErrHandler: 'an error usually means text number is too large for the variable type
MatrixFormatInv = -1
End Function

'returns a multiline String containing sudoku matrix a(0..n-1, 0..n-1)
Public Function SudokuMatrixFormat(a() As Long, n As Long) As String
    Dim i As Long, j As Long, m As Long, s As String

    If (n > 0) Then
        m = Sqr(n)
        If (m * m = n) Then
            For j = 1 To m
                s = s & "+" & String(m + m + 1, "-")
            Next j
            s = s & "+" & vbCrLf
            For i = 0 To n - 1
                s = s & "|"
                For j = 0 To n - 1
                    s = s & " " & Chr$(a(i, j))
                    If (j Mod m) = m - 1 Then s = s & " |"
                Next j
                s = s & vbCrLf
                If (i Mod m) = m - 1 Then
                    For j = 1 To m
                        s = s & "+" & String(m + m + 1, "-")
                    Next j
                    s = s & "+" & vbCrLf
                End If
            Next i
        End If
    End If
    SudokuMatrixFormat = s
End Function

'reads a multiline String into sudoku matrix a(0..n-1, 0..n-1)
'alphabet variables x(0..n) and size n must be given
'returns 0 on success
Private Function SudokuMatrixFormatInv(a() As Long, x() As Long, n As Long, s As String) As Long
    Dim c As Long, h As Long, i As Long, j As Long, m As Long, y() As Long

    If ((n <= 0) Or (n > 255)) Then 'n is positive integer?
        i = n
    Else
        m = Sqr(n)
        If (m * m <> n) Then        'n is a square?
            i = n
        Else
            For i = 0 To n          'elements of x(0..n) in the set {0..255} and unique?
                Select Case x(i)
                Case Is < 0:   Exit For
                Case Is > 255: Exit For
                End Select
                j = i + 1
                While (j <= n)
                    If (x(i) = x(j)) Then Exit For
                    j = j + 1
                Wend
            Next i
        End If
    End If
    If (i <= n) Then
        SudokuMatrixFormatInv = 1
        Exit Function
    End If

    ReDim y(255) 'build y inverse alpha
    For i = 1 To n
        y(x(i)) = i
    Next i
    y(x(0)) = n + 1 'wild char

    h = 1
    i = 0
    j = 0
    While (h <= Len(s))
        c = Asc(Mid$(s, h, 1))
        h = h + 1
        If (y(c) <> 0) Then
            If (j = n) Then
                SudokuMatrixFormatInv = 1
                Exit Function
            End If
            a(i, j) = c
            j = j + 1
        ElseIf (c = 10) Then
            If (j <> 0) Then
                If ((j <> n) Or (i = n)) Then
                    SudokuMatrixFormatInv = 1
                    Exit Function
                End If
                i = i + 1
                j = 0
            End If
        End If
    Wend
    If (i = n) Then
        SudokuMatrixFormatInv = 0
    Else
        If ((i = n - 1) And (j = n) And (h = Len(s) + 1)) Then
            SudokuMatrixFormatInv = 0
        Else
            SudokuMatrixFormatInv = 1
        End If
    End If
End Function

'place pentomino graphic on the chessboard
Private Sub PlacePentomino(u() As Long, d() As Long, l() As Long, r() As Long, c() As Long, ByVal k As Long)
    Dim h As Long, i As Long, j As Long

    While ((c(k) < 65) Or (c(k) > 76))
        k = r(k)
    Wend
    h = (c(k) - 65) * 5
    Do
        k = l(k)
        If (c(k) > 64) Then Exit Do
        i = (c(k) - 1) \ 8
        j = (c(k) - 1) Mod 8
        Shape1(h).Move Shape1(61).Left + m_psize * i, Shape1(61).top + m_psize * j
        h = h + 1
    Loop
End Sub

'move pentomino graphic off the chessboard
Private Sub RemovePentomino(o As Long)
    Select Case o
    Case 1  'F
        Shape1(0).Move Shape1(61).Left - 3 * m_psize, Shape1(61).top, m_psize, m_psize
        Shape1(1).Move Shape1(0).Left, Shape1(0).top + m_psize, m_psize, m_psize
        Shape1(2).Move Shape1(1).Left - m_psize, Shape1(1).top, m_psize, m_psize
        Shape1(3).Move Shape1(1).Left + m_psize, Shape1(1).top, m_psize, m_psize
        Shape1(4).Move Shape1(2).Left, Shape1(2).top + m_psize, m_psize, m_psize
    Case 2  'I
        Shape1(5).Move Shape1(61).Left + m_psize * 5, Shape1(61).top + m_psize * 12, m_psize, m_psize
        Shape1(6).Move Shape1(5).Left + m_psize, Shape1(5).top, m_psize, m_psize
        Shape1(7).Move Shape1(6).Left + m_psize, Shape1(6).top, m_psize, m_psize
        Shape1(8).Move Shape1(7).Left + m_psize, Shape1(7).top, m_psize, m_psize
        Shape1(9).Move Shape1(8).Left + m_psize, Shape1(8).top, m_psize, m_psize
    Case 3  'L
        Shape1(10).Move Shape1(61).Left, Shape1(61).top + m_psize * 11, m_psize, m_psize
        Shape1(11).Move Shape1(10).Left, Shape1(10).top + m_psize, m_psize, m_psize
        Shape1(12).Move Shape1(11).Left + m_psize, Shape1(11).top, m_psize, m_psize
        Shape1(13).Move Shape1(12).Left + m_psize, Shape1(12).top, m_psize, m_psize
        Shape1(14).Move Shape1(13).Left + m_psize, Shape1(13).top, m_psize, m_psize
    Case 4  'N
        Shape1(15).Move Shape1(61).Left, Shape1(61).top + m_psize * 9, m_psize, m_psize
        Shape1(16).Move Shape1(15).Left + m_psize, Shape1(15).top, m_psize, m_psize
        Shape1(17).Move Shape1(16).Left + m_psize, Shape1(16).top, m_psize, m_psize
        Shape1(18).Move Shape1(17).Left, Shape1(17).top + m_psize, m_psize, m_psize
        Shape1(19).Move Shape1(18).Left + m_psize, Shape1(18).top, m_psize, m_psize
    Case 5  'P
        Shape1(20).Move Shape1(61).Left - m_psize * 4, Shape1(61).top + m_psize * 7, m_psize, m_psize
        Shape1(21).Move Shape1(20).Left + m_psize, Shape1(20).top, m_psize, m_psize
        Shape1(22).Move Shape1(20).Left, Shape1(20).top + m_psize, m_psize, m_psize
        Shape1(23).Move Shape1(22).Left + m_psize, Shape1(22).top, m_psize, m_psize
        Shape1(24).Move Shape1(23).Left + m_psize, Shape1(23).top, m_psize, m_psize
    Case 6  'T
        Shape1(25).Move Shape1(61).Left - m_psize * 2, Shape1(61).top + m_psize * 3, m_psize, m_psize
        Shape1(26).Move Shape1(25).Left, Shape1(25).top + m_psize, m_psize, m_psize
        Shape1(27).Move Shape1(26).Left, Shape1(26).top + m_psize, m_psize, m_psize
        Shape1(28).Move Shape1(26).Left - m_psize, Shape1(26).top, m_psize, m_psize
        Shape1(29).Move Shape1(28).Left - m_psize, Shape1(28).top, m_psize, m_psize
    Case 7  'U
        Shape1(30).Move Shape1(61).Left + 9 * m_psize, Shape1(61).top + m_psize * 7, m_psize, m_psize
        Shape1(31).Move Shape1(30).Left, Shape1(30).top + m_psize, m_psize, m_psize
        Shape1(32).Move Shape1(31).Left + m_psize, Shape1(31).top, m_psize, m_psize
        Shape1(33).Move Shape1(32).Left + m_psize, Shape1(32).top, m_psize, m_psize
        Shape1(34).Move Shape1(33).Left, Shape1(33).top - m_psize, m_psize, m_psize
    Case 8  'V
        Shape1(35).Move Shape1(61).Left + m_psize * 9, Shape1(61).top + m_psize * 10, m_psize, m_psize
        Shape1(36).Move Shape1(35).Left + m_psize, Shape1(35).top, m_psize, m_psize
        Shape1(37).Move Shape1(36).Left + m_psize, Shape1(36).top, m_psize, m_psize
        Shape1(38).Move Shape1(37).Left, Shape1(37).top + m_psize, m_psize, m_psize
        Shape1(39).Move Shape1(38).Left, Shape1(38).top + m_psize, m_psize, m_psize
    Case 9  'W
        Shape1(40).Move Shape1(61).Left + 9 * m_psize, Shape1(61).top + m_psize * 3, m_psize, m_psize
        Shape1(41).Move Shape1(40).Left, Shape1(40).top + m_psize, m_psize, m_psize
        Shape1(42).Move Shape1(41).Left + m_psize, Shape1(41).top, m_psize, m_psize
        Shape1(43).Move Shape1(42).Left, Shape1(42).top + m_psize, m_psize, m_psize
        Shape1(44).Move Shape1(43).Left + m_psize, Shape1(43).top, m_psize, m_psize
    Case 10 'X
        Shape1(45).Move Shape1(61).Left - m_psize * 3, Shape1(61).top + m_psize * 10, m_psize, m_psize
        Shape1(46).Move Shape1(45).Left, Shape1(45).top + m_psize, m_psize, m_psize
        Shape1(47).Move Shape1(46).Left - m_psize, Shape1(46).top, m_psize, m_psize
        Shape1(48).Move Shape1(46).Left + m_psize, Shape1(46).top, m_psize, m_psize
        Shape1(49).Move Shape1(46).Left, Shape1(46).top + m_psize, m_psize, m_psize
    Case 11 'Y
        Shape1(50).Move Shape1(61).Left + m_psize * 4, Shape1(61).top + m_psize * 9, m_psize, m_psize
        Shape1(51).Move Shape1(50).Left + m_psize, Shape1(50).top, m_psize, m_psize
        Shape1(52).Move Shape1(51).Left + m_psize, Shape1(51).top, m_psize, m_psize
        Shape1(53).Move Shape1(52).Left + m_psize, Shape1(52).top, m_psize, m_psize
        Shape1(54).Move Shape1(52).Left, Shape1(52).top + m_psize, m_psize, m_psize
    Case 12 'Z
        Shape1(55).Move Shape1(61).Left + 9 * m_psize, Shape1(61).top, m_psize, m_psize
        Shape1(56).Move Shape1(55).Left, Shape1(55).top + m_psize, m_psize, m_psize
        Shape1(57).Move Shape1(56).Left + m_psize, Shape1(56).top, m_psize, m_psize
        Shape1(58).Move Shape1(57).Left + m_psize, Shape1(57).top, m_psize, m_psize
        Shape1(59).Move Shape1(58).Left, Shape1(58).top + m_psize, m_psize, m_psize
    End Select
End Sub

''remove rows from a DLX matrix which by themselves exclude the possibility of a solution.  For example,
''the "U" pentomino can not be upright and at the top of the chessboard because it isolates a square
''which can then never be covered by another pentomino.  This should be called immediately before
''the recursive part of DLX, after row construction is finished.
'Private Sub RemoveDLXDeadRows(u() As Long, d() As Long, l() As Long, r() As Long, c() As Long, ByVal kmax2 As Long)
'    Dim h As Long, i As Long, j As Long, k As Long, p As Long, q As Long, s As Long
'    Dim v() As Long
'
'    ReDim v(kmax2 - 1)
'
'    s = -1
'    q = r(0)
'    While (q <> 0)
'
'        'cover column q
'        l(r(q)) = l(q)        'remove q from the column header list
'        r(l(q)) = r(q)
'        p = d(q)              'p <- top data object in column q
'        While (p <> q)        'for each data object p in column q ...
'            j = r(p)              'j <- initial data object in p's row
'            While (j <> p)        'for all other data objects j in p's row ...
'                u(d(j)) = u(j)        'remove it from its column
'                d(u(j)) = d(j)
'                c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
'                j = r(j)          'j <- next data object in p's row
'            Wend
'            p = d(p)          'p <- next data object in column q
'        Wend
'
'        p = d(q) 'p <- top data object in column q
'        While (p <> q)
'            If (v(p) = 0) Then 'have not visited p yet
'
'                'cover the columns of all other data objects h in p's row  (left to right)
'                h = r(p)              'h <- initial data object in p's row
'                While (h <> p)        'for each data object h in p's row ...
'                    v(h) = 1
'                    'cover h's column c(h)
'                    l(r(c(h))) = l(c(h))  'remove c(h) from the column header list
'                    r(l(c(h))) = r(c(h))
'                    i = d(c(h))           'i <- top data object in column c(h)
'                    While (i <> c(h))     'for each data object i in column c(h) ...
'                        j = r(i)              'j <- initial data object in i's row
'                        While (j <> i)        'for all other data objects j in i's row ...
'                            u(d(j)) = u(j)        'remove it from its column
'                            d(u(j)) = d(j)
'                            c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
'                            j = r(j)          'j <- next data object in i's row
'                        Wend
'                        i = d(i)          'i <- next data object in column c(h)
'                    Wend
'                    h = r(h)          'h <- next data object in p's row
'                Wend
'                v(h) = 1
'
'                k = r(0)              'are any column object counts zero?
'                While (k <> 0) And (c(k) <> 0)
'                    k = r(k)
'                Wend
'                If (k <> 0) Then      'yes; add row p to linked list for removal
'                    v(p) = s
'                    s = p
'                End If
'
'                'uncover the columns of all other data objects h in p's row  (right to left)
'                h = l(p)              'h <- initial data object in p's row
'                While (h <> p)        'for each data object h in p's row ...
'                    'uncover h's column c(h)
'                    l(r(c(h))) = c(h)     'restore c(h) to the column header list
'                    r(l(c(h))) = c(h)
'                    i = u(c(h))           'i <- bottom data object in column c(h)
'                    While (i <> c(h))     'for each data object i in column c(h) ...  (in reverse order of "cover h's column c(h)")
'                        j = l(i)              'j <- initial data object in i's row
'                        While (j <> i)        'for all other data objects j in i's row ...
'                            u(d(j)) = j           'restore it to its column
'                            d(u(j)) = j
'                            c(c(j)) = c(c(j)) + 1 'increment column c(j) count
'                            j = l(j)          'j <- next data object in i's row
'                        Wend
'                        i = u(i)          'i <- next data object in column c(h)
'                    Wend
'                    h = l(h)          'h <- next data object in p's row
'                Wend
'
'            End If
'            p = d(p) 'p <- next data object in column q
'        Wend
'
'        'uncover column q
'        l(r(q)) = q           'restore q to the column header list
'        r(l(q)) = q
'        p = u(q)              'p <- bottom data object in column q
'        While (p <> q)        'for each data object p in column q ...  (in reverse order of "cover column q")
'            j = l(p)              'j <- initial data object in p's row
'            While (j <> p)        'for all other data objects j in p's row ...
'                u(d(j)) = j           'restore it to its column
'                d(u(j)) = j
'                c(c(j)) = c(c(j)) + 1 'increment column c(j) count
'                j = l(j)          'j <- next data object in p's row
'            Wend
'            p = u(p)          'p <- next data object in column q
'        Wend
'
'        q = r(q)
'    Wend
'
'    'remove rows in linked list
'    While (s <> -1)       'remove row s
'        j = r(s)              'j <- initial data object in s's row
'        While (j <> s)        'for all other data objects j in s's row ...
'            u(d(j)) = u(j)        'remove it from its column
'            d(u(j)) = d(j)
'            c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
'            j = r(j)          'j <- next data object in s's row
'        Wend
'        u(d(j)) = u(j)        'remove data object s from its column
'        d(u(j)) = d(j)
'        c(c(j)) = c(c(j)) - 1 'decrement column c(s) count
'        s = v(s)          'next s in linked list
'    Wend
'
'End Sub

'add a data object to a DLX matrix
Private Sub InsertDLXDataObject(u() As Long, d() As Long, l() As Long, r() As Long, c() As Long, ByRef k As Long, ByVal StartNewRow As Long, ByVal j As Long)
    u(k) = u(j)
    u(j) = k
    d(u(k)) = k
    d(k) = j
    c(k) = j
    c(j) = c(j) + 1
    If (StartNewRow <> 0) Then
        l(k) = k
        r(k) = k
    Else
        l(k) = k - 1
        l(r(k - 1)) = k
        r(k) = r(k - 1)
        r(k - 1) = k
    End If
    k = k + 1
End Sub

'add a data object row to a DLX matrix
Private Sub InsertDLXRow(u() As Long, d() As Long, l() As Long, r() As Long, c() As Long, ByRef k As Long, ParamArray w() As Variant)
    Dim i As Long, j As Long

    For i = 0 To UBound(w)
        j = w(i)
        u(k) = u(j)
        u(j) = k
        d(u(k)) = k
        d(k) = j
        c(k) = j
        c(j) = c(j) + 1
        If (i = 0) Then  'first data object in this row
            l(k) = k
            r(k) = k
        Else
            l(k) = k - 1
            l(r(k - 1)) = k
            r(k) = r(k - 1)
            r(k - 1) = k
        End If
        k = k + 1
    Next i
End Sub

'add data object rows to a DLX matrix representing all possible placements of an yxx shape on an mxn grid.
'u(), d(), l(), r(), c() is data/column object memory; next available memory location is k
'grid locations are DLX col object indexes 1..mxn
'v(1..m*n) is a Mask for grid locations; zero if location (i,j) is available, non-zero if not.
'w(0..z-1) are columns in the DLX matrix occupied by yxx shape when in upper left corner of grid
'func2 is a mask for excluding any of the 8 orientations (reflection, rotation)
'  bit 0: none
'  bit 1: horizontal mirror
'  bit 2: rotate 180
'  bit 3: vertical mirror
'  bit 4: rotate right
'  bit 5: transverse transpose
'  bit 6: rotate left
'  bit 7: transpose
Private Sub InsertDLXShapeInGrid(u() As Long, d() As Long, l() As Long, r() As Long, c() As Long, ByRef k As Long, m As Long, n As Long, v() As Long, w() As Long, z As Long, ByVal func2 As Long)
    Dim f As Long, g As Long, h As Long, i As Long, j As Long, p As Long, q As Long, t As Long, x As Long, y As Long
    Dim s() As Long

    x = 1
    y = 1
    If (w(0) <= m * n) Then g = 1 Else g = 0
    For i = 1 To z - 1                  'sort w(0..z-1)
        h = w(i)
        For j = i - 1 To 0 Step -1
            If (w(j) < h) Then Exit For
            w(j + 1) = w(j)
        Next j
        w(j + 1) = h
        If (h <= m * n) Then
            g = g + 1  'count occupied grid locations g
            If (x < ((h - 1) Mod n) + 1) Then x = ((h - 1) Mod n) + 1
            If (y < ((h - 1) \ n) + 1) Then y = (h - 1) \ n + 1
        End If
    Next i

    ReDim s(g * 8 - 1)
    t = 0
    For f = 0 To 7                   'cycle through all 8 possible orientations
        If ((func2 And 1) = 0) Then  'if this orientation is not excluded by mask func2 ...
            For i = 0 To t - 1           'is this orientation indistinguishable from one already seen?
                h = g * i
                For j = 0 To g - 1
                    If (s(h + j) <> w(j)) Then Exit For
                Next j
                If (j = g) Then Exit For
            Next i
            If (i = t) Then              'if not, ...
                h = g * i                '  ... add it to previously seen orientations s() ...
                For j = 0 To g - 1
                    s(h + j) = w(j)
                Next j
                t = t + 1

                h = 0                    '  ...and place in all possible locations in the grid.
                For p = 0 To m - y           'vertical   location p
                    For q = 0 To n - x       'horizontal location q
                        For i = 0 To g - 1   'does mask v() exclude placment here?
                            If (v(w(i) + h) <> 0) Then Exit For
                        Next i
                        If (i = g) Then      'no; add DLX data object row for this location
                            For i = 0 To z - 1
                                If (i = 0) Then
                                    l(k) = k
                                    r(k) = k
                                Else
                                    l(k) = k - 1
                                    l(r(k - 1)) = k
                                    r(k) = r(k - 1)
                                    r(k - 1) = k
                                End If
                                If (i < g) Then
                                    j = w(i) + h
                                Else
                                    j = w(i)
                                End If
                                u(k) = u(j)
                                u(j) = k
                                d(u(k)) = k
                                d(k) = j
                                c(k) = j
                                c(j) = c(j) + 1
                                k = k + 1
                            Next i
                        End If
                        h = h + 1
                    Next q
                    h = h + x - 1
                Next p

            End If
        End If
        func2 = func2 \ 2
        'increment shape to next orientation
        Select Case f
        Case 0, 2, 5
            For h = 0 To g - 1 'HMirror   w(0..g-1)
                i = ((w(h) - 1) \ n)
                j = ((w(h) - 1) Mod n)
                w(h) = i * n + x - j
            Next h
        Case 1, 4, 6
            For h = 0 To g - 1 'VMirror   w(0..g-1)
                i = ((w(h) - 1) \ n)
                j = ((w(h) - 1) Mod n)
                w(h) = (y - i - 1) * n + j + 1
            Next h
        Case 3, 7
            For h = 0 To g - 1 'Transpose w(0..g-1)
                i = ((w(h) - 1) \ n)
                j = ((w(h) - 1) Mod n)
                w(h) = j * n + i + 1
            Next h
            j = x
            x = y
            y = j
        End Select
        For i = 1 To g - 1     'sort w(0..g-1)
            h = w(i)
            For j = i - 1 To 0 Step -1
                If (w(j) < h) Then Exit For
                w(j + 1) = w(j)
            Next j
            w(j + 1) = h
        Next i
    Next f

End Sub

Private Sub PentominoSearchInit()
    Dim j As Long, p As Long, q As Long
    Dim v() As Long

    ReDim v(64)
    v(28) = -1 'center four squares vacant
    v(29) = -1
    v(36) = -1
    v(37) = -1

    p = 64 + 12 'primary requirements (exact)
    q = p + 1   'total   requirements (primary (exact) + secondary (at most one))

    'allocate memory for root, column header list, and data objects
    ReDim l_l(9484)     'link left
    ReDim l_r(9484)     'link right
    ReDim l_u(9484)     'link up
    ReDim l_d(9484)     'link down
    ReDim l_c(9484)     'l_c(1..q) column data objects count  |  l_c(q+1.. ) points to the data object's column header object {1..q}
    ReDim l_o(12 - 1)   'solution vector

    'build 4-way linked representation of the exact cover problem
    'primary (exact) requirements:
    '  1..n       col i in a() has non-zero
    'secondary (at most one) requirements:
    '  none
    l_l(0) = p       'l_k=0 is root
    l_r(p) = 0
    l_k = 1               'l_k={1..q}    column header list
    While (l_k <= p)      'l_k={1..p} primary (exact) requirements
        l_c(l_k) = 0        'column l_k count
        l_l(l_k) = l_k - 1
        l_r(l_k - 1) = l_k
        l_d(l_k) = l_k
        l_u(l_k) = l_k
        l_k = l_k + 1
    Wend
    While (l_k <= q)      'l_k={p+1..q} secondary (at most one) requirements
        l_c(l_k) = 0        'column l_k count
        l_l(l_k) = l_k        'left  links to itself
        l_r(l_k) = l_k        'right links to itself
        l_d(l_k) = l_k
        l_u(l_k) = l_k
        l_k = l_k + 1
    Wend
    'remove center 4 square requirements
    For j = 1 To 64
        If (v(j) <> 0) Then
            l_r(l_l(j)) = l_r(j)
            l_l(l_r(j)) = l_l(j)
        End If
    Next j
    'column header oject complete.  now add data object rows.  columns 1 to 64 are squares on the chessboard.
    'columns 65 to 76 are the 12 pentominoes.  secondary requirement 77 prevents the P pentomino from being
    'flipped when pentomino X is at location 3,3.  this way the search finds "essentially different" solutions
    'and the rest are implied by symmetries.  solution vector l_o() memory is borrowed for row add procedures
    l_o(0) = 1  'Pentomino F
    l_o(1) = 2  '**
    l_o(2) = 10 ' **
    l_o(3) = 11 ' *
    l_o(4) = 18
    l_o(5) = 65
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    l_o(0) = 1  'Pentomino I
    l_o(1) = 2  '*****
    l_o(2) = 3
    l_o(3) = 4
    l_o(4) = 5
    l_o(5) = 66
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    l_o(0) = 1  'Pentomino L
    l_o(1) = 2  '****
    l_o(2) = 3  '*
    l_o(3) = 4
    l_o(4) = 9
    l_o(5) = 67
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    l_o(0) = 1  'Pentomino N
    l_o(1) = 2  '**
    l_o(2) = 10 ' ***
    l_o(3) = 11
    l_o(4) = 12
    l_o(5) = 68
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    l_o(0) = 1  'Pentomino P
    l_o(1) = 2  '***
    l_o(2) = 3  '**
    l_o(3) = 9
    l_o(4) = 10
    l_o(5) = 69
    l_o(6) = 77
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 170
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 7, 85
    l_o(0) = 1  'Pentomino T
    l_o(1) = 2  '***
    l_o(2) = 3  ' *
    l_o(3) = 10 ' *
    l_o(4) = 18
    l_o(5) = 70
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    l_o(0) = 1  'Pentomino U
    l_o(1) = 2  '***
    l_o(2) = 3  '* *
    l_o(3) = 9
    l_o(4) = 11
    l_o(5) = 71
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    l_o(0) = 1  'Pentomino V
    l_o(1) = 2  '***
    l_o(2) = 3  '*
    l_o(3) = 9  '*
    l_o(4) = 17
    l_o(5) = 72
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    l_o(0) = 1  'Pentomino W
    l_o(1) = 2  '**
    l_o(2) = 10 ' **
    l_o(3) = 11 '  *
    l_o(4) = 19
    l_o(5) = 73
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    InsertDLXRow l_u, l_d, l_l, l_r, l_c, l_k, 3, 10, 11, 12, 19, 74 'Pentomino X   ' *
    InsertDLXRow l_u, l_d, l_l, l_r, l_c, l_k, 4, 11, 12, 13, 20, 74                '***
    InsertDLXRow l_u, l_d, l_l, l_r, l_c, l_k, 11, 18, 19, 20, 27, 74, 77           ' *
    l_o(0) = 1  'Pentomino Y
    l_o(1) = 2  '****
    l_o(2) = 3  ' *
    l_o(3) = 4
    l_o(4) = 10
    l_o(5) = 75
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0
    l_o(0) = 1  'Pentomino Z
    l_o(1) = 2  '**
    l_o(2) = 10 ' *
    l_o(3) = 18 ' **
    l_o(4) = 19
    l_o(5) = 76
    InsertDLXShapeInGrid l_u, l_d, l_l, l_r, l_c, l_k, 8, 8, v, l_o, 6, 0

    l_b = 0 'number of solutions found
    l_k = 0 'level in search tree
    l_q = 0 'current column
End Sub

Private Function PentominoSearchDestruct()
    Erase l_l
    Erase l_r
    Erase l_u
    Erase l_d
    Erase l_c
    Erase l_o
    l_b = 0
    l_k = 0
    l_p = 0
    l_q = 0
End Function

'Returns:
' -1 form unloaded
'  0 solution found
'  1 search paused
'  3 search complete
Private Function PentominoSearchFindNext() As Long
    Dim h As Long, i As Long, j As Long

    Debug.Assert (m_Running = 0)
    m_Running = 1
    Do 'initialization complete; continue with the recursive part of the algorithm
        If (l_q = 0) Then 'establish current column l_q

            'choose column object l_q
            i = &H7FFFFFFF 's heuristic:
            j = l_r(0)       'choose column with minimum number of non-zeroes to minimize branching
            While (j <> 0)
                If (i > l_c(j)) Then
                    i = l_c(j)       'l_c(j) is minimum count yet seen
                    l_q = j
                End If
                j = l_r(j)
            Wend

            If (i = 0) Then 'backtrack one level; current state can not lead to a solution because column l_q is empty
                If (l_k = 0) Then Exit Do 'search complete
                l_k = l_k - 1
                l_p = l_o(l_k)
                l_q = l_c(l_p)
            Else

                'cover column l_q
                l_l(l_r(l_q)) = l_l(l_q)        'remove l_q from the column header list
                l_r(l_l(l_q)) = l_r(l_q)
                l_p = l_d(l_q)              'l_p <- top data object in column l_q
                While (l_p <> l_q)        'for each data object l_p in column l_q ...
                    j = l_r(l_p)              'j <- initial data object in l_p's row
                    While (j <> l_p)        'for all other data objects j in l_p's row ...
                        l_u(l_d(j)) = l_u(j)        'remove it from its column
                        l_d(l_u(j)) = l_d(j)
                        l_c(l_c(j)) = l_c(l_c(j)) - 1 'decrement column l_c(j) count
                        j = l_r(j)          'j <- next data object in l_p's row
                    Wend
                    l_p = l_d(l_p)          'l_p <- next data object in column l_q
                Wend

            End If
        Else
            If (l_p <> l_q) Then 'l_p's row has been searched, so we must ...

                'uncover the columns of all other data objects h in l_p's row  (right to left)
                h = l_l(l_p)              'h <- initial data object in l_p's row
                While (h <> l_p)        'for each data object h in l_p's row ...
                    'uncover h's column l_c(h)
                    l_l(l_r(l_c(h))) = l_c(h)     'restore l_c(h) to the column header list
                    l_r(l_l(l_c(h))) = l_c(h)
                    i = l_u(l_c(h))           'i <- bottom data object in column l_c(h)
                    While (i <> l_c(h))     'for each data object i in column l_c(h) ...  (in reverse order of "cover h's column l_c(h)")
                        j = l_l(i)              'j <- initial data object in i's row
                        While (j <> i)        'for all other data objects j in i's row ...
                            l_u(l_d(j)) = j           'restore it to its column
                            l_d(l_u(j)) = j
                            l_c(l_c(j)) = l_c(l_c(j)) + 1 'increment column l_c(j) count
                            j = l_l(j)          'j <- next data object in i's row
                        Wend
                        i = l_u(i)          'i <- next data object in column l_c(h)
                    Wend
                    h = l_l(h)          'h <- next data object in l_p's row
                Wend

                If ((m_Delay <> 0) And ((m_Running And 4) = 0)) Then 'remove pentomino from chessboard and momentarily pause
                    i = l_p
                    While ((l_c(i) < 65) Or (l_c(i) > 76))
                        i = l_r(i)
                    Wend
                    RemovePentomino l_c(i) - 64
                    DoEvents
                    Sleep m_Delay
                End If

            End If
            l_p = l_d(l_p)        'l_p <- next data object in column l_q
            If (l_p = l_q) Then 'l_p exhausted

                'uncover column l_q
                l_l(l_r(l_q)) = l_q           'restore l_q to the column header list
                l_r(l_l(l_q)) = l_q
                l_p = l_u(l_q)              'l_p <- bottom data object in column l_q
                While (l_p <> l_q)        'for each data object l_p in column l_q ...  (in reverse order of "cover column l_q")
                    j = l_l(l_p)              'j <- initial data object in l_p's row
                    While (j <> l_p)        'for all other data objects j in l_p's row ...
                        l_u(l_d(j)) = j           'restore it to its column
                        l_d(l_u(j)) = j
                        l_c(l_c(j)) = l_c(l_c(j)) + 1 'increment column l_c(j) count
                        j = l_l(j)          'j <- next data object in l_p's row
                    Wend
                    l_p = l_u(l_p)          'l_p <- next data object in column l_q
                Wend

                'backtrack one level
                If (l_k = 0) Then Exit Do 'search complete
                l_k = l_k - 1
                l_p = l_o(l_k)
                l_q = l_c(l_p)
            Else

                'cover the columns of all other data objects h in l_p's row  (left to right)
                h = l_r(l_p)              'h <- initial data object in l_p's row
                While (h <> l_p)        'for each data object h in l_p's row ...
                    'cover h's column l_c(h)
                    l_l(l_r(l_c(h))) = l_l(l_c(h))  'remove l_c(h) from the column header list
                    l_r(l_l(l_c(h))) = l_r(l_c(h))
                    i = l_d(l_c(h))           'i <- top data object in column l_c(h)
                    While (i <> l_c(h))     'for each data object i in column l_c(h) ...
                        j = l_r(i)              'j <- initial data object in i's row
                        While (j <> i)        'for all other data objects j in i's row ...
                            l_u(l_d(j)) = l_u(j)        'remove it from its column
                            l_d(l_u(j)) = l_d(j)
                            l_c(l_c(j)) = l_c(l_c(j)) - 1 'decrement column l_c(j) count
                            j = l_r(j)          'j <- next data object in i's row
                        Wend
                        i = l_d(i)          'i <- next data object in column l_c(h)
                    Wend
                    h = l_r(h)          'h <- next data object in l_p's row
                Wend

                If ((m_Delay <> 0) And ((m_Running And 4) = 0)) Then 'move pentomino into chessboard and momentarily pause
                    PlacePentomino l_u, l_d, l_l, l_r, l_c, l_p
                    DoEvents
                    Sleep m_Delay
                End If

                l_o(l_k) = l_p               'include l_p's row in the solution vector.
                If (l_r(0) = 0) Then Exit Do 'state is at exact cover solution

                l_k = l_k + 1 'advance to next level in search tree
                l_q = 0
                
                If (m_Running <> 1) Then Exit Do
            End If
        End If
    Loop
    If (m_Running And 4) Then        'form unloaded
        PentominoSearchFindNext = -1
    ElseIf (l_k = 0) Then            'search complete
        PentominoSearchFindNext = 3
    ElseIf (m_Running And 2) Then    'search paused
        PentominoSearchFindNext = 1
    Else 'If (l_r(0) = 0) Then         'solution found
        PentominoSearchFindNext = 0
    End If
    m_Running = 0
End Function

Private Sub cmdPentomino_Click()
    Dim i As Long

    If (m_Running = 0) Then
        If (l_k = 0) Then
            cmdClear_Click
            PentominoSearchInit
        End If
        cmdPentomino.Caption = "<Pause>"
        cmdClear.Enabled = False
        Select Case PentominoSearchFindNext
        Case -1 'form unloaded
        Case 0  'solution found
            cmdClear.Enabled = True
            cmdPentomino.Caption = "Next"
            l_b = l_b + 1
            If (m_Delay = 0) Then 'move pentominos into chessboard
                For i = 0 To l_k
                    PlacePentomino l_u, l_d, l_l, l_r, l_c, l_o(i)
                Next i
            End If
            Label1(0).Caption = "Solutions Found: " & CStr(l_b)
        Case 1  'search paused
            cmdClear.Enabled = True
            cmdPentomino.Caption = "Continue"
        Case 3  'search complete
            cmdClear.Enabled = True
            Option1(0).Value = True
            cmdClear_Click
        End Select
    Else
        m_Running = m_Running Or 2 'cancel
    End If
End Sub



'given binary matrix a(0..m-1, 0..n-1) find subset of rows such that each column contains exactly one non-zero
'returns number of solutions found
Public Function DLX_Matrix(a() As Long, m As Long, n As Long) As Long
    Dim b As Long, h As Long, i As Long, j As Long, k As Long, p As Long, q As Long
    Dim l() As Long, r() As Long, u() As Long, d() As Long, c() As Long  'column object
    Dim o() As Long, z As String
'    Dim g() As Long 'variable g() can be used if solutions must include row indexes of input matrix
    Dim timeout    As Long
    Const max_time As Long = 6000 'timeout after 6 seconds
    Const max_list As Long = 12   'maximum number of solutions to display

    p = n 'primary requirements (exact)
    q = n 'total   requirements (primary (exact) + secondary (at most one))

    'allocate memory for root, column header list, and data objects
    k = q
    For i = 0 To m - 1
        For j = 0 To n - 1
            If (a(i, j) <> 0) Then k = k + 1
        Next j
    Next i
    ReDim l(k)     'link left
    ReDim r(k)     'link right
    ReDim u(k)     'link up
    ReDim d(k)     'link down
    ReDim c(k)     'c(1..q) column data objects count  |  c(q+1.. ) points to the data object's column header object {1..q}
    ReDim o(m - 1) 'solution vector
'    ReDim g(k) 'n + 1 To k

    'build 4-way linked representation of the exact cover problem
    'primary (exact) requirements:
    '  1..n       col i in a() has non-zero
    'secondary (at most one) requirements:
    '  none
    l(0) = p       'k=0 is root
    r(p) = 0
    k = 1               'k={1..q}    column header list
    While (k <= p)      'k={1..p} primary (exact) requirements
        c(k) = 0        'column k count
        l(k) = k - 1
        r(k - 1) = k
        d(k) = k
        u(k) = k
        k = k + 1
    Wend
    While (k <= q)      'k={p+1..q} secondary (at most one) requirements
        c(k) = 0        'column k count
        l(k) = k        'left  links to itself
        r(k) = k        'right links to itself
        d(k) = k
        u(k) = k
        k = k + 1
    Wend
    'column header oject complete.  now add data objects.
    For i = 0 To m - 1
        h = 0
        For j = 0 To n - 1
            If (a(i, j) <> 0) Then
                u(k) = u(j + 1)
                u(j + 1) = k
                d(u(k)) = k
                d(k) = j + 1
                c(k) = j + 1
                c(j + 1) = c(j + 1) + 1 'increment column j+1 count
'                g(k) = i
                If (h = 0) Then
                    l(k) = k
                    r(k) = k
                Else
                    l(k) = h
                    l(r(h)) = k
                    r(k) = r(h)
                    r(h) = k
                End If
                h = k
                k = k + 1
            End If
        Next j
    Next i

    AddItemText "Search for subset of rows of a binary matrix such that each column contains exactly one non-zero" & vbCrLf & vbCrLf & MatrixFormat(a, m, n, "")

    timeout = GetTickCount + max_time
    b = 0 'number of solutions found
    k = 0 'level in search tree
    q = 0 'current column
    Do 'initialization complete; continue with the recursive part of the algorithm
        If (q = 0) Then 'establish current column q
            'choose column object q
            i = &H7FFFFFFF 's heuristic:
            j = r(0)       'choose column with minimum number of non-zeroes to minimize branching
            While (j <> 0)
                If (i > c(j)) Then
                    i = c(j)       'c(j) is minimum count yet seen
                    q = j
                End If
                j = r(j)
            Wend
            If (i = 0) Then 'backtrack one level; current state can not lead to a solution because column q is empty
                If ((k = 0) Or (GetTickCount >= timeout)) Then Exit Do  'search complete or timeout
                k = k - 1
                p = o(k)
                q = c(p)
            Else
                'cover column q
                l(r(q)) = l(q)        'remove q from the column header list
                r(l(q)) = r(q)
                p = d(q)              'p <- top data object in column q
                While (p <> q)        'for each data object p in column q ...
                    j = r(p)              'j <- initial data object in p's row
                    While (j <> p)        'for all other data objects j in p's row ...
                        u(d(j)) = u(j)        'remove it from its column
                        d(u(j)) = d(j)
                        c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
                        j = r(j)          'j <- next data object in p's row
                    Wend
                    p = d(p)          'p <- next data object in column q
                Wend
            End If
        Else
            If (p <> q) Then 'p's row has been searched, so we must ...
                'uncover the columns of all other data objects h in p's row  (right to left)
                h = l(p)              'h <- initial data object in p's row
                While (h <> p)        'for each data object h in p's row ...
                    'uncover h's column c(h)
                    l(r(c(h))) = c(h)     'restore c(h) to the column header list
                    r(l(c(h))) = c(h)
                    i = u(c(h))           'i <- bottom data object in column c(h)
                    While (i <> c(h))     'for each data object i in column c(h) ...  (in reverse order of "cover h's column c(h)")
                        j = l(i)              'j <- initial data object in i's row
                        While (j <> i)        'for all other data objects j in i's row ...
                            u(d(j)) = j           'restore it to its column
                            d(u(j)) = j
                            c(c(j)) = c(c(j)) + 1 'increment column c(j) count
                            j = l(j)          'j <- next data object in i's row
                        Wend
                        i = u(i)          'i <- next data object in column c(h)
                    Wend
                    h = l(h)          'h <- next data object in p's row
                Wend
            End If
            p = d(p)        'p <- next data object in column q
            If (p = q) Then 'p exhausted
                'uncover column q
                l(r(q)) = q           'restore q to the column header list
                r(l(q)) = q
                p = u(q)              'p <- bottom data object in column q
                While (p <> q)        'for each data object p in column q ...  (in reverse order of "cover column q")
                    j = l(p)              'j <- initial data object in p's row
                    While (j <> p)        'for all other data objects j in p's row ...
                        u(d(j)) = j           'restore it to its column
                        d(u(j)) = j
                        c(c(j)) = c(c(j)) + 1 'increment column c(j) count
                        j = l(j)          'j <- next data object in p's row
                    Wend
                    p = u(p)          'p <- next data object in column q
                Wend
                'backtrack one level
                If ((k = 0) Or (GetTickCount >= timeout)) Then Exit Do  'search complete or timeout
                k = k - 1
                p = o(k)
                q = c(p)
            Else
                'cover the columns of all other data objects h in p's row  (left to right)
                h = r(p)              'h <- initial data object in p's row
                While (h <> p)        'for each data object h in p's row ...
                    'cover h's column c(h)
                    l(r(c(h))) = l(c(h))  'remove c(h) from the column header list
                    r(l(c(h))) = r(c(h))
                    i = d(c(h))           'i <- top data object in column c(h)
                    While (i <> c(h))     'for each data object i in column c(h) ...
                        j = r(i)              'j <- initial data object in i's row
                        While (j <> i)        'for all other data objects j in i's row ...
                            u(d(j)) = u(j)        'remove it from its column
                            d(u(j)) = d(j)
                            c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
                            j = r(j)          'j <- next data object in i's row
                        Wend
                        i = d(i)          'i <- next data object in column c(h)
                    Wend
                    h = r(h)          'h <- next data object in p's row
                Wend
                o(k) = p    'include p's row in the solution vector.
                If (r(0) = 0) Then 'all primary columns are covered; we are at a solution
                    'visit exact cover solution
    Select Case b
    Case max_list
        AddItemText "... additional solutions being found but not displayed ..." & vbCrLf
        If (max_time <= 10000) Then timeout = GetTickCount + max_time 'timer reset for benchmarking purposes
    Case Is < max_list
        z = "Solution " & CStr(b + 1) & ": (" & CStr(k + 1) & " rows)" & vbCrLf
        For i = 0 To k
            j = o(i) 'c(j) <- leftmost column in i-th row of solution vector
            While (c(l(j)) < c(j))
                j = l(j)
            Wend
            For h = 1 To n
                If (h = c(j)) Then
                    z = z & " 1"
                    j = r(j)
                Else
                    z = z & " 0"
                End If
            Next h
            z = z & vbCrLf
        Next i
        AddItemText z
'        'list row indexes of solution vector
'        z = "solution rows: "
'        For i = 0 To k
'            z = z & CStr(g(o(i))) & " "
'        Next i
'        AddItemText z
    End Select
    b = b + 1
                    'end visit exact cover solution
                Else               'advance to next level in search tree
                    k = k + 1
                    q = 0
                End If
            End If
        End If
    Loop
    AddItemText "Search " & IIf(k = 0, "complete with ", "timeout after ") & CStr(b) & " solution" & IIf(b = 1, "", "s") & " found" & vbCrLf

    DLX_Matrix = b
End Function

'Solve the N queens problem: find ways to place N queens on an NxN chessboard such that no queen is able to capture
'another.  In other words, no two queens can occupy the same row, column, or diagonal.
'returns number of solutions found
Public Function DLX_NQueens(n As Long) As Long
    Dim b As Long, h As Long, i As Long, j As Long, k As Long, p As Long, q As Long
    Dim l() As Long, r() As Long, u() As Long, d() As Long, c() As Long  'column object
    Dim o() As Long, a() As Long, z As String
    Dim timeout    As Long
    Const max_time As Long = 6000 'timeout after 6 seconds
    Const max_list As Long = 12   'maximum number of solutions to display

    ReDim a(n - 1, n - 1)
    p = n * 2       'primary requirements (exact)
    q = 6 * (n - 1) 'total   requirements (primary (exact) + secondary (at most one))

    'allocate memory for root, column header list, and data objects
    k = q + p * p - 4 '0,1..q for root and header list  p*p for row,col,left diagonal,right diagonal  -4 because 4 corners are alone in one of the diagonals
    ReDim l(k)     'link left
    ReDim r(k)     'link right
    ReDim u(k)     'link up
    ReDim d(k)     'link down
    ReDim c(k)     'c(1..q) column data objects count  |  c(q+1.. ) points to the data object's column header object {1..q}
    ReDim o(n - 1) 'solution vector

    'build 4-way linked representation of the exact cover problem
    'primary (exact) requirements:
    '  1..n       row has queen
    '  n+1..p     col has queen
    'secondary (at most one) requirements:
    '  p+1..q     left, right diagonal has queen
    l(0) = p       'k=0 is root
    r(p) = 0
    k = 1               'k={1..q}    column header list
    While (k <= p)      'k={1..p} primary (exact) requirements
        c(k) = 0        'column k count
        l(k) = k - 1
        r(k - 1) = k
        d(k) = k
        u(k) = k
        k = k + 1
    Wend
    While (k <= q)      'k={p+1..q} secondary (at most one) requirements
        c(k) = 0        'column k count
        l(k) = k        'left  links to itself
        r(k) = k        'right links to itself
        d(k) = k
        u(k) = k
        k = k + 1
    Wend
    'Column header oject complete.  Now add rows.  There are NxN rows in the binary matrix, each of which
    'represent placing a queen at location (i,j) on the chessboard.  Each row has 4 non-zero data objects, one
    'for each of the 4 requirements it satisfies, except for the 4 corner locations which have 3.
    For i = 0 To n - 1
        For j = 0 To n - 1
            InsertDLXDataObject u, d, l, r, c, k, 1, i + 1      'row i occupied  1..n
            InsertDLXDataObject u, d, l, r, c, k, 0, j + 1 + n  'col j occupied  n+1..n*2
            h = i + j                                        'left diagonal occupied  n*2+1..n*4-3
            If ((h <> 0) And (h <> 2 * (n - 1))) Then
            InsertDLXDataObject u, d, l, r, c, k, 0, n * 2 + h
            End If
            h = n - 1 + j - i                                'right diagonal occupied  n*4-2..n*6-6
            If ((h <> 0) And (h <> 2 * (n - 1))) Then
            InsertDLXDataObject u, d, l, r, c, k, 0, n * 4 + h - 3
            End If
        Next j
    Next i
    Debug.Assert k - 1 = UBound(c)

    'In his paper, Knuth notes that the order in which the header nodes are linked together at
    'the start of the algorithm can have a significant effect on the running time.  This block
    'of code re-arranges the column object order per Knuth's recommendation for the N Queens
    'problem.  Comment out this block of code to use the simpler row 1..n, then col 1..n order.
    h = 0          'previous col index
    k = n \ 2 + 1  'current  col index
    If ((n And 1) = 0) Then 'n is even
        q = -1     'increment for col index
    Else
        q = 1
    End If
    l(h) = h
    r(h) = h
    For i = 0 To n - 1
        l(k) = h      'insert col object k
        l(r(h)) = k
        r(k) = r(h)
        r(h) = k
        l(k + n) = k  'insert col object k+n
        l(r(k)) = k + n
        r(k + n) = r(k)
        r(k) = k + n
        h = k         'get k for next round
        k = k + q
        If (q < 0) Then
            q = -(q - 1)
        Else
            q = -(q + 1)
        End If
    Next i

    AddItemText "Search for solution(s) to N Queens puzzle with N = " & CStr(n) & vbCrLf

    timeout = GetTickCount + max_time
    b = 0 'number of solutions found
    k = 0 'level in search tree
    q = 0 'current column
    Do 'initialization complete; continue with the recursive part of the algorithm
        If (q = 0) Then 'establish current column q
            'choose column object q
            i = &H7FFFFFFF 's heuristic:
            j = r(0)       'choose column with minimum number of non-zeroes to minimize branching
            While (j <> 0)
                If (i > c(j)) Then
                    i = c(j)       'c(j) is minimum count yet seen
                    q = j
                End If
                j = r(j)
            Wend
            If (i = 0) Then 'backtrack one level; current state can not lead to a solution because column q is empty
                If ((k = 0) Or (GetTickCount >= timeout)) Then Exit Do  'search complete or timeout
                k = k - 1
                p = o(k)
                q = c(p)
            Else
                'cover column q
                l(r(q)) = l(q)        'remove q from the column header list
                r(l(q)) = r(q)
                p = d(q)              'p <- top data object in column q
                While (p <> q)        'for each data object p in column q ...
                    j = r(p)              'j <- initial data object in p's row
                    While (j <> p)        'for all other data objects j in p's row ...
                        u(d(j)) = u(j)        'remove it from its column
                        d(u(j)) = d(j)
                        c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
                        j = r(j)          'j <- next data object in p's row
                    Wend
                    p = d(p)          'p <- next data object in column q
                Wend
            End If
        Else
            If (p <> q) Then 'p's row has been searched, so we must ...
                'uncover the columns of all other data objects h in p's row  (right to left)
                h = l(p)              'h <- initial data object in p's row
                While (h <> p)        'for each data object h in p's row ...
                    'uncover h's column c(h)
                    l(r(c(h))) = c(h)     'restore c(h) to the column header list
                    r(l(c(h))) = c(h)
                    i = u(c(h))           'i <- bottom data object in column c(h)
                    While (i <> c(h))     'for each data object i in column c(h) ...  (in reverse order of "cover h's column c(h)")
                        j = l(i)              'j <- initial data object in i's row
                        While (j <> i)        'for all other data objects j in i's row ...
                            u(d(j)) = j           'restore it to its column
                            d(u(j)) = j
                            c(c(j)) = c(c(j)) + 1 'increment column c(j) count
                            j = l(j)          'j <- next data object in i's row
                        Wend
                        i = u(i)          'i <- next data object in column c(h)
                    Wend
                    h = l(h)          'h <- next data object in p's row
                Wend
            End If
            p = d(p)        'p <- next data object in column q
            If (p = q) Then 'p exhausted
                'uncover column q
                l(r(q)) = q           'restore q to the column header list
                r(l(q)) = q
                p = u(q)              'p <- bottom data object in column q
                While (p <> q)        'for each data object p in column q ...  (in reverse order of "cover column q")
                    j = l(p)              'j <- initial data object in p's row
                    While (j <> p)        'for all other data objects j in p's row ...
                        u(d(j)) = j           'restore it to its column
                        d(u(j)) = j
                        c(c(j)) = c(c(j)) + 1 'increment column c(j) count
                        j = l(j)          'j <- next data object in p's row
                    Wend
                    p = u(p)          'p <- next data object in column q
                Wend
                'backtrack one level
                If ((k = 0) Or (GetTickCount >= timeout)) Then Exit Do  'search complete or timeout
                k = k - 1
                p = o(k)
                q = c(p)
            Else
                'cover the columns of all other data objects h in p's row  (left to right)
                h = r(p)              'h <- initial data object in p's row
                While (h <> p)        'for each data object h in p's row ...
                    'cover h's column c(h)
                    l(r(c(h))) = l(c(h))  'remove c(h) from the column header list
                    r(l(c(h))) = r(c(h))
                    i = d(c(h))           'i <- top data object in column c(h)
                    While (i <> c(h))     'for each data object i in column c(h) ...
                        j = r(i)              'j <- initial data object in i's row
                        While (j <> i)        'for all other data objects j in i's row ...
                            u(d(j)) = u(j)        'remove it from its column
                            d(u(j)) = d(j)
                            c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
                            j = r(j)          'j <- next data object in i's row
                        Wend
                        i = d(i)          'i <- next data object in column c(h)
                    Wend
                    h = r(h)          'h <- next data object in p's row
                Wend
                o(k) = p    'include p's row in the solution vector.
                If (r(0) = 0) Then 'all primary columns are covered; we are at a solution
                    'visit exact cover solution
    Select Case b
    Case max_list
        AddItemText "... additional solutions being found but not displayed ..." & vbCrLf
        If (max_time <= 10000) Then timeout = GetTickCount + max_time 'timer reset for benchmarking purposes
    Case Is < max_list
        z = "Solution " & CStr(b + 1) & ":" & vbCrLf
        For i = 0 To k
            j = o(i) 'c(j) <- leftmost column in i-th row of solution vector
            While (c(l(j)) < c(j))
                j = l(j)
            Wend
            a(c(j) - 1, c(r(j)) - n - 1) = 1
        Next i
        For i = 0 To n - 1
            For j = 0 To n - 1
                If (a(i, j) = 0) Then
                    z = z & " " & Chr$(149)
                Else
                    z = z & " Q"
                End If
                a(i, j) = 0
            Next j
            z = z & vbCrLf
        Next i
        AddItemText z
    End Select
    b = b + 1
                    'end visit exact cover solution
                Else               'advance to next level in search tree
                    k = k + 1
                    q = 0
                End If
            End If
        End If
    Loop
    AddItemText "Search " & IIf(k = 0, "complete with ", "timeout after ") & CStr(b) & " solution" & IIf(b = 1, "", "s") & " found" & vbCrLf

'    AddItemText "Time: " & CStr(GetTickCount + max_time - timeout)
    DLX_NQueens = b
End Function

'solve sudoku puzzle
'n        size of sudoku grid (width and height)
'a(0..n-1,0..n-1)  A sudoku grid.  All elements of a() must be one of x(0..n).
'x(0..n)  An alphabet of values that can appear in matrix a().  x(1..n) are "constant" characters.
'         x(0) is the "wild" character.  Solving the puzzle is done by replacing all instances of x(0)
'         with one of x(1..n) according to the rules of sudoku.  Each element of x(0..n) must be unique,
'         and in the range 0 <= x(i) < 256 because they are treated as ASCII values when printing solutions.
'         ei: Print Chr$(x(i))
Public Function DLX_Sudoku(a() As Long, x() As Long, n As Long) As Long
    Dim b As Long, h As Long, i As Long, j As Long, k As Long, m As Long, p As Long, q As Long
    Dim l() As Long, r() As Long, u() As Long, d() As Long, c() As Long  'column object
    Dim o() As Long, z As String
    Dim timeout    As Long
    Const max_time As Long = 6000 'timeout after 6 seconds
    Const max_list As Long = 12   'maximum number of solutions to display

    Dim v() As Long, y() As Long, jj As Long

    'assign m and validate input
    If ((n <= 0) Or (n > 255)) Then 'n is positive integer?
        i = n
    Else
        m = Sqr(n)
        If (m * m <> n) Then        'n is a square?
            i = n
        Else
            For i = 0 To n          'elements of x(0..n) in the set {0..255} and unique?
                Select Case x(i)
                Case Is < 0:   Exit For
                Case Is > 255: Exit For
                End Select
                j = i + 1
                While (j <= n)
                    If (x(i) = x(j)) Then Exit For
                    j = j + 1
                Wend
            Next i
        End If
    End If
    If (i <= n) Then
        AddItemText "Illegal input" & vbCrLf
        DLX_Sudoku = -1
        Exit Function
    End If

    ReDim y(255) 'build y inverse alpha
    For i = 1 To n
        y(x(i)) = i
    Next i

    p = n * n * 4   'primary requirements (exact)
    q = n * n * 4   'total   requirements (primary (exact) + secondary (at most one))

    'allocate memory for root, column header list, and data objects
    ReDim v(n - 1, n - 1, n - 1) '3 dimensional cube
    For i = 0 To n - 1
        For j = 0 To n - 1
            If (y(a(i, j)) <> 0) Then
                k = y(a(i, j)) - 1
                If (v(i, j, k) <> 0) Then
                    AddItemText "Illegal input" & vbCrLf
                    DLX_Sudoku = -1
                    Exit Function
                End If
                For h = 0 To n - 1 'zero at v(i, j, k), ones cover the rest in 3 dimensions
                    If (i <> h) Then v(h, j, k) = 1
                    If (j <> h) Then v(i, h, k) = 1
                    If (k <> h) Then v(i, j, h) = 1
                    jj = ((i \ m) * m * n + (j \ m) * m) + (h \ m) * n + (h Mod m) 'index 0..n-1 in box
                    If (jj <> (i * n + j)) Then v(jj \ n, jj Mod n, k) = 1
                Next h
            End If
        Next j
    Next i
    k = q 'column object size
    For h = 0 To n - 1
        For i = 0 To n - 1
            For j = 0 To n - 1
                If (v(h, i, j) = 0) Then k = k + 4 '4 non-zero bits per row
            Next j
        Next i
    Next h
    ReDim l(k)     'link left
    ReDim r(k)     'link right
    ReDim u(k)     'link up
    ReDim d(k)     'link down
    ReDim c(k)     'c(1..q) column data objects count  |  c(q+1.. ) points to the data object's column header object {1..q}
    ReDim o(n * n - 1) 'solution vector

    'build 4-way linked representation of the exact cover problem
    'primary (exact) requirements:
    '  1..n^2           data object a(i,j) has digit 1..n
    '  n^2+1..n^2*2     row i has digit 1..n
    '  n^2*2+1..n^2*3   col j has digit 1..n
    '  n^2*3+1..n^2*4   subblock has digit 1..n
    'secondary (at most one) requirements:
    '  none
    l(0) = p       'k=0 is root
    r(p) = 0
    k = 1               'k={1..q}    column header list
    While (k <= p)      'k={1..p} primary (exact) requirements
        c(k) = 0        'column k count
        l(k) = k - 1
        r(k - 1) = k
        d(k) = k
        u(k) = k
        k = k + 1
    Wend
    While (k <= q)      'k={p+1..q} secondary (at most one) requirements
        c(k) = 0        'column k count
        l(k) = k        'left  links to itself
        r(k) = k        'right links to itself
        d(k) = k
        u(k) = k
        k = k + 1
    Wend
    'Column header oject complete.  Now add rows.  Before considering given values in a sudoku grid, there
    'are n^3 rows in the binary matrix, each of which represent placing digit {1..n} at location (0..n-1, 0..n-1)
    'on the grid.  (this number is reduced based on the given values)  Each row has 4 non-zero data objects,
    'one for each of the 4 requirements it satisfies.
    h = k
    For i = 0 To n - 1
        For j = 0 To n - 1
            For k = 0 To n - 1
                If (v(i, j, k) = 0) Then 'i,j can be k+1
                    InsertDLXDataObject u, d, l, r, c, h, 1, i * n + j + 1             'element 0..n^2-1 has a digit
                    InsertDLXDataObject u, d, l, r, c, h, 0, i * n + k + n * n + 1     'row i has digit k
                    InsertDLXDataObject u, d, l, r, c, h, 0, j * n + k + n * n * 2 + 1 'col j has digit k
                    InsertDLXDataObject u, d, l, r, c, h, 0, ((i \ m) * m + (j \ m)) * n + k + n * n * 3 + 1 'mxm square has digit k
                End If
            Next k
        Next j
    Next i
    Debug.Assert h - 1 = UBound(c)

    AddItemText "Search for solution(s) to sudoku puzzle"
    AddItemText SudokuMatrixFormat(a, n)

    timeout = GetTickCount + max_time
    b = 0 'number of solutions found
    k = 0 'level in search tree
    q = 0 'current column
    Do 'initialization complete; continue with the recursive part of the algorithm
        If (q = 0) Then 'establish current column q
            'choose column object q
            i = &H7FFFFFFF 's heuristic:
            j = r(0)       'choose column with minimum number of non-zeroes to minimize branching
            While (j <> 0)
                If (i > c(j)) Then
                    i = c(j)       'c(j) is minimum count yet seen
                    q = j
                End If
                j = r(j)
            Wend
            If (i = 0) Then 'backtrack one level; current state can not lead to a solution because column q is empty
                If ((k = 0) Or (GetTickCount >= timeout)) Then Exit Do  'search complete or timeout
                k = k - 1
                p = o(k)
                q = c(p)
            Else
                'cover column q
                l(r(q)) = l(q)        'remove q from the column header list
                r(l(q)) = r(q)
                p = d(q)              'p <- top data object in column q
                While (p <> q)        'for each data object p in column q ...
                    j = r(p)              'j <- initial data object in p's row
                    While (j <> p)        'for all other data objects j in p's row ...
                        u(d(j)) = u(j)        'remove it from its column
                        d(u(j)) = d(j)
                        c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
                        j = r(j)          'j <- next data object in p's row
                    Wend
                    p = d(p)          'p <- next data object in column q
                Wend
            End If
        Else
            If (p <> q) Then 'p's row has been searched, so we must ...
                'uncover the columns of all other data objects h in p's row  (right to left)
                h = l(p)              'h <- initial data object in p's row
                While (h <> p)        'for each data object h in p's row ...
                    'uncover h's column c(h)
                    l(r(c(h))) = c(h)     'restore c(h) to the column header list
                    r(l(c(h))) = c(h)
                    i = u(c(h))           'i <- bottom data object in column c(h)
                    While (i <> c(h))     'for each data object i in column c(h) ...  (in reverse order of "cover h's column c(h)")
                        j = l(i)              'j <- initial data object in i's row
                        While (j <> i)        'for all other data objects j in i's row ...
                            u(d(j)) = j           'restore it to its column
                            d(u(j)) = j
                            c(c(j)) = c(c(j)) + 1 'increment column c(j) count
                            j = l(j)          'j <- next data object in i's row
                        Wend
                        i = u(i)          'i <- next data object in column c(h)
                    Wend
                    h = l(h)          'h <- next data object in p's row
                Wend
            End If
            p = d(p)        'p <- next data object in column q
            If (p = q) Then 'p exhausted
                'uncover column q
                l(r(q)) = q           'restore q to the column header list
                r(l(q)) = q
                p = u(q)              'p <- bottom data object in column q
                While (p <> q)        'for each data object p in column q ...  (in reverse order of "cover column q")
                    j = l(p)              'j <- initial data object in p's row
                    While (j <> p)        'for all other data objects j in p's row ...
                        u(d(j)) = j           'restore it to its column
                        d(u(j)) = j
                        c(c(j)) = c(c(j)) + 1 'increment column c(j) count
                        j = l(j)          'j <- next data object in p's row
                    Wend
                    p = u(p)          'p <- next data object in column q
                Wend
                'backtrack one level
                If ((k = 0) Or (GetTickCount >= timeout)) Then Exit Do  'search complete or timeout
                k = k - 1
                p = o(k)
                q = c(p)
            Else
                'cover the columns of all other data objects h in p's row  (left to right)
                h = r(p)              'h <- initial data object in p's row
                While (h <> p)        'for each data object h in p's row ...
                    'cover h's column c(h)
                    l(r(c(h))) = l(c(h))  'remove c(h) from the column header list
                    r(l(c(h))) = r(c(h))
                    i = d(c(h))           'i <- top data object in column c(h)
                    While (i <> c(h))     'for each data object i in column c(h) ...
                        j = r(i)              'j <- initial data object in i's row
                        While (j <> i)        'for all other data objects j in i's row ...
                            u(d(j)) = u(j)        'remove it from its column
                            d(u(j)) = d(j)
                            c(c(j)) = c(c(j)) - 1 'decrement column c(j) count
                            j = r(j)          'j <- next data object in i's row
                        Wend
                        i = d(i)          'i <- next data object in column c(h)
                    Wend
                    h = r(h)          'h <- next data object in p's row
                Wend
                o(k) = p    'include p's row in the solution vector.
                If (r(0) = 0) Then 'all primary columns are covered; we are at a solution
                    'visit exact cover solution
    Select Case b
    Case max_list
        AddItemText "... additional solutions being found but not displayed ..." & vbCrLf
        If (max_time <= 10000) Then timeout = GetTickCount + max_time 'timer reset for benchmarking purposes
    Case Is < max_list
        For i = 0 To k
            j = o(i) 'c(j) <- leftmost column in i-th row of solution vector
            While (c(l(j)) < c(j))
                j = l(j)
            Wend
            a((c(j) - 1) \ n, (c(j) - 1) Mod n) = x(((c(r(j)) - 1) Mod n) + 1)
        Next i
        AddItemText "Solution " & CStr(b + 1)
        AddItemText SudokuMatrixFormat(a, n)
    End Select
    b = b + 1
                    'end visit exact cover solution
                Else               'advance to next level in search tree
                    k = k + 1
                    q = 0
                End If
            End If
        End If
    Loop
    AddItemText "Search " & IIf(k = 0, "complete with ", "timeout after ") & CStr(b) & " solution" & IIf(b = 1, "", "s") & " found" & vbCrLf

    DLX_Sudoku = b
End Function

'remove rows of zeroes and duplicate rows from matrix a(0..m-1, 0..n-1)
Private Sub MatrixRidDupes(a() As Long, ByRef m As Long, n As Long)
    Dim h As Long, i As Long, j As Long

    Do While (0 < m) 'row 0 must not be all zeroes
        For j = 0 To n - 1
            If (a(0, j) <> 0) Then Exit Do
        Next j
        m = m - 1
        For j = 0 To n - 1
            a(0, j) = a(m, j)
        Next j
    Loop
    h = 0
    While (h < m - 1)
        i = h + 1
        While (i < m)
            For j = 0 To n - 1
                If (a(i, j) <> 0) Then Exit For
            Next j
            If (j <> n) Then
                For j = 0 To n - 1
                    If (a(i, j) <> a(h, j)) Then Exit For
                Next j
            End If
            If (j = n) Then 'row i is zero or equal to row h
                m = m - 1
                If (i <> m) Then
                    For j = 0 To n - 1
                        a(i, j) = a(m, j)
                    Next j
                End If
            Else
                i = i + 1
            End If
        Wend
        h = h + 1
    Wend
End Sub

'create a random matrix and display in txtMatrixBinary.Text
Private Sub cmdMatrixBinaryRnd_Click()
    Dim i As Long, j As Long, m As Long, n As Long, a() As Long
    Const MinDim As Long = 6
    Const MaxDim As Long = 16

    m = 3 * Int((MaxDim - MinDim + 1) * Rnd + MinDim)
    n = Int((MaxDim - MinDim + 1) * Rnd + MinDim)
    ReDim a(m - 1, n - 1)
    ReDim x(m - 1)
    For i = 0 To m - 1
        For j = 0 To n - 1
            a(i, j) = Int((1.4) * Rnd) 'Int((2) * Rnd)
        Next j
    Next i
    MatrixRidDupes a, m, n
    txtMatrixBinary.Text = MatrixFormat(a, m, n, "0")
End Sub

'read matrix from txtMatrixBinary.Text and call DLX_Matrix()
Private Sub cmdMatrixBinaryDLX_Click()
    Dim m As Long, n As Long, a() As Long

    If (MatrixFormatInv(txtMatrixBinary.Text, a, m, n) = 2) Then
        MatrixRidDupes a, m, n
        If ((m > 0) And (n > 0)) Then DLX_Matrix a, m, n
    End If
End Sub

'display a default sudoku puzzle in txtSudoku.Text
Private Sub cmdSudokuClr_Click()
    Dim i As Long, j As Long, n As Long, a() As Long

    n = 9
    ReDim a(n - 1, n - 1)
    For i = 0 To n - 1
        For j = 0 To n - 1
            a(i, j) = Asc("_")
        Next j
    Next i
    a(0, 0) = 4 + 48
    a(0, 5) = 1 + 48
    a(0, 7) = 2 + 48
    a(1, 1) = 6 + 48
    a(1, 2) = 8 + 48
    a(1, 6) = 5 + 48
    a(1, 8) = 7 + 48
    a(2, 2) = 2 + 48
    a(2, 3) = 7 + 48
    a(2, 8) = 9 + 48
    a(3, 2) = 1 + 48
    a(3, 3) = 4 + 48
    a(3, 5) = 6 + 48
    a(4, 3) = 5 + 48
    a(4, 4) = 8 + 48
    a(4, 5) = 7 + 48
    a(5, 3) = 1 + 48
    a(5, 5) = 9 + 48
    a(5, 6) = 2 + 48
    a(6, 0) = 1 + 48
    a(6, 5) = 8 + 48
    a(6, 6) = 4 + 48
    a(7, 0) = 5 + 48
    a(7, 2) = 3 + 48
    a(7, 6) = 9 + 48
    a(7, 7) = 8 + 48
    a(8, 1) = 9 + 48
    a(8, 3) = 2 + 48
    a(8, 8) = 6 + 48
    txtSudoku.Text = SudokuMatrixFormat(a, n)
End Sub

'read sudoku puzzle from txtSudoku.Text and find solutions
Private Sub cmdSudokuDLX_Click()
    Dim i As Long, n As Long
    Dim a() As Long, x() As Long

    n = 9
    ReDim x(n)
    ReDim a(n - 1, n - 1)
    x(0) = Asc("_")
    For i = 1 To n
        x(i) = 48 + i 'asc("0") +
    Next i
    If SudokuMatrixFormatInv(a, x, n, txtSudoku.Text) Then
        AddItemText "Invalid input matrix"
        Exit Sub
    End If
    
    DLX_Sudoku a, x, n
End Sub

'find solutions to a 4x4 sudoku problem
Private Sub cmdSudokuDLX4x4_Click()
    Dim i As Long, j As Long, k As Long, n As Long
    Dim a() As Long, x() As Long
    Dim t As String

    n = 16 '4x4 sudoku
    ReDim a(n - 1, n - 1)
    ReDim x(n)

    x(0) = Asc("_")             'wild char
    For i = 1 To n
        x(i) = Asc("A") + i - 1 'alphabet {"A".."P"}
    Next i

    t = t & "EJA_ __OL GP__ __H_" & vbCrLf & _
            "G_H_ F__E NMCO PA_L" & vbCrLf & _
            "ML__ K_GP A_J_ _FEO" & vbCrLf & _
            "_FOC BM__ _DE_ I_NG" & vbCrLf
    t = t & "KG__ E_J_ O___ _I__" & vbCrLf & _
            "_NJ_ _G__ MK_P BL_E" & vbCrLf & _
            "_BM_ L___ ___E F__K" & vbCrLf & _
            "OE_F NKMC __HD APGJ" & vbCrLf
    t = t & "NHBO DL__ FAKG E_IP" & vbCrLf & _
            "J__G O___ ___C _BL_" & vbCrLf & _
            "I_KM G_PA __D_ _OF_" & vbCrLf & _
            "__F_ ___K _O_B __JC" & vbCrLf
    t = t & "CO_N _EB_ __LK JGA_" & vbCrLf & _
            "BPG_ _C_N EJ_A __DM" & vbCrLf & _
            "F_EA JOKG D__H _C_I" & vbCrLf & _
            "_I__ __DF CG__ _EKB" & vbCrLf

    SudokuMatrixFormatInv a, x, n, t
    AddItemText "Solve an example 4x4 sudoku problem"
    DLX_Sudoku a, x, n
End Sub

'solve N queens with N = cboNQueens.Text
Private Sub cmdNQueensDLX_Click()
    Dim n As Long

    n = CLng(cboNQueens.Text)  'number of queens
    DLX_NQueens n
End Sub

Private Sub Option1_Click(Index As Integer)
    Frame1(0).Visible = False
    Frame1(1).Visible = False
    Frame1(2).Visible = False
    Frame1(3).Visible = False
    Frame1(Index).Visible = True
    If (Index = 0) Then
        Text1.Visible = False
        Picture2.Visible = True
    Else
        Picture2.Visible = False
        Text1.Visible = True
    End If
End Sub

Private Sub cmdClear_Click()
    If Option1(0).Value Then
        RemovePentomino 1
        RemovePentomino 2
        RemovePentomino 3
        RemovePentomino 4
        RemovePentomino 5
        RemovePentomino 6
        RemovePentomino 7
        RemovePentomino 8
        RemovePentomino 9
        RemovePentomino 10
        RemovePentomino 11
        RemovePentomino 12
        PentominoSearchDestruct
        cmdPentomino.Caption = "Search"
        Label1(0).Caption = "Scott's pentomino problem:  Find all possible ways to place the 12 pentominoes into a " & _
                    "chessboard leaving the center four squares vacant.  DLX will find each of the 65 " & _
                    "essentially different solutions exactly once, and the full set of 8 X 65 = 520 solutions " & _
                    "is easily obtained by rotation and reflection."
    Else
        Text1.Text = ""
    End If
End Sub

Private Sub HScroll1_Scroll()
    If (HScroll1.Value = HScroll1.Max) Then
        lblPentSpeed.Caption = "Max" '"Speed: Max"
    Else
        lblPentSpeed.Caption = CStr(HScroll1.Value) '"Speed: " & CStr(HScroll1.Value)
    End If
End Sub

Private Sub HScroll1_Change()
    Dim m As Double, b As Double
    Const MaxTime As Long = 400 'maximum delay between pentomino moves
    Const MinTime As Long = 30  'minimum delay between pentomino moves

    m = (Log(MinTime) - Log(MaxTime)) / (HScroll1.Max - HScroll1.Min)
    b = Log(MaxTime) - m * HScroll1.Min
    If (HScroll1.Value = HScroll1.Max) Then
        m_Delay = 0
    Else
        m_Delay = Exp(m * HScroll1.Value + b)
    End If
    HScroll1_Scroll
End Sub

Private Sub Form_Load()
    Dim i As Long, j As Long, k As Long
    Dim p As Variant
    Dim colorpent(12) As Long

    colorpent(0) = &H8000000F
    colorpent(1) = &H404000
    colorpent(2) = RGB(110, 45, 55)
    colorpent(3) = RGB(150, 75, 50)
    colorpent(4) = &HFFFF&
    colorpent(5) = &H80FF&
    colorpent(6) = &HFF&
    colorpent(7) = RGB(150, 40, 190)
    colorpent(8) = RGB(70, 160, 190)
    colorpent(9) = RGB(20, 130, 70)
    colorpent(10) = RGB(220, 40, 140)
    colorpent(11) = &HFF00&
    colorpent(12) = RGB(220, 200, 40)

    With m_ReserveSpace
        .Left = 120
        .Right = 120
        .top = 120
        .Bottom = 120
    End With
    Option1(0).ToolTipText = Frame1(0).Caption
    Option1(1).ToolTipText = Frame1(1).Caption
    Option1(2).ToolTipText = Frame1(2).Caption
    Option1(3).ToolTipText = Frame1(3).Caption
    For i = 2 To 14
        cboNQueens.AddItem CStr(i)
    Next i
  
    Label1(1).Caption = "N queens is a puzzle where N queens are to be placed on an NxN chessboard such that " & _
                        "no queen is able to capture another.  In other words, no two queens share the same " & _
                        "row, column, or diagonal."
    Label1(2).Caption = "Sudoku is a puzzle where a 9x9 grid must be filled with digits such that each row, " & _
                        "column, and each of the 9 3x3 subgrids contain all the digits 1 to 9.  The puzzle " & _
                        "begins with the grid partially filled, usually such that only one solution is possible."
    Label1(3).Caption = "Exact cover problems can be represented with a binary matrix where each column is a " & _
                        "requirement to be satisfied, and each row is some action we can take that satisfies " & _
                        "at least one of the requirements.  The objective is to find a subset of rows such that " & _
                        "there is exactly one non-zero in each of the columns."

    'load pentomino board
    m_psize = Shape1(0).Width
    For i = 1 To 61
        Load Shape1(i)
    Next i
    Picture2.BackColor = RGB(255, 255, 255)
    Shape1(61).Move m_psize * 5, m_psize, m_psize * 8, m_psize * 8
    Shape1(61).FillColor = Me.BackColor
    Shape1(60).Move Shape1(61).Left + m_psize * 3, Shape1(61).top + m_psize * 3, m_psize * 2, m_psize * 2
    Shape1(60).FillColor = Picture2.BackColor
    For i = 0 To 59
        Shape1(i).FillColor = colorpent(i \ 5 + 1)
        Shape1(i).BorderColor = Shape1(i).FillColor
    Next i
    For i = 1 To 61
        Shape1(i).Visible = True
    Next i

    'create example pentomino board
    shpPentominoDemo(0).Move (Frame1(0).Width - shpPentominoDemo(0).Width * 8) \ 2, 600
    p = Split("7 7 7 2 2 2 2 2 7 10 7 12 11 11 11 11 10 10 10 12 12 12 11 9 4 10 1 0 0 12 9 9 4 1 1 0 0 9 9 5 4 4 1 1 6 8 5 5 3 4 6 6 6 8 5 5 3 3 3 3 6 8 8 8", " ")
    For i = 0 To 63
        If (i <> 0) Then Load shpPentominoDemo(i)
        shpPentominoDemo(i).Move shpPentominoDemo(0).Left + shpPentominoDemo(0).Width * (i Mod 8), shpPentominoDemo(0).top + shpPentominoDemo(0).Height * (i \ 8)
        shpPentominoDemo(i).FillColor = colorpent(p(i))
        shpPentominoDemo(i).BorderColor = shpPentominoDemo(i).FillColor
        If (i <> 0) Then shpPentominoDemo(i).Visible = True
    Next i

    'create example chessboard
    shpChess(0).Move (Frame1(1).Width - shpChess(0).Width * 8) \ 2, 600
    p = Split("4 0 3 5 7 1 6 2", " ")
    For i = 0 To 7
        If (i <> 0) Then Load lblChess(i)
        lblChess(i).Move shpChess(0).Left + ((shpChess(0).Width - lblChess(i).Width) \ 2) + shpChess(0).Width * i, shpChess(0).top + ((shpChess(0).Height - lblChess(i).Height) \ 2) + shpChess(0).Height * p(i)
        lblChess(i).Visible = True
    Next i
    k = 0
    For i = 1 To 8
        If (k <> 0) Then
            Load shpChess(k)
            shpChess(k).Move shpChess(k - 8).Left, shpChess(k - 8).top + shpChess(k - 1).Height, shpChess(k - 8).Width, shpChess(k - 8).Height
            If (shpChess(k - 8).FillColor = &H0&) Then
                shpChess(k).FillColor = &HFFFFFF
            Else
                shpChess(k).FillColor = &H0&
            End If
            shpChess(k).Visible = True
        End If
        k = k + 1
        For j = 2 To 8
            Load shpChess(k)
            shpChess(k).Move shpChess(k - 1).Left + shpChess(k - 1).Width, shpChess(k - 1).top, shpChess(k - 1).Width, shpChess(k - 1).Height
            If (shpChess(k - 1).FillColor = &H0&) Then
                shpChess(k).FillColor = &HFFFFFF
            Else
                shpChess(k).FillColor = &H0&
            End If
            shpChess(k).Visible = True
            k = k + 1
        Next j
    Next i

    cboNQueens.ListIndex = 6
    HScroll1.Value = HScroll1.Max - 1
    Option1(0).Value = True
    cmdSudokuClr_Click
    cmdMatrixBinaryRnd_Click
    cmdClear_Click

    On Error Resume Next
    Text1.Font.Name = "Lucida Console"
    txtMatrixBinary.Font.Name = "Lucida Console"
    txtSudoku.Font.Name = "Lucida Console"

End Sub

Private Sub Form_Resize()
    Dim NewWidth     As Long
    Dim NewHeight    As Long
    Dim FormRect     As RECT

    If (Me.WindowState <> 1) Then       'only if not minimized
        GetClientRect Me.hwnd, FormRect 'find available space on Form

        NewWidth = FormRect.Right * Screen.TwipsPerPixelX - m_ReserveSpace.Left - m_ReserveSpace.Right - Frame1(0).Width - 120
        NewHeight = FormRect.Bottom * Screen.TwipsPerPixelY - m_ReserveSpace.top - m_ReserveSpace.Bottom
        If (NewWidth > 0) Then
            If (NewHeight > Option1(0).Height + 120) Then
                Text1.Move m_ReserveSpace.Left, m_ReserveSpace.top, NewWidth, NewHeight
                Picture2.Move Text1.Left, Text1.top, Text1.Width, Text1.Height
                Option1(0).Move m_ReserveSpace.Left + NewWidth + 120, m_ReserveSpace.top + NewHeight - Option1(0).Height
                Option1(1).Move Option1(0).Left + Option1(0).Width + 60, Option1(0).top
                Option1(2).Move Option1(1).Left + Option1(1).Width + 60, Option1(0).top
                Option1(3).Move Option1(2).Left + Option1(2).Width + 60, Option1(0).top
                Frame1(0).Move m_ReserveSpace.Left + NewWidth + 120, m_ReserveSpace.top, Frame1(0).Width, NewHeight - Option1(0).Height - 120
                Frame1(1).Move Frame1(0).Left, Frame1(0).top, Frame1(0).Width, Frame1(0).Height
                Frame1(2).Move Frame1(0).Left, Frame1(0).top, Frame1(0).Width, Frame1(0).Height
                Frame1(3).Move Frame1(0).Left, Frame1(0).top, Frame1(0).Width, Frame1(0).Height
                cmdClear.Move Frame1(0).Left + Frame1(0).Width - cmdClear.Width, Option1(0).top
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_Running = m_Running Or 4
End Sub
