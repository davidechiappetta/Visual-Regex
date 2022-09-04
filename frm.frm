VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm 
   Caption         =   "RegExp (by Davide Chiappetta)"
   ClientHeight    =   11100
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   ScaleHeight     =   740
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   836
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScrollPat 
      Height          =   5085
      Left            =   11520
      Max             =   0
      TabIndex        =   25
      Top             =   780
      Width           =   330
   End
   Begin VB.CheckBox chkTop 
      Caption         =   "top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   0
      TabIndex        =   24
      Top             =   60
      Width           =   705
   End
   Begin VB.HScrollBar HScroll 
      Height          =   315
      Left            =   30
      Max             =   100
      TabIndex        =   23
      Top             =   10740
      Width           =   11505
   End
   Begin VB.ComboBox cmbFontSize 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5850
      Width           =   855
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      BackColor       =   12648384
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3795
      Left            =   11580
      Max             =   0
      TabIndex        =   6
      Top             =   6930
      Width           =   330
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3885
      Left            =   30
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   765
      TabIndex        =   5
      Top             =   6840
      Width           =   11505
   End
   Begin VB.CheckBox chkIgnoreCase 
      Caption         =   "ignore case"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1800
      TabIndex        =   2
      Top             =   90
      Width           =   1425
   End
   Begin VB.CheckBox chkGlobal 
      Caption         =   "global"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   780
      TabIndex        =   1
      Top             =   90
      Value           =   1  'Checked
      Width           =   975
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   1170
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":009D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   1590
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   2430
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":01C7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   2850
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":0244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   6
      Left            =   0
      TabIndex        =   12
      Top             =   3270
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":02C1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   7
      Left            =   0
      TabIndex        =   13
      Top             =   3690
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":033E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   8
      Left            =   0
      TabIndex        =   14
      Top             =   4110
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":03BB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   9
      Left            =   0
      TabIndex        =   15
      Top             =   4530
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":0438
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   10
      Left            =   0
      TabIndex        =   16
      Top             =   4950
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":04B5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   11
      Left            =   0
      TabIndex        =   17
      Top             =   5370
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":0532
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   435
      Index           =   3
      Left            =   0
      TabIndex        =   20
      Top             =   2010
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm.frx":05AF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "index char/s (start:end) result from regexp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   6
      Left            =   60
      TabIndex        =   29
      Top             =   6540
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "index selected char/s with mouse (start:end)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   5
      Left            =   60
      TabIndex        =   28
      Top             =   6270
      Width           =   3825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "select index line pattern and index line text  corresponding"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   0
      Left            =   3420
      TabIndex        =   27
      Top             =   450
      Width           =   5625
   End
   Begin VB.Label lblValidRegExp 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valid"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   2100
      TabIndex        =   26
      Top             =   420
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "text font size:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   30
      TabIndex        =   22
      Top             =   5940
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "{F4} de-comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   8820
      TabIndex        =   19
      Top             =   90
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "{F5} comment part of strings of line pattern with char ""·"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   3420
      TabIndex        =   18
      Top             =   90
      Width           =   5355
   End
   Begin VB.Label lblInfoCollectionMatch 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3930
      TabIndex        =   4
      Top             =   6540
      Width           =   7545
   End
   Begin VB.Label lblInfoCaretTxtString 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0:0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3930
      TabIndex        =   3
      Top             =   6270
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pattern: Valid/Invalid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   2010
   End
   Begin VB.Menu pMenu 
      Caption         =   "pMenu"
      Visible         =   0   'False
      Begin VB.Menu pMenuCopy 
         Caption         =   "copy"
      End
      Begin VB.Menu pMenuPaste 
         Caption         =   "paste"
      End
      Begin VB.Menu pMenuCut 
         Caption         =   "cut"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadALL 
         Caption         =   "load pattern and string"
      End
      Begin VB.Menu mnuSaveALL 
         Caption         =   "save pattern and string"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "?"
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecuteA Lib "shell32.dll" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long



Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1


Const STRING_ALLOW = " ""!#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"

Dim g_pattern As String
Dim g_string As String

Private Type POINTSINGLE
        X As Single
        Y As Single
End Type
Dim ptMouseDown As POINTSINGLE

Private Type t_pattern
    pattern As String
    testo As String
    iSelectPat As Long
    arPar(100) As String
End Type
Dim p As t_pattern



Private Type t_gen
    maxLine As Long 'max lines on frame visible
    maxChar As Long 'max char on line
    cy As Long
    cx As Long
    iSelLine As Long
    iSelCharStart As Long
    iSelCharEnd As Long
    ShowCaret As Boolean
    ar(100) As String
End Type
Private gen As t_gen

Private Type t_match
    firstIndex As Long
    length As Long
    isMatch As Boolean
End Type
Private arMatch(100) As t_match



Private Sub chkGlobal_Click()
    processaRegExp
    draw
End Sub

Private Sub chkIgnoreCase_Click()
    processaRegExp
    draw
End Sub







Private Function getTextValidFromComment(ByVal s As String) As String
    t = ""
    For a = 1 To Len(s)
        ch = Mid(s, a, 1)
        If ch = "·" And comment = False Then
            comment = True
        ElseIf ch = "·" And comment = True Then
            comment = False
            GoTo salto
        End If
        
        If comment = False Then
            t = t & ch
        End If
salto:
    Next a

    getTextValidFromComment = t
End Function

Private Sub colorizaCommentiPattern(ByVal Index As Integer)

    LockWindowUpdate r(Index).hWnd

    r(Index).Tag = "skip"
 
    
    iStart = r(Index).SelStart
    
    s = r(Index).Text
    n = 0
    colora = False
    s = r(Index).Text
    
    r(Index).SelStart = 0
    r(Index).SelLength = Len(s)
    r(Index).SelColor = vbBlack
    
    If InStr(r(Index).Text, "·") = 0 Then
        r(Index).Tag = ""
        LockWindowUpdate False
        r(Index).SelStart = iStart
        Exit Sub
    End If
    
    For a = 1 To Len(s)
        ch = Mid(s, a, 1)
        If ch = "·" And colora = False Then
            colora = True
            i = a - 1
        ElseIf ch = "·" And colora = True Then
            colora = False
            r(Index).SelStart = i
            r(Index).SelLength = n + 1
            r(Index).SelColor = &HD0D0D0
            n = 0
        End If
        
        If colora = True Then
            n = n + 1
        End If



    Next a
    
    If colora = True Then
        r(Index).SelStart = i
        r(Index).SelLength = n
        r(Index).SelColor = &HD0D0D0
    End If
    
    r(Index).SelStart = iStart
    r(Index).Tag = ""
    
    LockWindowUpdate False
End Sub





Private Sub chkTop_Click()
    If chkTop.Value = 1 Then
        Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Sub

Private Sub cmbFontSize_Click()
    If cmbFontSize.Tag = "skip" Then Exit Sub
    initControl "Consolas", cmbFontSize.Text
    draw
End Sub



Private Sub Form_Load()

    p.pattern = getTextValidFromComment(r(0).Text)
    

    m = UBound(p.arPar) - r.Count + 1
    If m < 0 Then m = 0
    VScrollPat.max = m


    'put trash to array for test
    'For a = 0 To UBound(gen.ar)
    '    gen.ar(a) = a & "  " & Rnd(100000000) & Rnd(1000000000)
    'Next a
    
    gen.ar(0) = "234-234-2345"
    gen.ar(1) = "student-id@alumni.school.edu"

    gen.iSelLine = -1

    
    For a = 10 To 40 Step 2
        cmbFontSize.AddItem a
    Next a
    cmbFontSize.Tag = "skip"
    cmbFontSize.ListIndex = 3
    cmbFontSize.Tag = ""
    
    initControl "Consolas", cmbFontSize.Text


    loadALL

    drawPattern
    draw

End Sub

Sub initControl(ByVal fontname As String, ByVal fontsize As Long)
    pic.fontname = fontname
    pic.fontsize = fontsize
    gen.cx = pic.TextWidth("W")
    gen.cy = pic.TextHeight("Gj")
    gen.maxLine = pic.ScaleHeight \ gen.cy
    gen.maxChar = pic.ScaleWidth \ gen.cx
    m = UBound(gen.ar) - (gen.maxLine - 1)
    If m < 0 Then m = 0
    VScroll.max = m
    
End Sub
Sub drawPattern()
    For a = 0 To r.Count - 1
        n = a + VScrollPat.Value
        
        If n = p.iSelectPat Then
            If r(a).BackColor <> &HC0FFC0 Then r(a).BackColor = &HC0FFC0
        Else
            If r(a).BackColor <> vbWhite Then r(a).BackColor = vbWhite
        End If
        
        r(a).Tag = "skip"
        
 

        r(a).Text = p.arPar(n)
 
        
        colorizaCommentiPattern a
        r(a).Tag = ""
    Next a
End Sub
Sub draw()
    pic.Line (0, 0)-(pic.ScaleWidth, pic.ScaleHeight), vbWhite, BF
    
    hs = HScroll.Value * gen.cx
    
    Y = 0
    For a = 0 To gen.maxLine - 1
        n = a + VScroll.Value
        
        
        If n = gen.iSelLine Then
            If a = 0 Then
                pic.Line (0, 0)-(pic.ScaleWidth - 1, Y + gen.cy + 1), vbRed, B
                pic.Line (0, 0)-(pic.ScaleWidth - 2, Y + gen.cy + 2), vbRed, B
            Else
                pic.Line (0, Y - 1)-(pic.ScaleWidth - 1, Y + gen.cy + 1), vbRed, B
                pic.Line (0, Y - 2)-(pic.ScaleWidth - 1, Y + gen.cy + 2), vbRed, B
            End If
 
            
             'il cuore del disegno regexp
            For B = 0 To UBound(arMatch)
                If arMatch(B).isMatch = False Then Exit For
                
                fi = (arMatch(B).firstIndex * gen.cx) - hs
                le = (arMatch(B).length * gen.cx) '+ hs
                
                
                pic.Line (fi, Y)-(fi + le, Y + gen.cy), &HC0C0F0, BF
                pic.Line (fi, Y)-(fi + le, Y + gen.cy), &H505050, B
            Next B
            
            
            'il caret
            'If gen.ShowCaret = True Then
            If gen.ShowCaret = True Then
                caretStart = (gen.iSelCharStart * gen.cx) - hs
                caretEnd = (gen.iSelCharEnd * gen.cx) - hs
                If caretStart <> caretEnd Then
                    pic.Line (caretStart, Y)-(caretEnd, Y + gen.cy), &HA0A0A0, BF
                End If
            'pic.Line (caretEnd + 1, y)-(caretEnd + 1, y + gen.cy), vbBlue, BF
            
                SetCaretPos caretEnd, Y
            End If
            'End If
        End If

        pic.CurrentX = -hs
        pic.CurrentY = Y
        
        pic.Print gen.ar(n)
        
        
        Y = Y + gen.cy
    Next a
End Sub


Function processaRegExp()
    On Error GoTo hell
    Dim rx As New RegExp
    Dim a As Long
    Dim mColl As MatchCollection
    
    lblInfoCollectionMatch.Caption = ""
    Erase arMatch

    If gen.iSelLine = -1 Then
        rx.Global = chkGlobal.Value = 1
        rx.IgnoreCase = chkIgnoreCase.Value = 1
        rx.pattern = p.pattern
        Set mColl = rx.Execute("dummy") 'test is error
        Exit Function
    End If
    
    testo = gen.ar(gen.iSelLine)
    
    
    If Trim(testo) = "" Or Trim(p.pattern) = "" Then
        Exit Function
    End If

    
    
    rx.Global = chkGlobal.Value = 1
    rx.IgnoreCase = chkIgnoreCase.Value = 1
    rx.pattern = p.pattern
    
    Set mColl = rx.Execute(testo)
    
    lblValidRegExp.ForeColor = &H8000&
    lblValidRegExp.Caption = "Valid"

    For a = 0 To mColl.Count - 1
        arMatch(a).firstIndex = mColl.Item(a).firstIndex
        arMatch(a).length = mColl.Item(a).length
        arMatch(a).isMatch = True
        
        s = s & arMatch(a).firstIndex & "(" & arMatch(a).length & "), "
    Next a
    
    
    lblInfoCollectionMatch.Caption = s

Exit Function
hell:
    lblValidRegExp.ForeColor = vbRed
    lblValidRegExp.Caption = "Invalid"
End Function

Function getMin(ByVal min As Long, ByVal max As Long) As Long
    getMin = IIf(min < max, min, max)
End Function
Function getMax(ByVal min As Long, ByVal max As Long) As Long
    getMax = IIf(min > max, min, max)
End Function

Function switchMaxMin(ByRef n1 As Long, ByRef n2 As Long)
    Dim tmp As Long
    
    If n2 < n1 Then
        tmp = n1
        n1 = n2
        n2 = tmp
    End If
End Function
Sub moveCaretToLeft(ByVal n1 As Long, n2 As Long, ByVal offChar As Long, ByVal shift As Integer)
    If gen.iSelCharEnd = 0 Then
        If n1 <> n2 And shift = 1 Then Exit Sub
        gen.iSelCharStart = gen.iSelCharEnd
        draw
        Exit Sub
    ElseIf gen.iSelCharEnd - offChar = 0 And offChar > 0 Then
            
        gen.iSelCharEnd = gen.iSelCharEnd - 1
        If shift = 0 Then gen.iSelCharStart = gen.iSelCharEnd
        
        HScroll.Value = HScroll.Value - 1
        Exit Sub
    End If
    
    If shift = 1 Then
        gen.iSelCharEnd = gen.iSelCharEnd - 1
    Else
        If n1 = n2 Then
            gen.iSelCharStart = gen.iSelCharStart - 1
            gen.iSelCharEnd = gen.iSelCharStart
        Else
             gen.iSelCharStart = gen.iSelCharEnd
        End If
    End If
    draw
    pic.SetFocus

End Sub

Sub moveCaretToRight(ByVal n1 As Long, n2 As Long, ByVal offChar As Long, ByVal shift As Integer)
    If gen.iSelCharEnd = Len(gen.ar(gen.iSelLine)) Then
        If n1 <> n2 And shift = 1 Then Exit Sub
        gen.iSelCharStart = gen.iSelCharEnd
        draw
        Exit Sub
    ElseIf gen.iSelCharEnd - offChar = gen.maxChar Then
            
        gen.iSelCharEnd = gen.iSelCharEnd + 1
        If shift = 0 Then gen.iSelCharStart = gen.iSelCharEnd
        
        HScroll.Value = HScroll.Value + 1
        Exit Sub
    End If
    
    If shift = 1 Then
        gen.iSelCharEnd = gen.iSelCharEnd + 1
    Else
        If n1 = n2 Then
            gen.iSelCharStart = gen.iSelCharStart + 1
            gen.iSelCharEnd = gen.iSelCharStart
        Else
             gen.iSelCharStart = gen.iSelCharEnd
        End If
    End If
    
    If gen.iSelCharStart > gen.maxChar + offChar Then
        HScroll.Value = (gen.iSelCharStart - gen.maxChar)
    End If
    
    
    draw
    pic.SetFocus
End Sub


Private Sub Form_Resize()
    VScrollPat.Left = Me.ScaleWidth - VScrollPat.Width - 1
    VScroll.Left = Me.ScaleWidth - VScroll.Width - 1
    
    For a = 0 To r.Count - 1
        w = VScrollPat.Left - r(a).Left - 3
        If w < 0 Then w = 0
        r(a).Width = w
    Next a
    pic.Width = r(0).Width
    HScroll.Width = pic.Width
    
    w = VScroll.Left - lblInfoCollectionMatch.Left - 3
    If w < 0 Then w = 0
    lblInfoCollectionMatch.Width = w
    
    
    
    HScroll.Top = Me.ScaleHeight - HScroll.Height - 1
    
    
    h = HScroll.Top - pic.Top
    If h < 0 Then h = 0
    pic.Height = h
    
    VScroll.Height = pic.ScaleHeight
    
    gen.maxLine = pic.ScaleHeight \ gen.cy
    m = UBound(gen.ar) - (gen.maxLine - 1)
    If m < 0 Then m = 0
    VScroll.max = m
    
    
    draw
End Sub

Private Sub mnuAbout_Click()
    Call ShellExecuteA(Me.hWnd, "open", "https://www.facebook.com/davide.chiappetta", "", "", 4)
End Sub

Sub mnuLoadALL_Click()
    loadALL
    drawPattern
    draw
End Sub

Private Sub loadALL()
    Dim s As String, riga As String
    Dim i As Long
    Dim acceptPattern As Boolean, acceptText As Boolean

    s = leggifile(App.Path & "\pattern.dat")
    
    For a = 0 To r.Count - 1
        r(a).Tag = "skip"
        r(a).Text = ""
        r(a).Tag = ""
    Next a
    
    For a = 0 To UBound(p.arPar)
        p.arPar(a) = ""
    Next a
    
    For a = 0 To UBound(gen.ar)
        gen.ar(a) = ""
    Next a
    
    acceptPattern = False
    acceptText = False

    
    tmp = Split(s, vbCrLf)
    For a = 0 To UBound(tmp)
        riga = tmp(a)
        'If Trim(riga) = "" Then GoTo salto
        If riga = Chr(1) & "[PATTERN]" & Chr(1) Then
            i = 0
            acceptPattern = True
            GoTo salto
        ElseIf riga = Chr(1) & "[TEXT]" & Chr(1) Then
            i = 0
            acceptPattern = False
            acceptText = True
            GoTo salto
        End If
        
        If acceptPattern = True Then
            p.arPar(i) = tmp(a)
            i = i + 1
        ElseIf acceptText = True Then
            gen.ar(i) = tmp(a)
            i = i + 1
        End If
        
salto:
    Next a
    


End Sub
Sub mnuSaveALL_Click()
    saveAll
End Sub
Private Sub saveAll()
    Dim limit As Long, a As Long
    Dim s As String
    
    s = s & Chr(1) & "[PATTERN]" & Chr(1) & vbCrLf
    limit = 0
    For a = UBound(p.arPar) To 0 Step -1
        If Trim(p.arPar(a)) <> "" Then
            limit = a
            Exit For
        End If
        
    Next a
    
    For a = 0 To limit
        s = s & p.arPar(a) & vbCrLf
    Next a
    
    
    
    s = s & Chr(1) & "[TEXT]" & Chr(1) & vbCrLf
    
    '***********************************************
    limit = 0
    For a = UBound(gen.ar) To 0 Step -1
        If Trim(gen.ar(a)) <> "" Then
            limit = a
            Exit For
        End If
    Next a
    
    For a = 0 To limit
        s = s & gen.ar(a) & vbCrLf
    Next a
    
    If Right(s, 2) = vbCrLf Then s = Mid(s, 1, Len(s) - 2)
    
    SalvaFile App.Path & "\pattern.dat", s
    
End Sub

Public Function leggifile(ByVal filename As String) As String
    Dim s As String
    Dim f As Long
    On Error GoTo hell

    f = FreeFile
    Open filename For Binary As f
    s = Space(LOF(f))
    Get f, , s
    Close f
    
    leggifile = s
Exit Function
hell:
    Close fileNum
    MsgBox Err.Description
End Function

Public Function SalvaFile(ByVal filename As String, ByVal testo As String) As Boolean
On Error GoTo hell
    Open filename For Output As #1
    If testo <> "" Then
        Print #1, Trim(testo)
    End If
    Close #1
    SalvaFile = True
    Exit Function
hell:
    Close #1
    MsgBox (Err.Description)
End Function

Private Sub pic_GotFocus()
    CreateCaret pic.hWnd, 0, 2, gen.cy
    ShowCaret pic.hWnd
    gen.ShowCaret = True
    draw
End Sub
Private Sub pic_LostFocus()
    gen.ShowCaret = False
    HideCaret pic.hWnd
    'DestroyCaret
End Sub


Private Sub pic_KeyDown(KeyCode As Integer, shift As Integer)
    Dim n1 As Long, n2 As Long, offChar As Long

    i = gen.iSelLine
    
    offChar = HScroll.Value
    
    n1 = gen.iSelCharStart
    n2 = gen.iSelCharEnd
    
    'If n1 < 0 Then n1 = 0
    'If n2 < 0 Then n2 = 0
    
    If n1 <> n2 Then
        switchMaxMin n1, n2
    End If
    
    'Me.Caption = KeyCode & " " & shift
    
    If KeyCode = 46 Then 'canc
        If Len(gen.ar(i)) = n1 Then Exit Sub
        If n1 = n2 Then
            t1 = Mid(gen.ar(i), 1, n1)
            t2 = Mid(gen.ar(i), n1 + 2)
        Else
            t1 = Mid(gen.ar(i), 1, n1)
            t2 = Mid(gen.ar(i), n2 + 1)
        End If
        
        gen.ar(i) = t1 & t2

        If n1 <> n2 Then
            gen.iSelCharStart = getMin(n1, n2)
            gen.iSelCharEnd = gen.iSelCharStart
        End If
        
        processaRegExp
        draw
        
        pic.SetFocus
    ElseIf KeyCode = 39 Then 'right ->
    
        moveCaretToRight n1, n2, offChar, shift

    ElseIf KeyCode = 37 Then 'left <-
    
        moveCaretToLeft n1, n2, offChar, shift
        
    ElseIf KeyCode = 38 Then 'up

        If gen.iSelLine = 0 Then Exit Sub
        gen.iSelLine = gen.iSelLine - 1

        If gen.iSelCharStart > Len(gen.ar(gen.iSelLine)) Then
            gen.iSelCharStart = Len(gen.ar(gen.iSelLine))
            gen.iSelCharEnd = gen.iSelCharStart
        End If
        If gen.iSelLine < VScroll.Value Then
            VScroll.Value = VScroll.Value - 1
        Else
            draw
        End If
        pic.SetFocus
    ElseIf KeyCode = 40 Then 'down
        If gen.iSelLine = UBound(gen.ar) Then Exit Sub
        gen.iSelLine = gen.iSelLine + 1
        If gen.iSelCharStart > Len(gen.ar(gen.iSelLine)) Then
            gen.iSelCharStart = Len(gen.ar(gen.iSelLine))
            gen.iSelCharEnd = gen.iSelCharStart
        End If
        If gen.iSelLine > VScroll.Value + gen.maxLine - 1 Then
            VScroll.Value = VScroll.Value + 1
        Else
            draw
        End If
        pic.SetFocus
    ElseIf KeyCode = 86 And shift = 2 Then 'CTRL+V paste
        KeyCode = 0
        pMenuPaste_Click
    ElseIf KeyCode = 67 And shift = 2 Then 'CTRL+C copy
        KeyCode = 0
        pMenuCopy_Click
    ElseIf KeyCode = 88 And shift = 2 Then 'CTRL+X cut
        KeyCode = 0
        pMenuCut_Click
    End If
End Sub



Private Sub pic_KeyPress(KeyAscii As Integer)
    Dim n1 As Long, n2 As Long, i As Long, offChar As Long
    Dim ch As String
    
    
    ch = Chr(KeyAscii)
    
    offChar = HScroll.Value
    
    i = gen.iSelLine
    n1 = gen.iSelCharStart
    n2 = gen.iSelCharEnd
    
    'If n1 < 0 Then n1 = 0
    'If n2 < 0 Then n2 = 0
    
    
    If n1 <> n2 Then
        switchMaxMin n1, n2
    End If

    If InStr(STRING_ALLOW, ch) > 0 Then

        t1 = Mid(gen.ar(i), 1, n1)
        t2 = Mid(gen.ar(i), n2 + 1)

        gen.ar(i) = t1 & ch & t2
        
        moveCaretToRight n1, n2, offChar, 0
        

        processaRegExp
        draw
    ElseIf ch = Chr(8) Then 'back
        If gen.iSelCharStart = 0 Then Exit Sub
        If n1 = n2 Then
            t1 = Mid(gen.ar(i), 1, n1 - 1)
            t2 = Mid(gen.ar(i), n1 + 1)
        Else
            t1 = Mid(gen.ar(i), 1, n1)
            t2 = Mid(gen.ar(i), n2 + 1)
        End If
        
        gen.ar(i) = t1 & t2
        
        moveCaretToLeft n1, n2, offChar, 0


        processaRegExp
        draw
    End If
    
End Sub

Private Sub pic_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    Dim half As Single
    Dim offChar As Long, h As Long
    
    
    If Button = 1 Then
        gen.iSelLine = (Y \ gen.cy) + VScroll.Value
        offChar = HScroll.Value
        
        
        If gen.iSelLine > UBound(gen.ar) Then
            gen.iSelLine = UBound(gen.ar)
        End If
        
        half = X Mod gen.cx
        If half <= (gen.cx / 2) Then
            gen.iSelCharStart = (X \ gen.cx) + offChar
        Else
            gen.iSelCharStart = (X \ gen.cx) + 1 + offChar
        End If
        
        
        If gen.iSelCharStart >= Len(gen.ar(gen.iSelLine)) Then
            gen.iSelCharStart = Len(gen.ar(gen.iSelLine)) '+ offChar
        End If
        gen.iSelCharEnd = gen.iSelCharStart
'
        If gen.iSelCharStart < offChar Then
            h = Len(gen.ar(gen.iSelLine)) - 1
            If h < 0 Then h = 0
            HScroll.Value = h
        End If

        showInfoCaret

        processaRegExp

        draw
    
    End If
End Sub




Sub showInfoCaret()
    Dim n1 As Long, n2 As Long
    n1 = gen.iSelCharStart
    n2 = gen.iSelCharEnd
    'If n1 < 0 Then n1 = 0
    'If n2 < 0 Then n2 = 0
    
    switchMaxMin n1, n2
    
    If n1 = n2 Then
        lblInfoCaretTxtString.Caption = n1 & ":0"
    Else
        lblInfoCaretTxtString.Caption = n1 & ":" & n2 - n1
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        offChar = HScroll.Value
        
        If X < 0 Then X = 0
        If X > pic.ScaleWidth Then X = pic.ScaleWidth - 1
        
        half = X Mod gen.cx
        If half <= (gen.cx / 2) Then
            iChar = (X \ gen.cx)
        Else
            iChar = (X \ gen.cx) + 1
        End If


        If iChar + offChar >= Len(gen.ar(gen.iSelLine)) Then
            gen.iSelCharEnd = Len(gen.ar(gen.iSelLine))
        Else
            gen.iSelCharEnd = iChar + offChar
        End If

        showInfoCaret
        
        draw
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If gen.iSelLine = -1 Then Exit Sub
        If gen.ar(gen.iSelLine) = "" Then Exit Sub
        
        PopupMenu pMenu
    End If
End Sub

Private Sub pMenuCopy_Click()
    Dim n1 As Long, n2 As Long
    
    n1 = gen.iSelCharStart
    n2 = gen.iSelCharEnd
    
    'If n1 < 0 Then n1 = 0
    'If n2 < 0 Then n2 = 0
    
    switchMaxMin n1, n2
    
    s = gen.ar(gen.iSelLine)
    t = Mid(s, n1 + 1, n2 - n1)
    Clipboard.Clear
    Clipboard.SetText t
    pic.SetFocus
End Sub

Private Sub pMenuCut_Click()
    Dim n1 As Long, n2 As Long
    
    n1 = gen.iSelCharStart
    n2 = gen.iSelCharEnd

    switchMaxMin n1, n2

    
    s = gen.ar(gen.iSelLine)
    
    t1 = Mid(s, 1, n1)
    cp = Mid(s, n1 + 1, n2 - n1)
    t2 = Mid(s, n2 + 1)
    
    gen.ar(gen.iSelLine) = t1 & t2
    
    Clipboard.Clear
    Clipboard.SetText cp
    
    gen.iSelCharStart = n1
    gen.iSelCharEnd = n1
    processaRegExp
    draw
    pic.SetFocus
End Sub

Private Sub pMenuPaste_Click()
    Dim n1 As Long, n2 As Long
    
    n1 = gen.iSelCharStart
    n2 = gen.iSelCharEnd
    
    'If n1 < 0 Then n1 = 0
    'If n2 < 0 Then n2 = 0

    
    switchMaxMin n1, n2
    
    cp = Clipboard.GetText
    If cp = "" Then Exit Sub
    
    tmp = Split(cp, vbCrLf)
    cp = tmp(0)
    
    s = gen.ar(gen.iSelLine)
    
    t1 = Mid(s, 1, n1)
    t2 = Mid(s, n2 + 1)
    
    gen.ar(gen.iSelLine) = t1 & cp & t2
    
    gen.iSelCharStart = n1 + Len(cp)
    gen.iSelCharEnd = n1 + Len(cp)
    processaRegExp
    draw
    pic.SetFocus
End Sub



Private Sub r_Change(Index As Integer)
    If r(Index).Tag = "skip" Then Exit Sub
    
    p.arPar(Index + VScrollPat.Value) = r(Index).Text
    
    i = r(Index).SelStart
    colorizaCommentiPattern Index
    r(Index).SelStart = i

    p.pattern = getTextValidFromComment(r(Index).Text)
    processaRegExp
    draw

End Sub



Private Sub r_KeyDown(Index As Integer, KeyCode As Integer, shift As Integer)
    If KeyCode = 115 Then       'F4 de-comment
        i = r(Index).SelStart
        r(Index).SelText = Replace(r(Index).SelText, "·", "")
        colorizaCommentiPattern Index
        r(Index).SelStart = i
        
        p.pattern = getTextValidFromComment(r(Index).Text)
        processaRegExp
        draw
        
        
    ElseIf KeyCode = 116 Then   'F5 comment
    
        r(Index).Tag = "skip"
        If r(Index).SelLength = 0 Then
            r(Index).SelText = "·"
        Else
            r(Index).SelText = "·" & r(Index).SelText & "·"
        End If
        r(Index).Tag = ""
        
        i = r(Index).SelStart
        colorizaCommentiPattern Index
        r(Index).SelStart = i
        
        p.pattern = getTextValidFromComment(r(Index).Text)
        processaRegExp
        draw
        
    End If
End Sub

Sub pairParent(ByVal Index As Integer)
    If r(Index).SelLength = 1 Then
        m = r(Index).SelText
        i = r(Index).SelStart
        s = r(Index).Text
        n = 0
        
        If m = "(" Then
            mm = ")"
            dx = True
        ElseIf m = "[" Then
            mm = "]"
            dx = True
        ElseIf m = "{" Then
            mm = "}"
            dx = True
        ElseIf m = ")" Then
            mm = "("
            sx = True
        ElseIf m = "]" Then
            mm = "["
            sx = True
        ElseIf m = "}" Then
            mm = "{"
            sx = True
        End If
        
        If dx = True Then  'counting left to right
            For a = i + 2 To Len(s)
                ch = Mid(s, a, 1)
                
                If ch = "·" And Skip = False Then       'start comment, counting ignore text
                    Skip = True
                    GoTo salto
                ElseIf ch = "·" And Skip = True Then    'end comment
                    Skip = False
                    GoTo salto
                End If
                
                If Skip = True Then GoTo salto
                
                
                If ch = mm And n = 0 Then
                    r(Index).SelLength = a - i
                    SetFocus
                    Exit For
                End If
                If ch = m Then n = n + 1
                If ch = mm Then n = n - 1
salto:
            Next a
        ElseIf sx = True Then 'reverse order, counting right to left
            For a = i To 1 Step -1
                ch = Mid(s, a, 1)
                
                
                If ch = "·" And Skip = False Then       'start comment, counting ignore text
                    Skip = True
                    GoTo salto2
                ElseIf ch = "·" And Skip = True Then    'end comment, restore text
                    Skip = False
                    GoTo salto2
                End If
                
                If Skip = True Then GoTo salto2
                
                
                If ch = mm And n = 0 Then
                    r(Index).SelStart = a - 1
                    r(Index).SelLength = (i - a) + 2
                    SetFocus
                    Exit For
                End If
                If ch = m Then n = n + 1
                If ch = mm Then n = n - 1
salto2:
            Next a
        End If
        
    End If

End Sub




Private Sub r_LostFocus(Index As Integer)
    l = r(Index).SelLength
    r(Index).SelStart = r(Index).SelStart + l
    r(Index).SelLength = 0
End Sub

Private Sub r_MouseUp(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
    ch = r(Index).SelText
    If Len(ch) = 1 Then
        If InStr("()[]{}", ch) > 0 Then
            pairParent Index
        End If
    End If
End Sub

Sub segnapostoPattern(ByVal i As Long)
    For a = 0 To r.Count - 1
        r(a).BackColor = vbWhite
    Next a
    r(i).BackColor = &HC0FFC0
End Sub


Private Sub r_MouseDown(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
    p.iSelectPat = Index + VScrollPat.Value
    
    p.pattern = getTextValidFromComment(r(Index).Text)
    segnapostoPattern Index
    processaRegExp
    gen.ShowCaret = False
    draw
End Sub







Private Sub VScroll_Change()
    draw
    pic.SetFocus
End Sub

Private Sub VScroll_Scroll()
    VScroll_Change
End Sub
Private Sub hScroll_Change()
    draw
    pic.SetFocus
End Sub

Private Sub hScroll_Scroll()
    hScroll_Change
End Sub

Private Sub VScrollPat_Change()
    drawPattern
End Sub

Private Sub VScrollPat_Scroll()
    VScrollPat_Change
End Sub
