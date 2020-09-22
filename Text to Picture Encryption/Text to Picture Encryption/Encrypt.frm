VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEncrypt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PicEncrypt"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "Encrypt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stsbar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   14
      Top             =   5652
      Width           =   11064
      _ExtentX        =   19526
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "Ready..."
            TextSave        =   "Ready..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Number of characters scanned:"
            TextSave        =   "Number of characters scanned:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Number of pixels scanned:"
            TextSave        =   "Number of pixels scanned:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "00%"
            TextSave        =   "00%"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8280
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   5160
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   6720
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdPicNew 
         Height          =   375
         Left            =   1080
         Picture         =   "Encrypt.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Clear current image."
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdDecrypt 
         Caption         =   "Decrypt"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         ToolTipText     =   "Decrypt contents of image file."
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton cmdPicSave 
         Height          =   375
         Left            =   600
         Picture         =   "Encrypt.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Save current image"
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdPicOpen 
         Height          =   375
         Left            =   120
         Picture         =   "Encrypt.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Open existing image for decryption"
         Top             =   4560
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   4215
         Left            =   120
         ScaleHeight     =   281
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   265
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   1080
         Picture         =   "Encrypt.frx":13D8
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "New text file."
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "Encrypt"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         ToolTipText     =   "Encrypt current text."
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton cmdTxtSave 
         Height          =   375
         Left            =   600
         Picture         =   "Encrypt.frx":190A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Save the current text file."
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdTxtOpen 
         Height          =   375
         Left            =   120
         Picture         =   "Encrypt.frx":1E3C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Open a text file to encrypt."
         Top             =   4560
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   4215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         ToolTipText     =   "Type in the text to be encrypted."
         Top             =   240
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'PicEncrypt v1.1.3
'Written by Aaron Rider
'Copyright ©2002 Aaron Rider

Option Explicit
Dim Red, Green, Blue As Integer

'I am using api calls because PSet and Point seem to have er...unpredictable results

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Sub cmdClear_Click()
    Text1.Text = ""
End Sub

Private Sub cmdEncrypt_Click()
    On Error Resume Next
    Dim i, X, Y  As Integer
    
    Me.MousePointer = vbHourglass
    
    'Let1 RED value, Let2 BLUE value, Let3 GREEN value
    Dim Let1, Let2, Let3, buffer As String
    
    Dim counter As Double
    Picture1.AutoRedraw = True
    stsbar1.Panels(1).Text = "Scanning..."
    X = 1
    prg1.Min = 0
    prg1.Max = Len(Text1.Text) + 3
    Text1 = Text1 & "  "    'Make sure the length works
    
    'Begin running through the characters in the text box
    For i = 1 To (Len(Text1.Text)) Step 3
        Y = Y + 1
        prg1.Value = prg1.Value + 3         'Increment progress bar
        Let1 = Mid(Text1.Text, i, 1)        'Set RED value
        Let2 = Mid(Text1.Text, i + 1, 1)    'Set BLUE value
        Let3 = Mid(Text1.Text, i + 2, 1)    'Set GREEN value
        counter = counter + 3
        
        'Update status bar panels
        stsbar1.Panels(2).Text = "Number of characters scanned: " & counter
        stsbar1.Panels(3).Text = "Number of pixels scanned: " & counter / 3
        stsbar1.Panels(4).Text = Format((prg1.Value / prg1.Max), "percent")
        
        'display encryption output
        Picture1.Picture = Picture1.Image
        SetPixel Picture1.hdc, X, Y, RGB(Asc(Let1), Asc(Let2), Asc(Let3))
        
        'Make sure y does not go beyond picture1.scaleheight
        If Y >= Picture1.ScaleHeight Then
            X = X + 1
            Y = 1
        End If
    Next i
    
    prg1.Value = 0
    stsbar1.Panels(1).Text = "Ready..."
    stsbar1.Panels(4).Text = "100%"
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDecrypt_Click()
    On Error Resume Next
    Dim i, X, Y, z As Integer
    Dim buffer, strText As String
    Dim counter As Double
    Dim startPause As Boolean
    Me.MousePointer = vbHourglass
    
    stsbar1.Panels(4).Text = "Unknown..."
    
    Picture1.AutoRedraw = False
    Text1.Text = ""
    X = 1
    
    stsbar1.Panels(1).Text = "Scanning..."
    
    For i = 1 To (Picture1.ScaleHeight * Picture1.ScaleWidth)
        Y = Y + 1
        buffer = GetPixel(Picture1.hdc, X, Y)
        If buffer <> 0 Then
            startPause = True
        End If
        If buffer = 0 Then
            If startPause = True Then
                stsbar1.Panels(1).Text = "Ready..."
                stsbar1.Panels(4).Text = "100%"
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        
        'Seperate pixel color value into RGB format
        Blue = Int(buffer / 65536)
        Green = Int((buffer - Blue * 65536) / 256)
        Red = Int((buffer - Blue * 65536) - (Green * 256))
        Text1 = Text1 & Chr(Red) & Chr(Green) & Chr(Blue)
        Picture1.PSet (X, Y), (Picture1.Point(X, Y) + 100)
        counter = counter + 3
        stsbar1.Panels(2).Text = "Number of characters scanned: " & counter
        stsbar1.Panels(3).Text = "Number of pixels scanned: " & Int(counter / 3)
        If Y >= Picture1.ScaleHeight Then
            X = X + 1
            Y = 1
        End If
    Next i
    strText = ""
    Me.MousePointer = vbDefault
End Sub



Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdNew_Click()
    Text1 = ""
    Picture1.Cls
    Picture1.Picture = LoadPicture("")
    stsbar1.Panels(1).Text = "Ready..."
    stsbar1.Panels(2).Text = "Number of characters scanned: "
    stsbar1.Panels(3).Text = "Number of pixels scanned: "
    stsbar1.Panels(4).Text = "00%"
    Text1.Text = ""
End Sub

Private Sub cmdPicNew_Click()
    Picture1.Cls
End Sub

Private Sub cmdPicOpen_Click()
    cd1.Filter = "Bitmap Files|*.bmp|All Files|*.*"
    cd1.ShowOpen
    Picture1.Picture = LoadPicture(cd1.FileName)
End Sub

Private Sub cmdPicSave_Click()
    Dim strConfirm As String
    Picture1.AutoRedraw = True
    Picture1.Picture = Picture1.Image
    cd1.Filter = "Bitmap files|*.bmp|All Files|*.*"
    cd1.ShowSave
    If Len(cd1.FileName) = 0 Then Exit Sub
    
    If Len(Dir(cd1.FileName)) <> 0 Then
        strConfirm = MsgBox(cd1.FileName & " allready exists." & vbNewLine & "Are you sure you want to overrite this file?", vbYesNo + vbExclamation)
        If strConfirm = vbNo Then Exit Sub
    End If
        
    If cd1.FileName <> "" Then
        SavePicture Picture1.Picture, cd1.FileName
    End If
End Sub

Private Sub cmdTxtOpen_Click()
    Dim buffer As String
    cd1.Filter = "Text files|*.txt|All Files|*.*"
    cd1.ShowOpen
    If Len(cd1.FileName) = 0 Then Exit Sub
    Text1 = ""
    Open cd1.FileName For Input As #1
    Do Until EOF(1)
        Line Input #1, buffer
        Text1.Text = Text1 & buffer & vbNewLine
    Loop
    
    Close #1
End Sub

Private Sub cmdTxtSave_Click()
    Dim strConfirm As String
    cd1.Filter = "Text files|*.txt|All files|*.*"
    cd1.ShowSave
    If Len(cd1.FileName) = 0 Then Exit Sub
    
    If Len(Dir(cd1.FileName)) <> 0 Then
        strConfirm = MsgBox(cd1.FileName & " allready exists." & vbNewLine & "Are you sure you want to overrite this file?", vbYesNo + vbExclamation)
        If strConfirm = vbNo Then Exit Sub

    End If
    
    Open cd1.FileName For Output As #1
    Print #1, Text1.Text
    Close #1
End Sub
