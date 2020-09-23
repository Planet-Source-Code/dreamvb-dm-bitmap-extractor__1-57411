VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "DM Bitmap Extractor Version 1"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8985
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBase 
      BorderStyle     =   0  'None
      Height          =   4680
      Left            =   2685
      ScaleHeight     =   4680
      ScaleWidth      =   6180
      TabIndex        =   2
      Top             =   30
      Width           =   6180
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save Bitmap"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   15
         TabIndex        =   4
         Top             =   4155
         Width           =   1395
      End
      Begin VB.PictureBox picView 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   105
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   3
         Top             =   225
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   150
      Top             =   3195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglist 
      Left            =   165
      Top             =   3750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":12C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   4575
      Left            =   15
      TabIndex        =   1
      Top             =   60
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   8070
      _Version        =   393217
      Style           =   7
      ImageList       =   "imglist"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13556
            MinWidth        =   2364
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   1575
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   1575
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function FreeResource Lib "kernel32.dll" (ByVal hResData As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private ViewPicture As IPictureDisp
Private ResFileName As String
Private ResListArr() As String

Private Function ShowPicture(ResId As String) As Integer
    ShowPicture = LoadResBitmap(ResId)
End Function

Function FilltreeView(twv As TreeView)
Dim I As Long
On Error Resume Next
    
    twv.Nodes.Clear ' Clear all the items in treeview
    twv.Indentation = 20 '  set the indentation to 60
    twv.Nodes.Add , tvwFirst, "TOP", "BITMAP", 2, 1 ' Add the first top item

    For I = LBound(ResListArr) To UBound(ResListArr) ' Loop though the resource list array
       twv.Nodes.Add 1, tvwChild, "A" & I, Right(ResListArr(I), Len(ResListArr(I)) - 1), 3, 4
       ' add the resource list index number to the treeview
    Next
    
    I = 0
    
End Function


Function GetList(lzFile) As Long
Dim I As Long, LibHangle As Long, resIdx As Long
Dim ResCounter As Long
    
    Erase ResListArr ' erase it array
    
    LibHangle = LoadLibrary(lzFile) ' get the long hangle of a file
    
    If LibHangle = 0 Then
        MsgBox "There was an error while loading the file.", vbCritical, frmmain.Caption
        ' exit if we can't get a value
        Exit Function
    End If

    For I = 1 To 99999 ' I have not found anything yet to tell me the number of items so this will have to do
        resIdx = LoadBitmap(LibHangle, "#" & I) ' get the bitmaps offset index
        If resIdx <> 0 Then
            ResCounter = ResCounter + 1 ' add one to our counter
            ReDim Preserve ResListArr(ResCounter) ' resize array
            ResListArr(ResCounter) = "#" & I ' add resources bitmap index offset to the array
        End If
    Next
    
    ' Clean up
    I = 0
    GetList = ResCounter
    ResCounter = 0
    FreeLibrary LibHangle
    
End Function

Function GetFileExt(lzFile As String) As String
Dim ipos As Integer
    ' used to get the files ext eg GetFileExt "hello.txt" returns txt
    ipos = InStrRev(lzFile, ".", Len(lzFile), vbTextCompare)
    GetFileExt = Mid(lzFile, ipos + 1, Len(lzFile))
End Function

Public Function LoadResBitmap(ResName As String) As Long
Dim mHangle As Long
    
    LoadResBitmap = 1 ' Ok value if eveything went fine
    
    mHangle = LoadLibrary(ResFileName) ' Get the files libary habgle
    
    If mHangle = 0 Then ' Check if we found it
        MsgBox "Unable to open file", vbCritical, frmmain.Caption ' Nope o well retuen error
        LoadResBitmap = 0 ' not good
        Exit Function ' stop
    End If
    ' ok seems fine
    LoadResource mHangle, Iret ' load the resource of the file
    lReturned = LoadBitmap(mHangle, ResName) ' get the hangle of the bitmap
    If lReturned = 0 Then LoadResBitmap = 0: Exit Function ' O dear error
    
    Set ViewPicture = BitmapToPicture(lReturned) ' return the picture
    
    ' Free up Time
    FreeLibrary mHangle
    FreeResource hObject

End Function

Public Function BitmapToPicture(ByVal hBmp As Long) As IPicture
' used to convert the Bitmap to a Visuable basic picture type
    If (hBmp = 0) Then Exit Function
    
    Dim vbPic As Picture, tPicConv As PictDesc, IGuid As Guid
    
    ' Fill PictDesc structure with necessary parts:
    With tPicConv
        .cbSizeofStruct = Len(tPicConv)
        .picType = vbPicTypeBitmap
        .hImage = hBmp
    End With
    
    ' Magic GUID for picture
    With IGuid
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
    End With
    
    ' Create a picture object:
    OleCreatePictureIndirect tPicConv, IGuid, True, vbPic
    
    ' return the picture
    Set BitmapToPicture = vbPic
    
End Function

Private Sub cmdsave_Click()
On Error GoTo CanErr:

    With CDLG
        .FileName = "" ' Clear the filename
        .CancelError = True ' Turn on error check
        .DialogTitle = "Save Bitmap" ' update dialogs title
        .Filter = "Windows Bitmap Files(*.bmp)|*.bmp|" ' update dialogs filetypes
        .ShowSave ' show save dialog
        SavePicture picView.Picture, .FileName ' save the picture
    End With
    
    Exit Sub ' stop
    
CanErr:
    If Err = cdlCancel Then Err.Clear
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tv1.Height = (frmmain.ScaleHeight - StatusBar1.Height - tv1.Top)
    ' Line aboves resizes the treeview control to forms hieght
    Line1(0).X2 = frmmain.ScaleWidth: Line1(1).X2 = frmmain.ScaleWidth
    ' Line above adds a 3D lines along the top of the form
    
    PicBase.Width = (frmmain.ScaleWidth - PicBase.Left - 20)
    PicBase.Height = tv1.Height - PicBase.Top
    ' Lines above are used to reize the base picture box that will holder the view below
    
    picView.Left = (PicBase.ScaleWidth - picView.ScaleWidth) \ 2
    picView.Top = (PicBase.ScaleHeight - picView.ScaleHeight) \ 2
    ' Position the view picture box in the center of the base picturebox above
    cmdsave.Top = (PicBase.ScaleHeight - cmdsave.Height) ' Position the save bitmap button
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Clean upcode
    Set picView.Picture = Nothing
    Set ViewPicture = Nothing
    Set frmmain = Nothing
    ResFileName = ""
    Erase ResListArr
    End ' end this program
End Sub

Private Sub mnuabout_Click()
    frmabout.Show vbModal, frmmain 'show this programs about box
End Sub

Private Sub mnuexit_Click()
    Unload frmmain 'unload this program
End Sub

Private Sub mnuopen_Click()
On Error GoTo CanErr:
Dim FileExt As String * 3

    With CDLG
        .CancelError = True ' turn Cancel error on
        .DialogTitle = "Open File" ' update dialogs title
        .Filter = "Appliaction Files(*.exe)|*.exe|Dynamic Link Files(*.dll)|*.dll|Ocx Files(*.ocx)|*.ocx|" ' Update filetypes
        .ShowOpen ' show open dialog
        If Len(.FileName) = 0 Then Exit Sub ' stop if filename len = 0
        FileExt = LCase(GetFileExt(.FileName)) ' get the files ext
        
        If Not (FileExt = "exe" Or FileExt = "dll" Or FileExt = "ocx") Then ' check for vaild file types
            MsgBox "Unsopported file type" _
            & vbCrLf & vbCrLf & "Only files types of ocx,dll,exe are allowed", vbCritical, frmmain.Caption
            ' user selected other file type than supported for display error
            Exit Sub ' stop
        Else
            'Fill a and return the number of bitmap resources found in the selected file
            If GetList(.FileName) = 0 Then ' check if list was loaded
                cmdsave.Enabled = False ' disable save bitmap button
                picView.Visible = False ' hide pictureview
                MsgBox "There were no bitmap resources found in this file.", vbInformation, frmmain.Caption
                tv1.Nodes.Clear ' clear treeview of it's contents
                Exit Sub ' stop
            Else
                ResFileName = .FileName ' Store the filename becuase we need it later
                StatusBar1.Panels(1).Text = .FileName ' Update status bar with file been accessed
                StatusBar1.Panels(2).Text = UBound(ResListArr) ' update status bar with number of bitmaps found
                ' now we can fill our treeview with the resource list created
                FilltreeView tv1 ' Fill the treeview with the bitmap indexs
                cmdsave.Enabled = False ' disable save button
                picView.Visible = False ' hide pictureview
            End If
        End If
    End With
    
    Exit Sub ' stop
    
CanErr:
    If Err = cdlCancel Then Err.Clear
    
End Sub

Private Sub tv1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim resIdx As String ' Used to hold the bitmaps index in the file

    If tv1.Nodes.Count <> 0 Then ' Check if we have some items in the treeview
        If tv1.SelectedItem.Index > 1 Then ' Check if the top most item is selected this is the text BITMAP by the way
            resIdx = "#" & tv1.Nodes(tv1.SelectedItem.Index).Text ' Get the bitmaps index from the key
            If ShowPicture(resIdx) <> 1 Then ' Check if the Index was found in the file
                MsgBox "Unable to view this bitmap", vbInformation, frmmain.Caption ' No it was not o well show error
                picView.Visible = False ' Hide pictureview
                cmdsave.Enabled = False ' disable save bitmap button
                Exit Sub ' stop
            Else ' carry on
                cmdsave.Enabled = True ' enable save bitmap button
                picView.Visible = True ' Hide pictureview
                picView.Picture = ViewPicture ' Show the picture
                Form_Resize ' resize and update
           End If
        End If
    End If
End Sub
