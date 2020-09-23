VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExtract 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extract Icon Utility"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdOpen 
         Height          =   285
         Left            =   4200
         Picture         =   "frmExtract.frx":05CA
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   600
         Width           =   390
      End
      Begin VB.PictureBox pic32 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1560
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   8
         Top             =   1200
         Width           =   480
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "C:\"
         Top             =   585
         Width           =   2535
      End
      Begin VB.PictureBox pic16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3840
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   7
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "File Name"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Large Icon"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1313
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Small Icon "
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   1313
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   30
      TabIndex        =   2
      Top             =   2310
      Width           =   4935
      Begin VB.CommandButton cmdExtract 
         Caption         =   "&Extract"
         Height          =   345
         Left            =   2285
         TabIndex        =   5
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   345
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdSaveIcon 
         Caption         =   "Save &Icon"
         Height          =   345
         Left            =   1100
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "Save..."
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeColor 
         Caption         =   "Change BackColor"
      End
      Begin VB.Menu mnusepreg 
         Caption         =   "-"
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "Re&fresh"
      End
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ShInfo As modGlobal.SHFILEINFO
Private objPicture As PictureBox
Private objImage As PictureBox

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdExtract_Click()

On Error GoTo ErrorHandler
  Dim FileName As String
  FileName = txtFile.Text
  Dim hLIcon As Long, hSIcon As Long
  Dim imgObj As ListImage
  Dim hIcon As Long
  
  pic32.Cls
  pic16.Cls
  
  Dim r As Long
  'Get a handle to the small icon
  hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
           BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
  'Get a handle to the large icon
  hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
             BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON) '
           
  'If the handle(s) exists, load it into the picture box(es)
  If hLIcon <> 0 Then
    'Large Icon
    With pic32
      Set .Picture = LoadPicture("")
      .AutoRedraw = True
      r = ImageList_Draw(hLIcon, ShInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT)
      .Refresh
    End With
  End If
           
  'If the handle(s) exists, load it into the picture box(es)
  If hSIcon <> 0 Then
    'Small Icon
    With pic16
      Set .Picture = LoadPicture("")
      .AutoRedraw = True
      r = ImageList_Draw(hSIcon, ShInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT)
      .Refresh
    End With
  End If
  
  If hLIcon = 0 And hSIcon = 0 Then
    MsgBox "Please make sure that file does exist and path is correct", vbInformation
  End If
'  Set imgObj = ImageList1.ListImages.Add(ImageList1.ListImages.Count + 1, FileName & "hd", pic32.Image)
Exit Sub
ErrorHandler:
  MsgBox Err.Description
End Sub

Private Sub cmdOpen_Click()

On Error GoTo ErrorHandler

  Dim strFileName As String
   With cdlg
      .CancelError = True
      .DialogTitle = "Open File"
      .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNFileMustExist
      .Filter = "All Files (*.*)|*.*"
       strFileName = vbNullString
      .FileName = vbNullString
      .ShowOpen
      strFileName = .FileName
   End With
   If strFileName <> vbNullString Then
      txtFile.Text = strFileName
   End If
    On Error GoTo 0
  Exit Sub

ErrorHandler:
  If Err.Number = 32755 Then Err.Clear: Exit Sub
  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuChangeColor_Click of Form frmExtract"

End Sub

Private Sub cmdSaveIcon_Click()
 
  SavePicture pic32.Image, "c:\test123.bmp"
End Sub



Private Sub mnuChangeColor_Click()

On Error GoTo ErrorHandler
  
  With cdlg
    .CancelError = True
    .DialogTitle = "Set BackColor"
    .ShowColor
    objPicture.BackColor = .Color
    objPicture.Refresh
  End With
  
  If objPicture.DataChanged = True Then Call cmdExtract_Click
  
  On Error GoTo 0
  Exit Sub

ErrorHandler:
  If Err.Number = 32755 Then Err.Clear: Exit Sub
  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuChangeColor_Click of Form frmExtract"
End Sub

Private Sub mnuFile_Click()
  pic32.Refresh
  pic16.Refresh
End Sub

Private Sub mnuSave_Click()

On Error GoTo mnuSave_Click_Error
Dim strFileName As String
   With cdlg
      .CancelError = True
      .DefaultExt = "*.bmp"
      .DialogTitle = "Save File"
      .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
      .Filter = "Bitmaps (*.bmp)|*.bmp|All Files (*.*)|*.*"
       strFileName = vbNullString
      .FileName = GetFileName
      .ShowSave
      strFileName = .FileName
   End With
   SavePicture objPicture.Image, strFileName
   On Error GoTo 0
Exit Sub
mnuSave_Click_Error:
  If Err.Number = 32755 Then Err.Clear: Exit Sub
  MsgBox "Error " & Err.Number & " (" & Err.Description & ") " & vbCrLf & "in procedure mnuSave_Click of Form frmExtract", vbCritical

End Sub

Private Sub mnuSaveAs_Click()
  On Error GoTo mnuSave_Click_Error
  Dim strFileName As String
     With cdlg
        .CancelError = True
        .DialogTitle = "Save File As"
        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        .Filter = "Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|GIF Images (*.gif)|*.gif|JPEG Images (*.jpg)|*.jpg|Metafiles (*.wmf;*.emf)|*.wmf;*.emf|Icons (*.ico;*.cur)|*.ico;*.cur|All Files (*.*)|*.*"
         strFileName = vbNullString
        .FileName = GetFileName
        .ShowSave
        strFileName = .FileName
     End With
     SavePicture objPicture.Image, strFileName
     On Error GoTo 0
  Exit Sub
mnuSave_Click_Error:
    If Err.Number = 32755 Then Err.Clear: Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") " & vbCrLf & "in procedure mnuSave_Click of Form frmExtract", vbCritical
End Sub
'---------------------------------------------------------------------------------------
' Procedure : GetFileName
' DateTime  : 5/30/2003 12:46
' Author    : rnimbalkar
' Purpose   : To retrieve the file name without any extension.
'---------------------------------------------------------------------------------------
'
Private Function GetFileName() As String
  Dim sFile As String
   On Error GoTo GetFileName_Error

  sFile = Right(txtFile.Text, Len(txtFile.Text) - InStr(1, txtFile.Text, "\"))
  If sFile <> vbNullString Then
    sFile = Left(sFile, InStrRev(sFile, ".") - 1)
  End If
  
  GetFileName = sFile

   On Error GoTo 0
   Exit Function

GetFileName_Error:

  MsgBox "Error " & Err.Number & " (" & Err.Description & ") " & vbCrLf & "in procedure GetFileName of Form frmExtract"
End Function

Private Sub pic16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Set objPicture = pic16

  If pic16.DataChanged = True Then
    mnuSave.Enabled = True
    mnuSaveAs.Enabled = True
    PopupMenu Me.mnuFile
  Else
    mnuSave.Enabled = False
    mnuSaveAs.Enabled = False
    PopupMenu Me.mnuFile
  End If
  
End Sub

Private Sub pic32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Set objPicture = pic32
  
  If pic32.DataChanged = True Then
    mnuSave.Enabled = True
    mnuSaveAs.Enabled = True
    PopupMenu Me.mnuFile
  Else
    mnuSave.Enabled = False
    mnuSaveAs.Enabled = False
    PopupMenu Me.mnuFile
  End If
  
End Sub
