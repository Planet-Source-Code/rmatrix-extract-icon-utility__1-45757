Attribute VB_Name = "modGlobal"
Option Explicit

Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long                      '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * 260      '  out: display name (or path)
        szTypeName As String * 80          '  out: type name
End Type
Public SFI As SHFILEINFO
Public Const ILD_TRANSPARENT = &H1       'Display transparent
    'ShellInfo Flags
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Public Const SHGFI_LARGEICON = &H0       'Large icon
Public Const SHGFI_SMALLICON = &H1       'Small icon
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
                 Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
                 Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE


Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long


Public Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal Flags&) As Long


Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal rGetIcon As Long) As Long

Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hinst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
