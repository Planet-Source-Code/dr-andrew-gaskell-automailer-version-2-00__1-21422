Attribute VB_Name = "basMailer"
Option Explicit
'Declarations for using the MAPIRTF dll
' For the Debug version of the DLL, remark the following lines
    Public Declare Function WriteRTF _
        Lib "mapirtf.dll" _
        Alias "writertf" (ByVal ProfileName As String, _
                          ByVal MessageID As String, _
                          ByVal StoreID As String, _
                          ByVal cText As String) _
        As Integer
'If read RTF from a message is required, uncomment the following
    'Public Declare Function ReadRTF _
    '    Lib "mapirtf.dll" _
    '    Alias "readrtf" (ByVal ProfileName As String, _
    '                   ByVal SrcMsgID As String, _
    '                     ByVal SrcStoreID As String, _
    '               ByRef MsgRTF As String) _
    '    As Integer

    ' For the Debug version of the DLL, un-remark the following lines
    'Public Declare Function WriteRTF _
        Lib "mapirtfd.dll" _
        Alias "writertf" (ByVal ProfileName As String, _
                          ByVal MessageID As String, _
                          ByVal StoreID As String, _
                          ByVal cText As String) _
        As Integer

    'Public Declare Function ReadRTF _
        Lib "mapirtfd.dll" _
        Alias "readrtf" (ByVal ProfileName As String, _
                       ByVal SrcMsgID As String, _
                         ByVal SrcStoreID As String, _
                   ByRef MsgRTF As String) _
        As Integer


Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Public Const SC_CLOSE = &HF060&
Public Const MF_BYCOMMAND = &H0&
