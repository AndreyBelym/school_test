Attribute VB_Name = "mdlRegister"
Public nam As String
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020
Public LPath As String
Public LName As String
Public Type AboutQ
    Quest As String * 255
    AnswerA As String * 50
    AnswerB As String * 50
    AnswerC As String * 50
    AnswerD As String * 50
    AnswerR As String * 1
End Type
Public Enum constTypeOption
    GetOption = 0
    SaveOption = 1
End Enum
Public Enum constKeyOption
    LastPath = 0
    LastName = 1
    Track1 = 2
End Enum
Public Function Options(ByVal TypeOptions As constTypeOption, ByVal Key As constKeyOption, Optional Setting As String)
Select Case TypeOptions
    Case Is = GetOption
        Options = GetSetting("EruditLoto", "Options", Key)
    Case Is = SaveOption
        SaveSetting "EruditLoto", "Options", Key, Setting
    End Select
End Function





