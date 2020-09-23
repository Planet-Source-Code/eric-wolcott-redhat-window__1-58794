Attribute VB_Name = "ModColors"
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Sub GetRGB(R As Integer, G As Integer, b As Integer, ByVal Color As Long)
    Dim TempValue As Long
    
    'First translate the color from a long v
    '     alue to a short value
    TranslateColor Color, 0, TempValue
    
    'Calculate the red, green, and blue valu
    '     es from the short value
    R = TempValue And &HFF&
    G = (TempValue And &HFF00&) / 2 ^ 8
    b = (TempValue And &HFF0000) / 2 ^ 16
End Sub

Public Function MakeGrey(ByVal Col As ColorConstants) As ColorConstants
    Dim R As Integer, G As Integer, b As Integer
    GetRGB R, G, b, Col 'EXTRACT COLOUR VARIABLES
    Dim X As Integer
    X = (R + G + b) / 3 'GET AVERAGE VALUE OF Each
    MakeGrey = RGB(X, X, X) 'Make the GREY colour
End Function


Public Function MakeBW(ByVal Col As ColorConstants) As ColorConstants
    Dim R As Integer, G As Integer, b As Integer
    GetRGB R, G, b, Col 'EXTRACT COLOUR VARIABLES
    Dim X As Integer
    X = (R + G + b) / 3 'GET AVERAGE VALUE OF Each


    If X < (255 / 2) Then X = 0 Else X = 255 'IF AVERAGE IS LESS THAN HALF OF MAX THEN
        'MAKE BLACK, ELSE MAKE WHITE
        MakeBW = RGB(X, X, X)
    End Function

Public Function AdjustBrightness(ByVal Color As Long, ByVal Amount As Single) As Long
    On Error Resume Next
    
    Dim R(1) As Integer, G(1) As Integer, b(1) As Integer
    
    'get red, green, and blue values
    GetRGB R(0), G(0), b(0), Color
    
    'add/subtract the amount to/from the ori
    '     ginal RGB values
    R(1) = SetBound(R(0) + Amount, 0, 255)
    G(1) = SetBound(G(0) + Amount, 0, 255)
    b(1) = SetBound(b(0) + Amount, 0, 255)
    
    'convert RGB back to Long value
    AdjustBrightness = RGB(R(1), G(1), b(1))
End Function

Private Function SetBound(ByVal Num As Single, ByVal MinNum As Single, ByVal MaxNum As Single) As Single
    If Num < MinNum Then
        SetBound = MinNum
    ElseIf Num > MaxNum Then
        SetBound = MaxNum
    Else
        SetBound = Num
    End If
End Function

Public Function InvertColor(ByVal hdc As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
    Dim hRect As RECT
    SetRect hRect, X1, Y1, X2, Y2
    InvertRect hdc, hRect
End Function


