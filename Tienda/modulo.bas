Option Explicit

'Delaraciones del API de windows
Private Declare Function SendMessage _
    Lib "user32" _
    Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long

Private Declare Function GetDialogBaseUnits Lib "user32.dll" () As Long

Private Declare Function GetDC Lib "user32.dll" ( _
    ByVal hWnd As Long) As Long

Private Declare Function SelectObject _
    Lib "gdi32.dll" ( _
        ByVal HDC As Long, _
        ByVal hObject As Long) As Long
    
Private Declare Function GetTextExtentPoint32 _
    Lib "gdi32.dll" _
    Alias "GetTextExtentPoint32A" ( _
        ByVal HDC As Long, ByVal _
        lpString As String, _
        ByVal cbString As Long, ByRef _
        lpSize As SIZE) As Long
    
Private Declare Function ReleaseDC _
    Lib "user32.dll" ( _
        ByVal hWnd As Long, _
        ByVal HDC As Long) As Long

'UDT
Private Type SIZE
    cx As Long
    cy As Long
End Type

'Constantes
Private Const WM_GETFONT As Long = &H31&
Private Const LB_SETTABSTOPS As Long = &H192
Private Const CH = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"


Public Sub SetHeader(ListBox As ListBox, WidthHeader As String)
    Dim Delimitador() As String
    Dim sTab() As Long, x As Long
   
    Delimitador = Split(WidthHeader, ",")
    ReDim sTab(UBound(Delimitador))
      
    For x = LBound(Delimitador) To UBound(Delimitador)
        sTab(x) = CLng(Delimitador(x))
    Next x
    
    x = UBound(sTab) + 1
      
    SendMessage ListBox.hWnd, LB_SETTABSTOPS, x, sTab(0)
End Sub

Public Sub CrearEncabezado(LB As ListBox, WidthHeader As String, _
    ElLabel As Object, LabelCaptions As String)

    Dim s() As String
    Dim Stops() As Single, U As Single, x As Single, y As Single, i As Long
    Dim ParentForm As Form
    
    SetHeader LB, WidthHeader
    
    Set ParentForm = LB.Parent
    

    s = Split(WidthHeader, ",")
    
    ReDim Stops(UBound(s))
      
    For i = LBound(s) To UBound(s)
        Stops(i) = CLng(s(i))
    Next i

    For i = ElLabel.UBound + 1 To UBound(Stops)
        Load ElLabel(i)
        ElLabel(i).Visible = True
    Next i
    
    U = GetDialogUnitsPerPixel(LB.hWnd)
    s = Split(LabelCaptions, ",")
      
    y = LB.Top - (ElLabel(0).Height * 1.1)
    
    For i = 0 To UBound(Stops)
        If i = 0 Then
            x = LB.Left + ParentForm.ScaleX(2 + U, vbPixels, ParentForm.ScaleMode)
        Else
            x = (Stops(i) + 1) * U
            x = ParentForm.ScaleX(x + 0, vbPixels, ParentForm.ScaleMode) + LB.Left
        End If
        
        
            ElLabel(i).Left = x
            ElLabel(i).Top = y
            ElLabel(i).Caption = s(i)
        
    Next i
    
    Set ParentForm = Nothing
End Sub

Private Function GetDialogUnitsPerPixel(ByVal hWnd As Long) As Single
    Dim DB As Long, HDC, hf, WidthChar As Long, oldFont As Long, sSize As SIZE
    
    HDC = GetDC(hWnd)
    
    If HDC = 0 Then Exit Function
        
        hf = SendMessage(hWnd, WM_GETFONT, 0&, ByVal 0&)
    
        oldFont = SelectObject(HDC, hf)
        
        If GetTextExtentPoint32(HDC, CH, Len(CH), sSize) <> 0 Then
           WidthChar = sSize.cx / Len(CH)

           DB = GetDialogBaseUnits And &HFFFF&
           GetDialogUnitsPerPixel = (2 * WidthChar) / DB
        End If
        
        Call SelectObject(HDC, oldFont)
        Call ReleaseDC(hWnd, HDC)
    
End Function