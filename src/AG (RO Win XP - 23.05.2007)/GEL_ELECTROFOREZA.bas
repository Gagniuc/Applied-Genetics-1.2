Attribute VB_Name = "GEL_ELECTROFOREZA"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

Public am_x_segmente As Variant


'##############################################################################################
'###                                                                                        ###
'###                                        GEL                                             ###
'###                                                                                        ###
'##############################################################################################
Function GEL(ByVal nr_secvente As String, ByVal enzima_cu_nr As Variant)
principal.Gelul.Cls
principal.Gelul.Refresh

If nr_secvente = 0 Then nr_secvente = 1
latime_gel = principal.Gelul.ScaleWidth
grosime_linii = latime_gel / nr_secvente

For i = 0 To nr_secvente
DoEvents
sec = Len(secvente_rerstrictate(i + 1))

If sec = 0 Then GoTo 1
Y = Int(((principal.Gelul.ScaleHeight / 100) * (100 / sec)) * principal.timpul.Value)

If (Y > principal.Gelul.ScaleHeight - 1) Then Y = principal.Gelul.ScaleHeight - 1

principal.Gelul.Line ((grosime_linii * i) + 10, Y - 1)-((grosime_linii * i) + grosime_linii - 5, Y), Mai_mult.GEL_CUL.ForeColor, BF
'principal.Gelul.Line (1, Y - 1)-(4, Y), vbWhite, BF

Call Marele_GEL(Y, enzima_cu_nr)

1:
Next i

Gel_vizivil = True
'MsgBox Lista_cu_geluri.Gel_nr_X.Count
ee = Lista_cu_geluri.Gel_nr_X.Count
Load Lista_cu_geluri.Gel_nr_X(ee)
'ArangeControl Lista_cu_geluri.Gel_nr_X(ee), Lista_cu_geluri
    'Lista_cu_geluri.Gel_nr_X(ee).ZOrder 0
    Lista_cu_geluri.Gel_nr_X(ee).Visible = True
'
'If ee = 0 Then
'Lista_cu_geluri.Gel_nr_X(ee).Left = 1
'Else
'Lista_cu_geluri.Gel_nr_X(ee).Left = Lista_cu_geluri.Gel_nr_X(ee - 1).Left '(theForm.ScaleWidth - CtrObj.Width) / 2
'End If
Lista_cu_geluri.Gel_nr_X(ee - 1).Cls
Lista_cu_geluri.Gel_nr_X(ee - 1).Picture = principal.Gelul.Image
Lista_cu_geluri.Gel_nr_X(ee - 1).ToolTipText = nr_secvente & "@" & Mai_mult.enzime_de_test.List(enzima_cu_nr)
Lista_cu_geluri.Gel_nr_X(ee - 1).Refresh

  am_x_segmente = nr_secvente ' ca sa nu fie 0 pt ca var asta o seteaza gelul 2
    'Clipboard.Clear
    'Clipboard.SetData principal.Gelul.Image
    'SendMessage principal.out_put.hwnd, WM_PASTE, 0, 0
'If nr_secvente <> i Then principal.Gelul.Cls
End Function
'##############################################################################################
'##############################################################################################

'##############################################################################################
'##############################################################################################


Function Marele_GEL(ByVal ygrec As Variant, ByVal i As Variant)
latime_gel = principal.Gel_general.ScaleWidth
grosime_linii = (latime_gel / Mai_mult.enzime_de_test.ListCount)
If Mai_mult.enzime_de_test.ListCount < 3 Then grosime_linii = (latime_gel / 8)
principal.Gel_general.Line ((grosime_linii * i) + 5, ygrec)-((grosime_linii * i) + grosime_linii - 10, ygrec), Mai_mult.GEL_CUL.ForeColor, BF
End Function


Function Trece_OUT_PUT_GEL(ByVal nr As Variant)
    Clipboard.Clear
    Clipboard.SetData principal.Gelul.Image
    SendMessage principal.out_put.hwnd, principal.Gelul.Picture, 0, 0
End Function
