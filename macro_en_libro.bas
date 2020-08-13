Attribute VB_Name = "macro_en_libro"
Sub unirTodo()

Application.ScreenUpdating = False

Dim X As Variant

X = Application.GetOpenFilename _
    ("Excel Files (*.xlsx), *.xlsx*", 2, "Abrir archivos", , True)

If IsArray(X) Then
    Workbooks.Add
           
    ANNNNN = ActiveWorkbook.Name
           
'creando las hojas

    Sheets.Add after:=ActiveSheet
    Sheets("Hoja2").Select
    Sheets("Hoja2").Name = "atributosTodos"
    
    Sheets.Add after:=ActiveSheet
    Sheets("Hoja3").Select
    Sheets("Hoja3").Name = "indicadoresTodos"

    Sheets.Add after:=ActiveSheet
    Sheets("Hoja4").Select
    Sheets("Hoja4").Name = "nuevaHoja"
  
    For y = LBound(X) To UBound(X)
    Workbooks.Open X(y)
    b = ActiveWorkbook.Name
    
'Cambiando la primer hoja de donde va a extraer
    Sheets("atributos").Select
    Range("A1:D100").Select 'cambiando el rango que va a extraer
    Selection.Copy
    
'Cambiando el nombre de la hoja a copiar
    Application.Workbooks(ANNNNN).Worksheets("atributosTodos").Activate
    Range("A1").Select 'cambiando donde quiero que empiece a pegar

    Do While Not IsEmpty(ActiveCell)
    ActiveCell.Offset(1, 0).Select
    Loop


    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False ' pegado especial

'Cambiando la segunda hoja de donde va a extraer
    Application.Workbooks(b).Worksheets("indicadores").Activate
    Range("A1:G100").Select 'cambiando el rango que va a extraer
    Selection.Copy

'Cambiando el nombre de la hoja a copiar 2
    Application.Workbooks(ANNNNN).Worksheets("indicadoresTodos").Activate
    Range("A1").Select 'cambiando donde quiero que empiece a pegar
            
    Do While Not IsEmpty(ActiveCell)
    ActiveCell.Offset(1, 0).Select
    Loop
          
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False ' pegado especial
                                          
    Workbooks(b).Application.CutCopyMode = False
    Workbooks(b).Close False
     
    Next
    
End If
Application.ScreenUpdating = False
End Sub




