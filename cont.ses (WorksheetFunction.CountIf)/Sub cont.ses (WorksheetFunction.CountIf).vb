Sub contar()
    Dim i As Integer
    Dim intervalo, col
    
    'condicao = todas as linhas da coluna "ofensor" da (planilha 2) TOP OFENSORES
    'intervalo = apenas a coluna (planilha 1) que tiver o nome da CATEGORIA procurada (planilha 2)
    'O FOR percorre a coluna C da (planilha 2) TOP OFENSORES e conta, caso a condição (da planilha 2)
    'esteja na (planilha 1) backlog
    
    For i = 3 To 400 'linhas
        'procura na planilha BACKLOG, onde está a categoria procurada (da planilha TOP OFENSORES).
        Set cell = Sheets(1).Range("BACKLOG!1:1048576").Find(Sheets(2).Range("A" & i).Value)
 
        'diz em qual coluna está a categoria procurada, na planilha BACKLOG
        coluna = Split(cell.Address, "$")(1)
    
        intervalo = coluna & ":" & coluna   'exemplo:  coluna:coluna   Y:Y  X:X ....
        
        Sheets(2).Range("C" & i) = WorksheetFunction.CountIf(Sheets(1).Range(intervalo), Sheets(2).Range("B" & i).Value)
    Next
    
    
    'filtre: do maior para o menor ofensor ( quantidade )
        Range("C3:C400").Select
    ActiveWorkbook.Worksheets("TOP OFENSORES").ListObjects("Tabela1").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("TOP OFENSORES").ListObjects("Tabela1").Sort. _
        SortFields.Add Key:=Range("C3"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("TOP OFENSORES").ListObjects("Tabela1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E10").Select
End Sub
