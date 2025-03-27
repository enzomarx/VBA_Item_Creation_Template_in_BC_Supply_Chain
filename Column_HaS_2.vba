Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ColunasOcultasD_K As Range
    Dim ColunasOcultasC_H As Range
    Dim ws As Worksheet
    Set ws = Me
    
    Set ColunasOcultasD_K = ws.Range("J:Q") ' Colunas D a K
    Set ColunasOcultasC_H = ws.Range("C:H") ' Colunas C a H

    ' Condição para a coluna I
    If Not Intersect(Target, ws.Range("I:I")) Is Nothing Then
        If Target.Value = "NO" Then
            ColunasOcultasD_K.EntireColumn.Hidden = True
        ElseIf Target.Value = "YES" Then
            ColunasOcultasD_K.EntireColumn.Hidden = False
        End If
    End If

    ' Condição para a coluna A
    If Not Intersect(Target, ws.Range("A:A")) Is Nothing Then
        If Target.Value = "PURCHASE" Then
            ColunasOcultasC_H.EntireColumn.Hidden = True
        ElseIf Target.Value = "YES" Then
            ColunasOcultasC_H.EntireColumn.Hidden = False
        ElseIf Target.Value = "PRODUCTION" Then
            ColunasOcultasC_H.EntireColumn.Hidden = False
        End If
    End If

    ' Condição para inserir ou remover colunas com base na coluna R
    If Not Intersect(Target, ws.Range("R:R")) Is Nothing Then
        Application.ScreenUpdating = False ' Evita piscar a tela
        
        Dim colDesc As Integer
        Dim colNCM As Integer
        Dim colX As Integer
        Dim colY As Integer
        Dim colZ As Integer
        Dim i As Integer
        
        ' Inicializa as variáveis
        colDesc = -1
        colNCM = -1
        colX = -1
        colY = -1
        colZ = -1
        
        ' Verifica se as colunas já existem
        For i = ws.Cells(10, ws.Columns.Count).End(xlToLeft).Column To 1 Step -1
            Select Case ws.Cells(10, i).Value
                Case "Descrição do item em português": colDesc = i
                Case "NCM": colNCM = i
                Case "X": colX = i
                Case "Y": colY = i
                Case "Z": colZ = i
            End Select
        Next i
        
        If Target.Value = "BRASIL - RESOLUX DO BRASIL" Then
            ' Adiciona colunas se necessário
            If colDesc = -1 Then
                ws.Columns(Target.Column + 1).Insert Shift:=xlToRight
                ws.Columns("AD:AD").Copy
                ws.Columns(Target.Column + 1).PasteSpecial Paste:=xlPasteFormats
                ws.Cells(10, Target.Column + 1).Value = "Descrição do item em português"
                ws.Range("B11").Copy
                ws.Range("S11:S100").PasteSpecial Paste:=xlPasteAll
            End If
            If colNCM = -1 Then
                ws.Columns(Target.Column + 2).Insert Shift:=xlToRight
                ws.Columns("AD:AD").Copy
                ws.Columns(Target.Column + 2).PasteSpecial Paste:=xlPasteFormats
                ws.Cells(10, Target.Column + 2).Value = "NCM"
                ws.Range("B11").Copy
                ws.Range("T11:T100").PasteSpecial Paste:=xlPasteAll
            End If
        ElseIf Target.Value = "INDIA - RESOLUX INDIA PRIVATE LTD" Then
            ' Adiciona colunas X, Y, Z se necessário
            If colX = -1 Then
                ws.Columns(Target.Column + 1).Insert Shift:=xlToRight
                ws.Cells(10, Target.Column + 1).Value = "X"
            End If
            If colY = -1 Then
                ws.Columns(Target.Column + 2).Insert Shift:=xlToRight
                ws.Cells(10, Target.Column + 2).Value = "Y"
            End If
            If colZ = -1 Then
                ws.Columns(Target.Column + 3).Insert Shift:=xlToRight
                ws.Cells(10, Target.Column + 3).Value = "Z"
            End If
        Else
            ' Remove as colunas se existirem
            If colZ <> -1 Then ws.Columns(colZ).Delete
            If colY <> -1 Then ws.Columns(colY).Delete
            If colX <> -1 Then ws.Columns(colX).Delete
            If colNCM <> -1 Then ws.Columns(colNCM).Delete
            If colDesc <> -1 Then ws.Columns(colDesc).Delete
        End If
        
        Application.ScreenUpdating = True ' Reativa a atualização da tela
    End If
End Sub

