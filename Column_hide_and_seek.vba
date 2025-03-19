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
        Dim i As Integer

        ' Inicializa as variáveis
        colDesc = -1
        colNCM = -1

        ' Verifica se as colunas já existem
        For i = ws.Cells(10, ws.Columns.Count).End(xlToLeft).Column To 1 Step -1
            If ws.Cells(10, i).Value = "Descrição do item em português" Then
                colDesc = i
            ElseIf ws.Cells(10, i).Value = "NCM" Then
                colNCM = i
            End If
        Next i

        If Target.Value = "BRASIL - RESOLUX DO BRASIL" Then
            ' Adiciona colunas se necessário
            If colDesc = -1 Then
                ws.Columns(Target.Column + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                ws.Cells(10, Target.Column + 1).Value = "Descrição do item em português"
            End If
            If colNCM = -1 Then
                ws.Columns(Target.Column + 2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                ws.Cells(10, Target.Column + 2).Value = "NCM"
            End If
        Else
            ' Remove as colunas se existirem (sempre removendo da maior para a menor para evitar deslocamento incorreto)
            If colNCM <> -1 Then ws.Columns(colNCM).Delete
            If colDesc <> -1 Then ws.Columns(colDesc).Delete
        End If

        Application.ScreenUpdating = True ' Reativa a atualização da tela
    End If
End Sub

