Sub RenomearAba()
    Worksheets("Sheet1").Name = "Análise - Segunda-feira"
End Sub

Sub PularLinhas()
  Range("B6:M23").Delete Shift:=xlShiftUp
  Range("B9:E30").Cut Range("B30")
End Sub

Sub CriarBlocoNPS()

    ' Cabeçalho preto
    Range("F1").Value = "Net Promoter Score (NPS) = %Promotores - %Detratores"
    Range("F1:I1").Merge
    Range("F1").Interior.Color = RGB(0, 0, 0)
    Range("F1").Font.Color = RGB(255, 255, 255)
    Range("F1").Font.Bold = True
    Range("F1").HorizontalAlignment = xlCenter
    
    ' Coluna F (NPS)
    Range("F2").Value = "NPS"
    Range("F2:F3").Font.Size = 14
    
    Range("F3").Formula = "=G3 - I3"        ' Fórmula do NPS
    
    Range("F2:F3").Interior.Color = RGB(0, 0, 0)
    Range("F2:F3").Font.Color = RGB(255, 255, 255)
    Range("F2:F3").Font.Bold = True
    Range("F2:F3").HorizontalAlignment = xlCenter

    ' Promotores (verde)
    Range("G2").Value = "Promotores (9 a 10)"
    Range("G3").Formula = "=SUM(L7:M7)/SUM($C$7:$M$7)"
    Range("G2:G3").Interior.Color = RGB(173, 217, 158)
    Range("G2:G3").Font.Bold = True
    
    Range("G2").Font.Size = 10
    Range("G3").Font.Size = 12
    Range("G2:G3").HorizontalAlignment = xlCenter

    ' Passivos (amarelo)
    Range("H2").Value = "Passivos (7 a 8)"
    Range("H3").Formula = "=SUM(J7:K7)/SUM($C$7:$M$7)"
    Range("H2:H3").Interior.Color = RGB(240, 228, 66)
    Range("H2:H3").Font.Bold = True
    
    Range("H2").Font.Size = 10
    Range("H3").Font.Size = 12
    Range("H2:H3").HorizontalAlignment = xlCenter

    ' Detratores (rosa)
    Range("I2").Value = "Detratores (0 a 6)"
    Range("I3").Formula = "=SUM(C7:I7)/SUM($C$7:$M$7)"
    Range("I2:I3").Interior.Color = RGB(200, 150, 150)
    Range("I2:I3").Font.Bold = True
    
    Range("I2").Font.Size = 10
    Range("I3").Font.Size = 12
    Range("I2:I3").HorizontalAlignment = xlCenter

    ' Ajustar colunas
    Range("F1:I3").NumberFormat = "0%"
    Range("F:I").ColumnWidth = 15
End Sub

Sub Criar_RenomearAba()
  Worksheets.Add
  Worksheets("Planilha1").Name = "Respostas - Segunda-feira"
  ThisWorkbook.Sheets("Respostas - Segunda-feira").Move After:=ThisWorkbook.Sheets("Análise - Segunda-feira")
End Sub

Sub ImportarDadosDeRespostas()

    Dim CaminhoDoArquivo As String
    Dim NomeDoArquivo As String
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim UltimaLinhaOrigem As Long
    Dim UltimaColunaOrigem As Long
    
    ' --- AJUSTE AQUI ---
    ' Defina o caminho e nome do arquivo de respostas externo
    CaminhoDoArquivo = "C:\Users\marcu\Downloads\Cursos\NEO4\MOD-22\FEEDBACK\"
    NomeDoArquivo = "resposta.xlsx"
    Set wsDestino = ThisWorkbook.Sheets("Respostas - Segunda-feira") ' Sua aba de destino
    ' ------------------
    
    If Dir(CaminhoDoArquivo & NomeDoArquivo) = "" Then
        MsgBox "Arquivo de origem não encontrado: " & NomeDoArquivo, vbCritical
        Exit Sub
    End If
    
    Set wbOrigem = Workbooks.Open(CaminhoDoArquivo & NomeDoArquivo)
    Set wsOrigem = wbOrigem.Sheets(1)
    wsOrigem.UsedRange.Copy
    
    wsDestino.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wbOrigem.Close SaveChanges:=False ' Fecha sem salvar alterações
    
    Application.CutCopyMode = False ' Remove a seleção piscando
    MsgBox "Dados de respostas importados com sucesso!", vbInformation

End Sub

Sub FormatacaoRespostas()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Respostas - Segunda-feira")
    
    ws.Activate
    
    ws.Range("B:J").Delete Shift:=xlShiftLeft
    
    With ws.Columns("B")
    
        .FormatConditions.Add( _
            Type:=xlCellValue, _
            Operator:=xlBetween, _
            Formula1:="0", _
            Formula2:="7").Interior.Color = RGB(255, 199, 206)
        
        .FormatConditions.Add( _
            Type:=xlCellValue, _
            Operator:=xlBetween, _
            Formula1:="7", _
            Formula2:="8").Interior.Color = RGB(255, 235, 156)
        
        .FormatConditions.Add( _
            Type:=xlCellValue, _
            Operator:=xlGreaterEqual, _
            Formula1:="9").Interior.Color = RGB(198, 239, 206)

    End With
    
    MsgBox ActiveSheet.Name
    Range("A:C").ColumnWidth = 15
    
End Sub

Sub ExecutarSubs()
  RenomearAba
  PularLinhas
  CriarBlocoNPS
  Criar_RenomearAba
  ImportarDadosDeRespostas
  FormatacaoRespostas
End Sub
