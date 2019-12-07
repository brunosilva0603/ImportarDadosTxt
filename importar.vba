Private Sub CommandButton1_Click()
Dim Conteudo As String
    Dim lngRowCount As Long
    Dim fOrigem$, fDestino$, Linha$
    
    'Apresenta mensagem para seleção de arquivo na proxima tela
    MsgBox "Na tela seguinte selecione o arquivo texto"
    
    'Através da caixa de diálogo faz uma busca e selecionando o arquivo que será utilizado.
    fOrigem = Application.GetOpenFilename()
   
    Open fOrigem For Input Access Read As #1
    
    fDestino = FreeFile
    
    fDestino = "\temp.txt"
    
    Line Input #1, Linha
    If EOF(1) Then
       Linha = Replace(Linha, vbLf, vbCrLf)
       Close #1
       Open fDestino For Output Access Write As #1
       Print #1, Linha
       Close #1
    End If
   
    Open fDestino For Input Access Read As #1
    
    Range("A2:ZZ5000").Value = Empty 'LIMPA campos de A2 á ZZ5000
    
' RECOMPONDO O CABEÇALHO, ANTERIORMENTE EXCLUIDO PELA FÓRMULA
    
    Range("A2") = "DATA O.F"
    Range("B2") = "COD.CLIENTE"
    Range("C2") = "CLIENTE"
    Range("D2") = "COD.ITEM"
    Range("E2") = "ITEM"
    Range("F2") = "QTD ITEM"
    Range("G2") = "VALOR UNITÁRIO"
    Range("H2") = "VALOR TOTAL"
    Range("I2") = "OBSERVAÇÃO DA FATURA"
    Range("A3").Select

    Do While EOF(1) = False
        Line Input #1, Conteudo
        'Identificador do Header1
                
        If IsNumeric(Mid(Conteudo, 1, 1)) = "1" Then  'envia as informações pra planilha
        
        Cells(ActiveCell.Row, 1) = Mid(Conteudo, 30, 8) 'DATA DA FATURA
        Cells(ActiveCell.Row, 2) = Mid(Conteudo, 178, 15) 'CODIGO DO CLIENTE
        Cells(ActiveCell.Row, 9) = Mid(Conteudo, 450, 500) 'OBSERVAÇÕES
        
        'Pula de linha na arquivo temporário
        'Identificador de Detail
        Line Input #1, Conteudo
            
        Cells(ActiveCell.Row, 4) = Mid(Conteudo, 2, 15) 'CODIGO DO ITEM
        Cells(ActiveCell.Row, 5) = "=VLOOKUP(RC[-1],ITENS!C[-4]:C,2,TRUE)" 'COLOCA PROCV NA CELULA
        Dim lValor As Variant
        Cells(ActiveCell.Row, 3) = "=VLOOKUP(RC[-1],CLIENTES!C[-2]:C[1],2,TRUE)" 'COLOCA PROCV NA CELULA
        Cells(ActiveCell.Row, 6) = Mid(Conteudo, 287, 23) 'QUANTIDADE
        Cells(ActiveCell.Row, 7) = Mid(Conteudo, 325, 23) 'PREÇO INFORMADO

        'Pula de linha na planilha
        Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
        
        
        End If
    Loop
    
    Close 1

'EXCLUI O ARQUIVO TEMPORÁRIO
    Kill fDestino
    
    MsgBox "Fim de execução da macro"
End Sub
