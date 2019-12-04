Sub Importar_dados_txt()
    Dim LocaldoArquivo As String
    Dim N1 As Integer
    Dim ConteudoDaLinha As String
    
    LocaldoArquivo = Application.GetOpenFilename()
    'Através da caixa de diálogo faz uma busca e selecionando o arquivo que será utilizado.
    
    N1 = FreeFile()
    'Atribui o primeiro número de arquivo disponível (E.g.: #1)

    Open LocaldoArquivo For Input As N1
    'Abre o arquivo para fazer busca de dados
    
    Do While EOF(N1) = False
    'Faz o loop no TXT
    
        Line Input #N1, ConteudoDaLinha
        
        If IsNumeric(Mid(ConteudoDaLinha, 30, 8)) = True Then 'envia as informações pra planilha
        
        Cells(ActiveCell.Row, 1) = Mid(ConteudoDaLinha, 30, 8) 'Alimenta a planilha
        Cells(ActiveCell.Row, 2) = Mid(ConteudoDaLinha, 178, 15) 'Alimenta a planilha
        Cells(ActiveCell.Row, 3) = Mid(ConteudoDaLinha, 2, 15) 'Alimenta a planilha
        Cells(ActiveCell.Row, 4) = Mid(ConteudoDaLinha, 17, 250) 'Alimenta a planilha
        Cells(ActiveCell.Row, 5) = Mid(ConteudoDaLinha, 287, 23) 'Alimenta a planilha
        Cells(ActiveCell.Row, 6) = Mid(ConteudoDaLinha, 325, 23) 'Alimenta a planilha
        Cells(ActiveCell.Row, 7) = Mid(ConteudoDaLinha, 450, 500) 'Alimenta a planilha
        Cells(ActiveCell.Row + 1, ActiveCell.Column).Select 'Pula de linha na planilha
        
        End If
    
    Loop
    
    'pula de linha
    
    Close N1
    'Fecha o arquivo (o número em NumArquivo poder ser reutilizado)

'avisa que terminou


End Sub

