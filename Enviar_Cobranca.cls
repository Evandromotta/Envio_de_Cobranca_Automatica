'Programa em VBA que utiliza a Planila "ENVIAR_COBRANCA" para o disparo automï¿½tico das documentaï¿½ï¿½es para os Clientes

Sub enviar_email()

Set objeto_outlook = CreateObject("Outlook.Application")

For linha = 2 To Cells(2, 10).Value
    
    
    Set Email = objeto_outlook.createitem(0)
'   exibe o e-mail na tela
    
    Email.display
    
    
    Email.To = Cells(linha, 1).Value
'   adiciona o remetente no e-mail
    
    Email.Subject = "CTe e Boleto " & Cells(linha, 5) & "- Gabardo para " & Cells(linha, 3).Value & "*** EM ABERTO ***"
    moeda = Format(Cells(linha, 8).Text, "Currency")
'   adiciona o assunto no email, para inserir a informaï¿½ï¿½o inserida na planilha,
'   basta inserir "& Cells(linha, coluna)" lembrando que a variï¿½vel "linha" mudarï¿½ conforme o for irï¿½ percorre linha por linha
'   e em coluna insira o nï¿½mero correspondet a coluna
'   sempre que pecisar adicionar textos fixos ao e-mail, basta referenciï¿½-los entre aspas "  ".
    
        
    If Cells(linha, 9).Value = Cells(2, 12).Value Then
'   esse "If" informarï¿½ ao cï¿½digo se ï¿½ para enviar e-mail com instruï¿½ï¿½o de cartï¿½rio caso na coluna 12, contenha o nï¿½ 2, caso contrï¿½rio envia sï¿½ a cobranï¿½a normal.

        Email.Body = "Olï¿½ " & Cells(linha, 3).Value & "," & Chr(10) & Chr(10) _
        & "Poderiam verificar o Boleto do CTe " & Cells(linha, 5).Value & " no valor de " & moeda & "?" & Chr(10) & Chr(10) _
        & "Pois venceu dia: " & Cells(linha, 7).Value & " e nï¿½o identificamos o pagamento." & Chr(10) & Chr(10) _
        & Chr(10) & Chr(10) _
        & "Fico a disposiï¿½ï¿½o," & Chr(10) & Chr(10) _
        & " Nome da Empresa" & Chr(10) & Chr(10) _
        & "fone: 01 1234 5678 Ramal 3022" & Chr(10) & Chr(10) _
        & "whats: 01 123456789  "
    
'       Abaixo segue as linhas de cï¿½digos responsï¿½veis por anexar a documentaï¿½ï¿½o do e-mail, neste caso boleto e Nota
'       Lembrando que as informaï¿½ï¿½es fixas estï¿½o entres aspas e as informaï¿½ï¿½es que mudam conforme as linhas da Planilha,
'       entï¿½o com as variï¿½veis:"& Cells(linha, coluna)"
    
        Email.Attachments.Add (ThisWorkbook.Path & "\" & Cells(linha, 4).Value & "-" & Cells(linha, 5).Value & "-" & Cells(linha, 6).Value & "(DocumentoEletronico).pdf")
'       anexo da Nota
        Email.Attachments.Add (ThisWorkbook.Path & "\" & "Boleto-" & Cells(linha, 5).Value & ".pdf")
'       anexo do Boleto
        Email.send
        
    Else
        Email.Body = "Olï¿½ " & Cells(linha, 3).Value & "," & Chr(10) & Chr(10) _
        & "Poderiam verificar o Boleto do CTe " & Cells(linha, 5).Value & " no valor de " & moeda & "?" & Chr(10) & Chr(10) _
        & "Pois venceu dia: " & Cells(linha, 7).Value & " e nï¿½o identificamos o pagamento." & Chr(10) & Chr(10) _
        & Chr(10) & Chr(10) _
        & " Peï¿½o a devida atenï¿½ï¿½o para que o tï¿½tulo nï¿½o seja enviado a cartï¿½rio." & Chr(10) & Chr(10) _
        & Chr(10) & Chr(10) _
        & "Fico a disposiï¿½ï¿½o," & Chr(10) & Chr(10) _
        & " Nome da Empresa" & Chr(10) & Chr(10) _
        & "fone: 01 1234 5678 Ramal 3022" & Chr(10) & Chr(10) _
        & "whats: 01 123456789  "
    
        Email.Attachments.Add (ThisWorkbook.Path & "\" & Cells(linha, 4).Value & "-" & Cells(linha, 5).Value & "-" & Cells(linha, 6).Value & "(DocumentoEletronico).pdf")
        Email.Attachments.Add (ThisWorkbook.Path & "\" & "Boleto-" & Cells(linha, 5).Value & ".pdf")
    
    
        Email.send
    End If
        
Next


End Sub



