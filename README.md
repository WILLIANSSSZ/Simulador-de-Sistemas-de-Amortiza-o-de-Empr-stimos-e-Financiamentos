# Simulador-de-Sistemas-de-Amortiza-o-de-Empr-stimos-e-Financiamentos
Sistema criado  em VBA exel .
Sub carregar()

With ListView1

If Me.TXT_TIPO.Text = "SAC" Then

                            .Gridlines = True
                            .View = lvwReport
                            .FullRowSelect = True
                            .MultiSelect = True
                            .ColumnHeaders.Add Text:="Periodo", Width:=80, Alignment:=0
                            .ColumnHeaders.Add Text:="PARCELA", Width:=130, Alignment:=1
                            .ColumnHeaders.Add Text:="Amortização", Width:=130, Alignment:=1
                            .ColumnHeaders.Add Text:="juros", Width:=130, Alignment:=1
                            .ColumnHeaders.Add Text:="Saldo devedor", Width:=130, Alignment:=1
            End If
            
      If Me.TXT_TIPO.Text = "PRICE" Then
      
            
         
                                  .Gridlines = True
                                .View = lvwReport
                                .FullRowSelect = True
                                .MultiSelect = True
                                .ColumnHeaders.Add Text:="Periodo", Width:=80, Alignment:=0
                                .ColumnHeaders.Add Text:="PRESTAÇÃO", Width:=130, Alignment:=1
                                .ColumnHeaders.Add Text:="Amortização", Width:=130, Alignment:=1
                                .ColumnHeaders.Add Text:="juros", Width:=130, Alignment:=1
                                .ColumnHeaders.Add Text:="Saldo devedor", Width:=130, Alignment:=1
                                
            End If
            
End With


Dim linha As Double
Dim lista As Object
linha = 14

Me.ListView1.ListItems.Clear


With Plan1

    While .Cells(linha, 3).Value <> Empty
                 
                 
                  With Me.ListView1
        
                                              
                            Set lista = Me.ListView1.ListItems.Add(Text:=Plan1.Cells(linha, 3).Value)
                            
                                    lista.ListSubItems.Add Text:=Plan1.Cells(linha, 4).Text
                                    lista.ListSubItems.Add Text:=Plan1.Cells(linha, 5).Text
                                     lista.ListSubItems.Add Text:=Plan1.Cells(linha, 6).Text
                                      lista.ListSubItems.Add Text:=Plan1.Cells(linha, 7).Text
                        End With
                       
                        
            linha = linha + 1

Wend
End With

Set lista = Nothing


End Sub
