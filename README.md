# Simulador-de-Sistemas-de-Amortiza-o-de-Empr-stimos-e-Financiamentos
Sistema criado  em VBA exel .
Private Sub cb_ano_mes_Change()
Me.lbl_taxa_perce = Me.cb_ano_mes
End Sub




Private Sub Frame2_Click()

End Sub

Private Sub Label22_Click()

Application.Visible = True
Unload simulador
End Sub

Private Sub Label23_Click()
Application.DisplayAlerts = False
Application.Quit

Application.DisplayAlerts = True

End Sub

Private Sub Label24_Click()

End Sub

Private Sub Label9_Click()
Instru.Show
End Sub

Private Sub lb_saldo_Click()

End Sub

Private Sub lbl_amort_Click()

End Sub

Private Sub LBL_JUROS_Click()

End Sub

Private Sub opt_ano_Click()
If Me.opt_ano.Value = True Then

        Plan1.Range("H2") = "Anual"
         Me.LBL_TAXAAISO = " % a.a"
        Call carregar
        End If
End Sub

Private Sub opt_mes_Click()

If Me.opt_mes.Value = True Then

        Plan1.Range("H2") = "mensal"
      Me.LBL_TAXAAISO = " % a.m"
        Call carregar
        
        End If
End Sub

Private Sub spin_pe_Change()

Call carregar
Me.TXT_P.Value = Me.spin_pe

Me.spin_pe.Max = 30

Me.lb_saldo = Plan1.Range("D11").Text
Me.lbl_amort = Plan1.Range("E11").Text
Me.LBL_JUROS = Plan1.Range("F11").Text

End Sub



Private Sub TabStrip1_Change()

End Sub

Private Sub TXT_P_Change()
Plan1.Range("D7").Value = Me.TXT_P.Value
Call carregar


End Sub


Private Sub TXT_TAXA_Change()
Plan1.Range("D5").Value = Me.TXT_TAXA.Value / 100


Call Me.carregar
End Sub

Private Sub TXT_TIPO_AfterUpdate()

End Sub

Private Sub TXT_TIPO_Change()




Me.ListView1.ListItems.Clear
Me.LBL_SITEMA = Me.TXT_TIPO.Value
Plan1.Range("D9") = Me.TXT_TIPO.Value

Me.lb_saldo = Plan1.Range("D11").Text
Me.lbl_amort = Plan1.Range("E11").Text
Me.LBL_JUROS = Plan1.Range("F11").Text
Call carregar

End Sub

Private Sub TXT_VALOR_FINAN_Change()
Plan1.Range("D3").Value = Me.TXT_VALOR_FINAN.Value


End Sub

Private Sub TXT_VALOR_FINAN_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Me.TXT_VALOR_FINAN.Value = Format(Me.TXT_VALOR_FINAN.Value, "R$ #,##0.00")
End Sub

Private Sub TXT_VALOR_FINAN_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Me.lb_saldo = Plan1.Range("D11").Text
Me.lbl_amort = Plan1.Range("E11").Text
Me.LBL_JUROS = Plan1.Range("F11").Text
End Sub

Private Sub UserForm_Activate()
Call carregar
Me.ListView1.ListItems.Clear


If Me.opt_mes.Value = True Then

        Plan1.Range("H2") = "mensal"

ElseIf Me.opt_ano = True Then

            Plan1.Range("H2") = "Anual"

End If

End Sub


Private Sub UserForm_Initialize()
 

Me.ListView1.ListItems.Clear

Dim TXT_TAXA As Double
Dim TXT_TIPO As String
Dim TXT_VALOR_FINAN As Double
Dim TXT_P As Integer

Me.TXT_P.Text = ""
Me.TXT_TAXA.Value = ""
Me.TXT_TIPO.Text = ""
Me.TXT_VALOR_FINAN.Value = ""
Me.LBL_SITEMA = ""
Me.TXT_P.Value = ""

'SELECIONANDO O SISTEMA DE AMORTIZAÇÃO



Me.TXT_TIPO.AddItem ("PRICE")
Me.TXT_TIPO.AddItem ("SAC")
Me.TXT_TIPO.AddItem ("AMERICANO")
Me.TXT_TIPO.AddItem ("MISTO")
Call carregar

End Sub

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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then Cancel = True
