Option Explicit
Public z           As Integer  'Última linha

Sub classref()

Dim SessObj     As Object   'Para acessar o obj NFC
Dim S           As String   'Sessão do NFC
Dim B           As Integer  'só uma confirmação, pois quero dar prioridade a origem de uma RU1 Brasil
Dim C           As Integer  'Para clicar na página seguinte
Dim cel         As Range    'Para utilizar o for
Dim REF         As String   'Referencia a ser buscada
Dim pag1        As Integer  'pagina atual NFC
Dim pag2        As Integer  'número de paginas no NFC
Dim org         As String   'Origem peça final
Dim org1        As String   'Origem da peça quando Brasil e RU1
Dim org2        As String   'Origem da peça quando CKD (RU2)
Dim org3        As String   'Origem da peça quando CKD (RU1)
Dim utl         As String   'Onde é utilizada
Dim dt1         As Date     'Data de início
Dim dt2         As Variant  'Data de fim
Dim cls         As String   'Classe da REF
Dim l           As Integer  'para percorrer as linhas do NFC
Dim l2          As Integer  'contador para buscar a origem acima da linha de utilização
Dim inex        As String   'Para pegar variáveis inexistentes
Dim RU1         As Integer  'Apenas para classificarmos dentro do if
Dim RU2         As Integer  'Apenas para classificarmos dentro do if (prioridade)

'Lógica de classificação:
'Guardamos a referência na var REF, ativamos o NFC e damos enter no local certo de busca
'Verificamos se é uma REF inexistente
'Verificamos quantas páginas a REF gerou no NFC
'Depois percorremos todas essas páginas na coluna X buscando utilização Brasil
'Se acharmos Brasil, verificamos a coluna de data com a var "dt1"
'Se dt1 for menor que a data atual, verificamos a var "dt2"
'Se dt2 for maior que a data atual, esta REF será  RU4
'Depois é verificado na linha acima, coluna Y, a origem e guardamos na variável "org"

S = UCase(InputBox("Qual sessão do Emulation 3270 - NFC você está usando? A ou C?", "Sessão NFC"))

Application.ScreenUpdating = False

If S <> "C" Or S <> "A" Then
    Exit Sub
End If

z = Range("A1048576").End(xlUp).Row

For Each cel In Range("A2:A" & z)

    If cel.Value Like "??????????" Then 'Para ser classificada deve haver 10 caracteres
    
        REF = cel.Value
        cls = Empty
        org = Empty
        org1 = Empty
        org2 = Empty
        org3 = Empty
        B = 0
        RU1 = 0
        RU2 = 0
        
        Set SessObj = CreateObject("PCOMM.autECLSession")
         
        SessObj.SetConnectionByName (S)
        
        SessObj.autECLPS.SendKeys REF
        SessObj.autECLPS.Wait 750
        SessObj.autECLPS.SendKeys "[enter]"
        SessObj.autECLPS.Wait 750
        
        inex = SessObj.autECLPS.GetText(24, 2, 35)
        
        SessObj.autECLOIA.WaitForInputReady

        If inex = "NUMERO DE PRODUIT INEXISTANT EN NFC" Then
            cls = "Inexistente"
            org = ""
            GoTo saida1
        End If
        
        pag1 = 1
        pag2 = SessObj.autECLPS.GetText(5, 76, 2)
        
        
        C = 0
        
        Do While pag1 <= pag2   'Buscando em todas as páginas
            
            l = 11
            
            If C <> 0 Then
                SessObj.autECLPS.SendKeys "[pf8]"
            End If
            
            Do While l <= 18    'Buscando em todas as linhas
            
                utl = SessObj.autECLPS.GetText(l, 10, 3)
                l2 = 1
                
                If utl Like "78?" Then   'Buscando apenas Brasil (78) no NFC
                
                    dt1 = SessObj.autECLPS.GetText(l, 56, 10)
                    dt2 = SessObj.autECLPS.GetText(l, 69, 10)
                    
                    If dt1 <= Date Then
                        If dt2 > Date Or dt2 = 0 Then
                        
                            cls = "RU4"
                            Do While org = Empty
                            org = Trim(SessObj.autECLPS.GetText(l - l2, 4, 5))
                            l2 = l2 + 1
                            Loop
                            GoTo saida1 'Chegando aqui, a REF já está classificada e saimos fora
                        
                        End If
                    ElseIf dt1 > Date Then
                    
                        RU1 = 1
                        B = 1
                        Do While org1 = Empty
                        org1 = Trim(SessObj.autECLPS.GetText(l - l2, 4, 5))
                        l2 = l2 + 1
                        Loop
                        
                    End If
                              
                ElseIf utl Like "???" And utl <> "   " Then
                
                    dt1 = SessObj.autECLPS.GetText(l, 56, 10)
                    dt2 = SessObj.autECLPS.GetText(l, 69, 10)
                    
                    If dt1 <= Date Then
                        If dt2 > Date Or dt2 = 0 Then
                        
                            RU2 = 1
                            Do While org2 = Empty
                            org2 = Trim(SessObj.autECLPS.GetText(l - l2, 4, 5))
                            l2 = l2 + 1
                            Loop
                                                
                        End If
                        
                    ElseIf dt1 > Date Then
                    
                        RU1 = 1
                        Do While org3 = Empty
                        org3 = Trim(SessObj.autECLPS.GetText(l - l2, 4, 5))
                        l2 = l2 + 1
                        Loop
                    
                    End If
                
                End If
                
                l = l + 1
                C = C + 1
                            
            Loop
            pag1 = pag1 + 1
        Loop
        
        If RU2 = 1 Then
            cls = "RU2"
            org = org2
        ElseIf RU2 = 0 And RU1 = 1 And B = 1 Then
            cls = "RU1"
            org = org1
        ElseIf RU2 = 0 And RU1 = 1 And B = 0 Then
            cls = "RU1"
            org = org3
        End If
        
saida1:
    
        
        cel.Offset(0, 1).Value = org
        cel.Offset(0, 2).Value = cls
        
        If cls <> "Inexistente" Then
            SessObj.autECLPS.SendKeys "[tab]"
            SessObj.autECLPS.SendKeys "[tab]"
            SessObj.autECLPS.SendKeys "[tab]"
            SessObj.autECLPS.SendKeys "[tab]"
            SessObj.autECLOIA.WaitForInputReady
        End If
    
    End If
       
Next

Application.ScreenUpdating = True
MsgBox "Classificação realizada!"

End Sub


