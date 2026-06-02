#include 'Totvs.ch'
//------------------------------------------------------------------
/*/{Protheus.doc} MT410TOK & MRatLOk
VALIDA A INCLUSÃO DO PEDIDO DE VENDA | VALIDA O PREENCHIMENTO DE RATEIO DENTRO DO PEDIDO DE VENDA PARA TES 509 E 513

@author Helder Vieira Medeiros
@since 24/04/2026
@version 1.0
/*/
//------------------------------------------------------------------
User Function MT410TOK()
    Local lRet      := .T.
    Local nI        := 0
    Local nPosTES   := GdFieldPos("C6_TES")
    
    // Verificamos se houve rateio para os itens que precisam
    For nI := 1 To Len(aCols)
        If !aCols[nI][Len(aCols[nI])]
            If aCols[nI][nPosTES] $ "509/513"
                
                // Se a variável não existe ou a soma gravada for < 100
                If Type("nSomaJall") == "U" .Or. nSomaJall < 100
                    MsgStop("O rateio deste pedido não foi realizado ou está incompleto (0%)." + CRLF + "Acesse 'Outras Ações -> Rateio'.", "JallCard")
                    lRet := .F.
                    Exit
                EndIf
                
            EndIf
        EndIf
    Next nI

    // O RESET: Se a validação passou e o pedido vai salvar, limpa a variável
    // para o próximo pedido começa "zerado"
    If lRet .And. Type("nSomaJall") != "U"
        nSomaJall := 0 
    EndIf

Return lRet

User Function MRatLOk()
    Local nX        := 0
    Local nPosPerc  := 0
    Local nSoma     := 0
    
    // 1. Calcula a soma do que está na tela do rateio
    If Type("aHeader") == "A" .And. Type("aCols") == "A"
        nPosPerc := aScan(aHeader, {|x| AllTrim(x[2]) == "AGG_PERC"})
        If nPosPerc > 0
            For nX := 1 To Len(aCols)
                If !aCols[nX][Len(aCols[nX])]
                    nSoma += aCols[nX][nPosPerc]
                EndIf
            Next nX
        EndIf
    EndIf

    // 2. "Manda" para o MT410TOK via variável de memória
    // Se a variável já existe, atualiza. Se não, cria como Private.
    If Type("nSomaJall") == "U"
        Public nSomaJall := nSoma // Usamos Public aqui para garantir que atravesse fontes
    Else
        nSomaJall := nSoma
    EndIf

Return .T.
