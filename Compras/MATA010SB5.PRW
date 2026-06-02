#include 'Protheus.ch'
/*
//-------------------------------------------------------------------
{Protheus.doc} U_GravaSB5()
-Inclui informações adicionais na tabela SB5 a partir da importação da SB1

@author Helder Vieira Medeiros
@since 27/05/2026
@version 1.0
//-------------------------------------------------------------------
*/
User Function GravaSB5()
    Local aArea   := GetArea()
    Local cCod    := ""
    Local cCeme   := ""
    Local cNbs    := ""

    // 1. Pega o CÓDIGO DO PRODUTO direto da variável do MILE (Linha atual do TXT)
    If Type("M->CB1COD") <> "U"
        cCod := M->CB1COD
    EndIf

    // 2. Lê as outras variáveis do MILE
    If Type("M->CB5CEME") <> "U"
        cCeme := M->CB5CEME
    EndIf
    
    If Type("M->CB5NBS") <> "U"
        cNbs := M->CB5NBS
    EndIf

    // 3. Só grava se tiver código de produto preenchido
    If !Empty(cCod)
        DbSelectArea("SB5")
        DbSetOrder(1) // B5_FILIAL + B5_COD
        
        // Se a SB5 para este produto ainda não existir, cria uma nova. Caso contrário, altera a existente.
        If !DbSeek(xFilial("SB5") + cCod)
            RecLock("SB5", .T.)
            SB5->B5_FILIAL := xFilial("SB5")
            SB5->B5_COD    := cCod
        Else
            RecLock("SB5", .F.)
        EndIf
        
        // Grava as informações nas colunas adicionais
        SB5->B5_CEME := cCeme
        SB5->B5_NBS  := cNbs
        
        MsUnlock()
    EndIf
    
    RestArea(aArea)
Return .T.
