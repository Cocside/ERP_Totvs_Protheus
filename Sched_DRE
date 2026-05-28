#Include 'Protheus.ch'
#Include "Totvs.ch"
#Include "FWMVCDef.ch"
/*/
//-------------------------------------------------------------------
{Protheus.doc} U_GravaSB5()
-Schedule para gerar a planilha de DRE

@author Helder Vieira Medeiros
@since 27/05/2026
@version 1.0
//-------------------------------------------------------------------
/*/
Function U_SCHDRE()
	Local cQuery    := ""
	Local cPath     := " " //<-caminho da pasta
	Local cFileName := "Relatorio_DRE_" + cValToChar(Year(Date())) + ".xlsx"
	Local cAno := cValToChar(Year(Date()))
	Local oExcel    := FWMSExcelXLSX():New()
	Local cAliasQ   := GetNextAlias()
	Local cTitulo := "DRE_DATA"
	Local cWorkSheet := "DRE"

	// inicializa o ambiente
	RpcSetEnv("25", "01")  //empresa e Filial

	// Verifica se o diretório existe, se não, cria
	If !ExistDir(cPath)
		MakeDir(cPath)
	EndIf

//inicio da query

cQuery := ""
cQuery += " SELECT "
cQuery += "     T_FIN.ANO, "
cQuery += "     T_FIN.MES, "
cQuery += "     T_FIN.LINHA_DRE, "
cQuery += "     T_FIN.CONTA_DETALHE, "
cQuery += "     T_FIN.VALOR "
cQuery += " FROM ( "
cQuery += "     SELECT "
cQuery += "         T_RES.ANO, "
cQuery += "         T_MAP.MES_DESTINO AS MES, "

cQuery += "         CASE "
cQuery += "             WHEN T_FRM.ORD = 1  THEN '01. RECEITA BRUTA' "
cQuery += "             WHEN T_FRM.ORD = 2  THEN '02. (-) DEDUCOES' "
cQuery += "             WHEN T_FRM.ORD = 3  THEN '03. (=) RECEITA LIQUIDA' "
cQuery += "             WHEN T_FRM.ORD = 4  THEN '04. (-) CUSTOS' "
cQuery += "             WHEN T_FRM.ORD = 5  THEN '05. (=) LUCRO BRUTO' "
cQuery += "             WHEN T_FRM.ORD = 6  THEN '06. (-) DESPESAS OPERACIONAIS' "
cQuery += "             WHEN T_FRM.ORD = 7  THEN '07. (-) DESPESAS ADMINISTRATIVAS' "
cQuery += "             WHEN T_FRM.ORD = 8  THEN '08. (-) DEPRECIACAO' "
cQuery += "             WHEN T_FRM.ORD = 9  THEN '09. (-) TRIBUTOS' "
cQuery += "             WHEN T_FRM.ORD = 10 THEN '10. (+) RESULTADO EXTRAORDINARIO' "
cQuery += "             WHEN T_FRM.ORD = 11 THEN '11. (-) DESPESAS FINANCEIRAS' "
cQuery += "             WHEN T_FRM.ORD = 12 THEN '12. (+) RECEITAS FINANCEIRAS' "
cQuery += "             WHEN T_FRM.ORD = 13 THEN '13. (=) RESULTADO FINANCEIRO' "
cQuery += "             WHEN T_FRM.ORD = 14 THEN '14. (=) RESULTADO OPERACIONAL' "
cQuery += "             WHEN T_FRM.ORD = 15 THEN '15. (-) IMPOSTOS' "
cQuery += "             WHEN T_FRM.ORD = 16 THEN '16. (=) RESULTADO LIQUIDO' "
cQuery += "         END AS LINHA_DRE, "

cQuery += "         CASE "
cQuery += "             WHEN T_FRM.ORD IN (1,2,4,6,7,8,9,10,11,12,15) "
cQuery += "             THEN T_RES.CONTA_DETALHE "
cQuery += "             ELSE 'TOTAL DO GRUPO' "
cQuery += "         END AS CONTA_DETALHE, "

cQuery += "         CASE "
cQuery += "             WHEN T_FRM.ORD = 1  THEN T_RES.V_RECEITA "
cQuery += "             WHEN T_FRM.ORD = 2  THEN T_RES.V_DEDUCAO "
cQuery += "             WHEN T_FRM.ORD = 3  THEN (T_RES.V_RECEITA + T_RES.V_DEDUCAO) "
cQuery += "             WHEN T_FRM.ORD = 4  THEN T_RES.V_CUSTO "
cQuery += "             WHEN T_FRM.ORD = 5  THEN (T_RES.V_RECEITA + T_RES.V_DEDUCAO + T_RES.V_CUSTO) "
cQuery += "             WHEN T_FRM.ORD = 6  THEN T_RES.V_DESP_OPER "
cQuery += "             WHEN T_FRM.ORD = 7  THEN T_RES.V_DESP_ADMIN "
cQuery += "             WHEN T_FRM.ORD = 8  THEN T_RES.V_DEPRECIACAO "
cQuery += "             WHEN T_FRM.ORD = 9  THEN T_RES.V_TRIBUTOS "
cQuery += "             WHEN T_FRM.ORD = 10 THEN T_RES.V_EXTRA "
cQuery += "             WHEN T_FRM.ORD = 11 THEN T_RES.V_FIN_DESP "
cQuery += "             WHEN T_FRM.ORD = 12 THEN T_RES.V_FIN_REC "
cQuery += "             WHEN T_FRM.ORD = 13 THEN (T_RES.V_FIN_DESP + T_RES.V_FIN_REC) "
cQuery += "             WHEN T_FRM.ORD = 14 THEN ( "
cQuery += "                 T_RES.V_RECEITA + "
cQuery += "                 T_RES.V_DEDUCAO + "
cQuery += "                 T_RES.V_CUSTO + "
cQuery += "                 T_RES.V_DESP_OPER + "
cQuery += "                 T_RES.V_DESP_ADMIN + "
cQuery += "                 T_RES.V_DEPRECIACAO + "
cQuery += "                 T_RES.V_TRIBUTOS + "
cQuery += "                 T_RES.V_EXTRA + "
cQuery += "                 T_RES.V_FIN_DESP + "
cQuery += "                 T_RES.V_FIN_REC "
cQuery += "             ) "
cQuery += "             WHEN T_FRM.ORD = 15 THEN T_RES.V_IMPOSTOS "
cQuery += "             WHEN T_FRM.ORD = 16 THEN ( "
cQuery += "                 T_RES.V_RECEITA + "
cQuery += "                 T_RES.V_DEDUCAO + "
cQuery += "                 T_RES.V_CUSTO + "
cQuery += "                 T_RES.V_DESP_OPER + "
cQuery += "                 T_RES.V_DESP_ADMIN + "
cQuery += "                 T_RES.V_DEPRECIACAO + "
cQuery += "                 T_RES.V_TRIBUTOS + "
cQuery += "                 T_RES.V_EXTRA + "
cQuery += "                 T_RES.V_FIN_DESP + "
cQuery += "                 T_RES.V_FIN_REC + "
cQuery += "                 T_RES.V_IMPOSTOS "
cQuery += "             ) "
cQuery += "         END AS VALOR, "

cQuery += "         T_FRM.ORD AS ORDEM_LINHA "

cQuery += "     FROM ( "

cQuery += "         SELECT '01' AS MES_ORIGEM, '01 - JAN' AS MES_DESTINO FROM DUAL UNION ALL "
cQuery += "         SELECT '02', '02 - FEV' FROM DUAL UNION ALL "
cQuery += "         SELECT '03', '03 - MAR' FROM DUAL UNION ALL "
cQuery += "         SELECT '04', '04 - ABR' FROM DUAL UNION ALL "
cQuery += "         SELECT '05', '05 - MAI' FROM DUAL UNION ALL "
cQuery += "         SELECT '06', '06 - JUN' FROM DUAL UNION ALL "
cQuery += "         SELECT '07', '07 - JUL' FROM DUAL UNION ALL "
cQuery += "         SELECT '08', '08 - AGO' FROM DUAL UNION ALL "
cQuery += "         SELECT '09', '09 - SET' FROM DUAL UNION ALL "
cQuery += "         SELECT '10', '10 - OUT' FROM DUAL UNION ALL "
cQuery += "         SELECT '11', '11 - NOV' FROM DUAL UNION ALL "
cQuery += "         SELECT '12', '12 - DEZ' FROM DUAL "

cQuery += "     ) T_MAP "

cQuery += "     CROSS JOIN ( "
cQuery += "         SELECT LEVEL AS ORD "
cQuery += "         FROM DUAL "
cQuery += "         CONNECT BY LEVEL <= 16 "
cQuery += "     ) T_FRM "

cQuery += "     INNER JOIN ( "

cQuery += "         SELECT "
cQuery += "             T_TMP.ANO, "
cQuery += "             T_TMP.MES, "
cQuery += "             T_TMP.CONTA_DETALHE, "

cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'RECEITA'      THEN T_TMP.SALDO ELSE 0 END) AS V_RECEITA, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'DEDUCAO'      THEN T_TMP.SALDO ELSE 0 END) AS V_DEDUCAO, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'CUSTO'        THEN T_TMP.SALDO ELSE 0 END) AS V_CUSTO, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'OPERACIONAL'  THEN T_TMP.SALDO ELSE 0 END) AS V_DESP_OPER, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'ADMIN'        THEN T_TMP.SALDO ELSE 0 END) AS V_DESP_ADMIN, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'DEPRECIACAO' THEN T_TMP.SALDO ELSE 0 END) AS V_DEPRECIACAO, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'TRIBUTOS'     THEN T_TMP.SALDO ELSE 0 END) AS V_TRIBUTOS, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'EXTRA'        THEN T_TMP.SALDO ELSE 0 END) AS V_EXTRA, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'FIN_DESP'     THEN T_TMP.SALDO ELSE 0 END) AS V_FIN_DESP, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'FIN_REC'      THEN T_TMP.SALDO ELSE 0 END) AS V_FIN_REC, "
cQuery += "             SUM(CASE WHEN T_TMP.GRUPO = 'IMPOSTOS'     THEN T_TMP.SALDO ELSE 0 END) AS V_IMPOSTOS "

cQuery += "         FROM ( "

cQuery += "             SELECT "
cQuery += "                 SUBSTR(LANC.DATA_REF,1,4) AS ANO, "
cQuery += "                 SUBSTR(LANC.DATA_REF,5,2) AS MES, "
cQuery += "                 MAP.GRUPO, "
cQuery += "                 TRIM(LANC.CONTA) AS CONTA_CODIGO, "
cQuery += "                 TRIM(LANC.CONTA) || ' - ' || MAX(CTA.DESCRICAO) AS CONTA_DETALHE, "
cQuery += "                 SUM(LANC.CREDITO - LANC.DEBITO) AS SALDO "

cQuery += "             FROM TB_LANCAMENTOS LANC "

cQuery += "             LEFT JOIN TB_MAPEAMENTO MAP "
cQuery += "                 ON TRIM(LANC.CONTA) BETWEEN MAP.CONTA_INI AND MAP.CONTA_FIM "

cQuery += "             LEFT JOIN TB_CONTAS CTA "
cQuery += "                 ON CTA.CONTA = LANC.CONTA "

cQuery += "             WHERE LANC.FLAG_EXCLUSAO = 'N' "

cQuery += "             GROUP BY "
cQuery += "                 SUBSTR(LANC.DATA_REF,1,4), "
cQuery += "                 SUBSTR(LANC.DATA_REF,5,2), "
cQuery += "                 MAP.GRUPO, "
cQuery += "                 LANC.CONTA "

cQuery += "         ) T_TMP "

cQuery += "         GROUP BY "
cQuery += "             T_TMP.ANO, "
cQuery += "             T_TMP.MES, "
cQuery += "             T_TMP.CONTA_DETALHE "

cQuery += "     ) T_RES "
cQuery += "         ON T_RES.MES = T_MAP.MES_ORIGEM "

cQuery += " ) T_FIN "

cQuery += " WHERE T_FIN.VALOR <> 0 "

cQuery += " ORDER BY "
cQuery += "     T_FIN.ANO DESC, "
cQuery += "     T_FIN.MES, "
cQuery += "     T_FIN.ORDEM_LINHA, "
cQuery += "     T_FIN.CONTA_DETALHE "
//fim da query

	//sem ChangeQuery
	DbUseArea(.T., "TOPCONN", TcGenQry(,,cQuery), cAliasQ, .T., .T.)

//configuração do Excel
	//AddworkSheet(variavel_aba)
	oExcel:AddworkSheet(cWorkSheet)

	//AddColumn(Nome_da_aba, Nome_da_Tabela, Titulo, Alinhamento, Tipo, Negrito)
	oExcel:AddColumn(cWorkSheet, cTitulo, "ANO",           1, 1, .T.)
	oExcel:AddColumn(cWorkSheet, cTitulo, "MES",           1, 1, .T.)
	oExcel:AddColumn(cWorkSheet, cTitulo, "LINHA_DRE",     1, 1, .T.)
	oExcel:AddColumn(cWorkSheet, cTitulo, "CONTA_DETALHE", 1, 1, .T.)
	oExcel:AddColumn(cWorkSheet, cTitulo, "VALOR",         3, 2, .T.) // 3=Direita, 4=Numérico 2 casas

	//percorre os registros e adiciona as linhas
	While (cAliasQ)->(!Eof())

		oExcel:AddRow(cWorkSheet, ,{;
			(cAliasQ)->ANO,;
			(cAliasQ)->MES,;
			(cAliasQ)->LINHA_DRE,;
			(cAliasQ)->CONTA_DETALHE,;
			(cAliasQ)->VALOR;
			})
		(cAliasQ)->(DbSkip())
	EndDo

	// Gera o arquivo no servidor
	oExcel:Activate()
	oExcel:GetFile(cPath + cFileName)

	(cAliasQ)->(DbCloseArea())
	FreeObj(oExcel)

	ConOut("[OK] " + Time() + " - Schedule DRE finalizado: " + cPath + cFileName)

Return
