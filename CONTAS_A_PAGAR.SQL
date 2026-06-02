/*
@author Helder Vieira Medeiros
@since 10/12/2025
@version 1.0	
Exemplo de relatório unificado de títulos a pagar para SmartView para Grupo de empresas
*/
SELECT
    'Emp01' AS empresa,
    e.e2_num as numero_Tit,
    e.e2_parcela as parcela,
    e.e2_filial as filial,
    RTRIM(e.e2_nomfor) as fornecedor,
    e.e2_valor as valor,
    e.e2_saldo as saldo,
    e.e2_saldo + e.e2_acresc - e.E2_DECRESC as Total,
    e.e2_naturez as natureza,
    e.e2_tipo as tipo,
    RTRIM(sed.ed_descric) AS desc_natureza,
	CAST(E.E2_EMISSAO AS DATE) AS emissao,   --Tratamento para formatação de data na pltaforma do smartview. O convencional 
	CAST(E.E2_VENCTO AS DATE) AS vencimento, --TO_CHAR(TO_DATE(E.E2_EMISSAO, 'YYYYMMDD'), 'DD/MM/YYYY') AS EMISSAO,
	CAST(E.E2_VENCREA AS DATE) AS venc_real, --não é reconhecido dentro da plataforma, a conversão para data é feita aqui
    e.e2_xfluig as num_fluig,				 -- e a formatação 'DD/MM/YYYY' dentro da plataforma
    e.e2_ccusto as centro_custo,
    e.e2_xdesccc as descricao_centro_custo
FROM
    se2010 e
    LEFT JOIN sed010 sed ON sed.ed_codigo = e.e2_naturez
    AND sed.d_e_l_e_t_ = ' '
WHERE
    e.d_e_l_e_t_ = ' '
    AND e.e2_saldo <> 0
    AND e.e2_tipo <> 'RE'
UNION
ALL
SELECT
    'Emp02' AS empresa,
    e.e2_num as numero_Tit,
    e.e2_parcela as parcela,
    e.e2_filial as filial,
    RTRIM(e.e2_nomfor) as fornecedor,
    e.e2_valor as valor,
    e.e2_saldo as saldo,
    e.e2_saldo + e.e2_acresc - e.E2_DECRESC as Total,
    e.e2_naturez as natureza,
    e.e2_tipo as tipo,
    RTRIM(sed.ed_descric) AS desc_natureza,
	CAST(E.E2_EMISSAO AS DATE) AS emissao,   --Tratamento para formatação de data na pltaforma do smartview. O convencional 
	CAST(E.E2_VENCTO AS DATE) AS vencimento, --TO_CHAR(TO_DATE(E.E2_EMISSAO, 'YYYYMMDD'), 'DD/MM/YYYY') AS EMISSAO,
	CAST(E.E2_VENCREA AS DATE) AS venc_real, --não é reconhecido dentro da plataforma, a conversão para data é feita aqui
    e.e2_xfluig as num_fluig,				 -- e a formatação 'DD/MM/YYYY' dentro da plataforma
    e.e2_ccusto as centro_custo,
    e.e2_xdesccc as descricao_centro_custo
FROM
    se2020 e
    LEFT JOIN sed020 sed ON sed.ed_codigo = e.e2_naturez
    AND sed.d_e_l_e_t_ = ' '
WHERE
    e.d_e_l_e_t_ = ' '
    AND e.e2_saldo <> 0
    AND e.e2_tipo <> 'RE'
UNION
ALL
SELECT
    'Emp03' AS empresa,
    e.e2_num as numero_Tit,
    e.e2_parcela as parcela,
    e.e2_filial as filial,
    RTRIM(e.e2_nomfor) as fornecedor,
    e.e2_valor as valor,
    e.e2_saldo as saldo,
    e.e2_saldo + e.e2_acresc - e.E2_DECRESC as Total,
    e.e2_naturez as natureza,
    e.e2_tipo as tipo,
    RTRIM(sed.ed_descric) AS desc_natureza,
	CAST(E.E2_EMISSAO AS DATE) AS emissao,   --Tratamento para formatação de data na pltaforma do smartview. O convencional 
	CAST(E.E2_VENCTO AS DATE) AS vencimento, --TO_CHAR(TO_DATE(E.E2_EMISSAO, 'YYYYMMDD'), 'DD/MM/YYYY') AS EMISSAO,
	CAST(E.E2_VENCREA AS DATE) AS venc_real, --não é reconhecido dentro da plataforma, a conversão para data é feita aqui
    e.e2_xfluig as num_fluig,				 -- e a formatação 'DD/MM/YYYY' dentro da plataforma
    e.e2_ccusto as centro_custo,
    e.e2_xdesccc as descricao_centro_custo
FROM
    se2030 e
    LEFT JOIN sed030 sed ON sed.ed_codigo = e.e2_naturez
    AND sed.d_e_l_e_t_ = ' '
WHERE
    e.d_e_l_e_t_ = ' '
    AND e.e2_saldo <> 0
    AND e.e2_tipo <> 'RE'
UNION
ALL
SELECT
    'Emo04' AS empresa,
    e.e2_num as numero_Tit,
    e.e2_parcela as parcela,
    e.e2_filial as filial,
    RTRIM(e.e2_nomfor) as fornecedor,
    e.e2_valor as valor,
    e.e2_saldo as saldo,
    e.e2_saldo + e.e2_acresc - e.E2_DECRESC as Total,
    e.e2_naturez as natureza,
    e.e2_tipo as tipo,
    RTRIM(sed.ed_descric) AS desc_natureza,
	CAST(E.E2_EMISSAO AS DATE) AS emissao,   --Tratamento para formatação de data na pltaforma do smartview. O convencional 
	CAST(E.E2_VENCTO AS DATE) AS vencimento, --TO_CHAR(TO_DATE(E.E2_EMISSAO, 'YYYYMMDD'), 'DD/MM/YYYY') AS EMISSAO,
	CAST(E.E2_VENCREA AS DATE) AS venc_real, --não é reconhecido dentro da plataforma, a conversão para data é feita aqui
    e.e2_xfluig as num_fluig,				 -- e a formatação 'DD/MM/YYYY' dentro da plataforma
    e.e2_ccusto as centro_custo,
    e.e2_xdesccc as descricao_centro_custo
FROM
    se2040 e
    LEFT JOIN sed040 sed ON sed.ed_codigo = e.e2_naturez
    AND sed.d_e_l_e_t_ = ' '
WHERE
    e.d_e_l_e_t_ = ' '
    AND e.e2_saldo <> 0
    AND e.e2_tipo <> 'RE'
ORDER BY
    Emissao DESC
