#Include 'Protheus.ch'
#Include 'TbiConn.ch'
#Include 'TryException.ch'
#DEFINE CRLF (Chr(13)+Chr(10))

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ ORCDESPR บAutor  ณEvandro Cleto       บ Data ณ  18/02/22   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Relatorio Excel para Or็amento e Despesa.                  บฑฑ
ฑฑบ          ณ                                                            บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Meu mesmo.....                                             บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function ORCDESPR() 
	Processa({|| GerRExcel()},"Gerando Rel. Receitas e Despesas")
Return

*****************************************************************************
Static Function GerRExcel()
	Local nCnt 	  := 0
	Local nCount  := 0
	Local aStruQry  := {}
	Local aStruQRc  := {}
	Local cArquivo  := ""
	Local oExcelApp := Nil
	Local cPath     := "C:\Users\Evandro\Documents\Financeiro\2022"
	Local nRet      := 0
	Local aLinExcel := {}
	Local oBrush1
	Local oExcel
	Local oExcelApp
	Local oException
	Local cAba1 
	Local cTabela1
	Local nXa := 0
	Local nX  := 0
	Local nY  := 0
	Local nValDesp := 0
	Local nValRec := 0
	Local nTPerdesp := 0
	Local nTPerRec  := 0
	Local nPerdesp  := 0
	Local nPerRec   := 0
	Private cMes := ""
	Private nTotReg := 0
	Private nPoslin := 0 
	Private _aResp  := {}	
    Private _cQuery   := ""
	Private  cQry     := ""
  	Private cMesSE7   := ""
	
	if MyPerg(@_aResp,.T.)
	
	do case
    	case Alltrim(_aResp[2]) == '1'
			cMesSE7   := 'E7_VALJAN1'
    	case Alltrim(_aResp[2]) == '2'
			cMesSE7   := 'E7_VALFEV1'
    	case Alltrim(_aResp[2]) == '3'
			cMesSE7   := 'E7_VALMAR1'		
    	case Alltrim(_aResp[2]) == '4'
			cMesSE7   := 'E7_VALABR1'		
    	case Alltrim(_aResp[2]) == '5'
			cMesSE7   := 'E7_VALMAI1'		
    	case Alltrim(_aResp[2]) == '6'
			cMesSE7   := 'E7_VALJUN1'		
    	case Alltrim(_aResp[2]) == '7'
			cMesSE7   := 'E7_VALJUL1'
    	case Alltrim(_aResp[2]) == '8'
			cMesSE7   := 'E7_VALAGO1'
    	case Alltrim(_aResp[2]) == '9'
			cMesSE7   := 'E7_VALSET1'
    	case Alltrim(_aResp[2]) == '10'
			cMesSE7   := 'E7_VALOUT1'
		case Alltrim(_aResp[2]) == '11'
			cMesSE7   := 'E7_VALNOV1'
    	case Alltrim(_aResp[2]) == '12'
			cMesSE7   := 'E7_VALDEZ1'
	endcase

	//CALCULO DE DESPESA
	/*
    _cQuery := "SELECT YEAR(E5_DATA) as ANO,                                                   " + CRLF
	_cQuery += "CASE MONTH(E5_DATA)                                                         " + CRLF
	_cQuery += " WHEN '1' THEN 'JANEIRO'                                                    " + CRLF
	_cQuery += " WHEN '2' THEN 'FEVEREIRO'                                                  " + CRLF
	_cQuery += " WHEN '3' THEN 'MARCO'                                                      " + CRLF
	_cQuery += " WHEN '4' THEN 'ABRIL'                                                      " + CRLF
	_cQuery += " WHEN '5' THEN 'MAIO'                                                       " + CRLF
	_cQuery += " WHEN '6' THEN 'JUNHO'                                                      " + CRLF
	_cQuery += " WHEN '7' THEN 'JULHO'                                                      " + CRLF
	_cQuery += " WHEN '8' THEN 'AGOSTO'                                                     " + CRLF
	_cQuery += " WHEN '9' THEN 'SETEMBRO'                                                   " + CRLF
	_cQuery += " WHEN '10' THEN 'OUTUBRO'                                                   " + CRLF
	_cQuery += " WHEN '11' THEN 'NOVEMBRO'                                                  " + CRLF
	_cQuery += " WHEN '12' THEN 'DEZEMBRO'                                                  " + CRLF
	_cQuery += "END AS MES,                                                                 " + CRLF
	_cQuery += "ED_DESCRIC DESC_NATUREZA,E5_NATUREZ COD_NATUREZA, SUM( E5_VALOR) AS DESPESA, " + CRLF
	_cQuery += ""+cMesSE7+" AS ORCADO,                                                       " + CRLF
	_cQuery += "CASE WHEN ("+cMesSE7+" IS NULL OR "+cMesSE7+" = 0) THEN Round(((SUM( E5_VALOR) / 1)),2)   " + CRLF
	_cQuery += "ELSE Round(((SUM( E5_VALOR) / "+cMesSE7+")*100),2) END AS PERCENTUAL,         " + CRLF
    _cQuery += " Round(("+cMesSE7+" - SUM( E5_VALOR)),2) AS RESULTADO                        " + CRLF
    _cQuery += "FROM       "+RetSqlName("SE5")+" SE5                                        " + CRLF 
    _cQuery += "INNER JOIN "+RetSqlName("SED")+" SED                                        " + CRLF
	_cQuery += "ON E5_NATUREZ = ED_CODIGO AND SE5.D_E_L_E_T_ = ' '                          " + CRLF
    _cQuery += "LEFT JOIN "+RetSqlName("SE7")+" SE7                                        " + CRLF
	_cQuery += "ON E5_NATUREZ = E7_NATUREZ AND YEAR(E5_DATA) = E7_ANO AND SE7.D_E_L_E_T_ = ' '" + CRLF
	_cQuery += "WHERE SE5.E5_FILIAL = '"+xFilial("SE5")+"'                                  " + CRLF
	_cQuery += "AND YEAR(E5_DATA) = '"+Alltrim(_aResp[1])+"'                                " + CRLF
	_cQuery += "AND MONTH(E5_DATA) = '"+Alltrim(_aResp[2])+"'                               " + CRLF
	_cQuery += "AND E5_MOEDA NOT IN ('TB')                                                  " + CRLF
	_cQuery += "AND E5_RECPAG = 'P'                                                         " + CRLF
	_cQuery += "AND E5_SITUACA = ' '                                                        " + CRLF
	_cQuery += "AND E5_TIPODOC NOT IN ('BA','DA','DC')                                      " + CRLF
    _cQuery += "AND E5_LA NOT IN('S')                                                       " + CRLF
	_cQuery += "AND E5_BANCO = '077'                                                        " + CRLF
	_cQuery += "AND SE5.D_E_L_E_T_ = ' '                                                    " + CRLF
	_cQuery += "AND E5_NATUREZ < '09'                                                       " + CRLF
	_cQuery += "GROUP BY YEAR(E5_DATA), MONTH(E5_DATA), ED_DESCRIC, E5_NATUREZ, "+cMesSE7+"             " + CRLF
	_cQuery += "ORDER BY 7 DESC "   
	*/
	_cQuery := "SELECT E7_ANO,                                                                   		                         " + CRLF
	_cQuery += "CASE "+Alltrim(_aResp[2])+"                                                          							 " + CRLF
	_cQuery += " WHEN '1' THEN 'JANEIRO'                                                             							 " + CRLF
	_cQuery += " WHEN '2' THEN 'FEVEREIRO'                                                           							 " + CRLF
	_cQuery += " WHEN '3' THEN 'MARCO'                                                               							 " + CRLF
	_cQuery += " WHEN '4' THEN 'ABRIL'                                                               							 " + CRLF
	_cQuery += " WHEN '5' THEN 'MAIO'                                                                							 " + CRLF
	_cQuery += " WHEN '6' THEN 'JUNHO'                                                               							 " + CRLF
	_cQuery += " WHEN '7' THEN 'JULHO'                                                               							 " + CRLF
	_cQuery += " WHEN '8' THEN 'AGOSTO'                                                     							         " + CRLF
	_cQuery += " WHEN '9' THEN 'SETEMBRO'                                                        							     " + CRLF
	_cQuery += " WHEN '10' THEN 'OUTUBRO'                                                           						     " + CRLF
	_cQuery += " WHEN '11' THEN 'NOVEMBRO'                                                           							 " + CRLF
	_cQuery += " WHEN '12' THEN 'DEZEMBRO'                                                           							 " + CRLF
	_cQuery += "END AS MES,                                                                       								 " + CRLF
	_cQuery += "ED_DESCRIC DESC_NATUREZA, E7_NATUREZ COD_NATUREZA, "+cMesSE7+" ORCADO, ISNULL(TSE2.TITULO_PAGAR,0) TITULO_PAGAR, " + CRLF     
	_cQuery += "ISNULL(TSE5.DESPESA,0) DESPESA, ISNULL(TSE2.TITULO_PAGAR,0) + ISNULL(TSE5.DESPESA,0) TOTAL_GASTO,                " + CRLF     	
	_cQuery += "CASE WHEN ("+cMesSE7+" IS NULL OR "+cMesSE7+" = 0) THEN Round(((ISNULL(TSE2.TITULO_PAGAR,0) + ISNULL(TSE5.DESPESA,0)) / 1),2) " + CRLF
	_cQuery += "ELSE Round(((((ISNULL(TSE2.TITULO_PAGAR,0) + ISNULL(TSE5.DESPESA,0)) / E7_VALMAR1))*100),2) END AS PERCENTUAL,   " + CRLF
	_cQuery += " Round(("+cMesSE7+" - (ISNULL(TSE2.TITULO_PAGAR,0) + ISNULL(TSE5.DESPESA,0))),2) AS RESULTADO 					 " + CRLF
	_cQuery += " FROM "+RetSqlName("SE7")+" SE7                                                      							 " + CRLF
	_cQuery += "INNER JOIN "+RetSqlName("SED")+" SED                                                 							 " + CRLF
	_cQuery += "ON E7_NATUREZ = ED_CODIGO AND SED.D_E_L_E_T_ = ' '                                   							 " + CRLF
	_cQuery += "left JOIN (SELECT TSE5.E5_NATUREZ, SUM(TSE5.E5_VALOR) AS DESPESA                     							 " + CRLF
	_cQuery += "FROM "+RetSqlName("SE5")+" TSE5                                                      							 " + CRLF
	_cQuery += "WHERE YEAR(E5_DATA) = '"+Alltrim(_aResp[1])+"'                                       							 " + CRLF
	_cQuery += "AND MONTH(TSE5.E5_DATA) = '"+Alltrim(_aResp[2])+"'                                   							 " + CRLF
	_cQuery += "AND TSE5.E5_MOEDA NOT IN ('TB')                                                      							 " + CRLF
	_cQuery += "AND TSE5.E5_RECPAG = 'P'                                                             							 " + CRLF
	_cQuery += "AND TSE5.E5_SITUACA = ' '                                                            							 " + CRLF
	_cQuery += "AND TSE5.E5_TIPODOC NOT IN ('BA','DA','DC')                                          							 " + CRLF
	_cQuery += "AND TSE5.E5_BANCO = '077'                                                            							 " + CRLF
	_cQuery += "AND TSE5.E5_NATUREZ < '09'                                                           							 " + CRLF
	_cQuery += "AND TSE5.E5_LA NOT IN('S')                                                           							 " + CRLF
	_cQuery += "AND TSE5.D_E_L_E_T_ = ' '                                                            							 " + CRLF
	_cQuery += "GROUP BY TSE5.E5_NATUREZ) TSE5                                                       							 " + CRLF
	_cQuery += "ON SE7.E7_NATUREZ = TSE5.E5_NATUREZ                                                  							 " + CRLF
	_cQuery += "LEFT JOIN (SELECT TSE2.E2_NATUREZ, SUM(TSE2.E2_SALDO) AS TITULO_PAGAR                							 " + CRLF     
    _cQuery += "FROM "+RetSqlName("SE2")+" TSE2                                                      							 " + CRLF             
    _cQuery += "WHERE YEAR(E2_VENCTO) = '"+Alltrim(_aResp[1])+"'                                     							 " + CRLF                                      
	_cQuery += "AND MONTH(TSE2.E2_VENCTO) = '"+Alltrim(_aResp[2])+"'                                 							 " + CRLF                                   
    _cQuery += "AND TSE2.E2_SALDO > 0                                                                							 " + CRLF 
	_cQuery += "AND TSE2.D_E_L_E_T_ = ' '                                                            							 " + CRLF 
	_cQuery += "GROUP BY TSE2.E2_NATUREZ) TSE2                                                       							 " + CRLF 
	_cQuery += "ON SE7.E7_NATUREZ = TSE2.E2_NATUREZ                                                  							 " + CRLF
	_cQuery += "WHERE E7_ANO = '"+Alltrim(_aResp[1])+"'                                              							 " + CRLF
	_cQuery += "AND SE7.D_E_L_E_T_ = ' '                                                             							 " + CRLF
	_cQuery += "ORDER BY 9 DESC "  

		If !Empty(_cQuery)
	
			If !lIsDir(cPath)
				nRet := MakeDir( cPath, Nil, .F. )
				
				if nRet != 0
					Alert( "Nใo foi possํvel criar o diret๓rio "+cPath+", crie manualmente. Erro: " + cValToChar( FError() ) )
				endif
				
			Endif
	
			TryException

				cQry      := getNextAlias()		
				DbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery),cQry,.F.,.F.)
				(cQry)->(DbEval({|| nCnt++}))   			
				nTotReg += nCnt
				(cQry)->(DbGoTop())
				//Atribui as colunas da query no array(Que forma o cabe็alho das colunas)  aStruQry[N][1] = Nome Campo, aStruQry[N][2] = tipo Campo, aStruQry[N][3] = Tamanho e aStruQry[N][4] = Decimal
				aStruQry  := (cQry)->(dbStruct())
	
				While !(cQry)->(Eof())
					//Calculo Totalizador despesa
					nValDesp += (cQry)->DESPESA
					(cQry)->(dbSkip())
				End
   		       (cQry)->(DBGOTOP())
			CatchException using oException
				Alert("Houve um erro na execu็ใo da Query, por favor verifique! " + TcSqlError())
				Return
			EndException
		Endif

		oBrush1 := TBrush():New(, RGB(193,205,205))
 
		// Verifica se o Excel estแ instalado na mแquina
		/* 
		If !ApOleClient("MSExcel")
			MsgAlert("Microsoft Excel nใo instalado!", "Aten็ใo")
	   		Return		
		EndIf
		*/
		oExcel  := FWMSExcel():New()//M้todo construtor da classe
		cAba1    := "Rel. Orcamento Despesa"
		cTabela1 := "Relat๓rio Orcamento Despesa"
		
		// Cria็ใo de nova aba 
		oExcel:AddworkSheet(cAba1)//Adiciona uma Worksheet ( Planilha )
		
		// Cria็ใo de tabela
		oExcel:AddTable (cAba1,cTabela1)//Adiciona uma tabela na Worksheet. Uma WorkSheet pode ter apenas uma tabela
		
		// Cria็ใo de colunas 
		For nCnt := 1 To Len(aStruQry)                                                                                     //Totalizo se for coluna DESPESA
			oExcel:AddColumn(cAba1,cTabela1,aStruQry[nCnt,1],IIF(aStruQry[nCnt,2]=="N",3,1),IIF(aStruQry[nCnt,2]=="N",1,1),.F.)//Adiciona uma coluna a tabela de uma Worksheet
		Next nCnt

      // Adiciona Linha das Despesas
		While !(cQry)->(Eof())
		    cMes := alltrim((cQry)->MES)
			nPoslin++
 		 	Incproc("Processando Registro " + StrZero(nPoslin,6) + " de " + StrZero(nTotReg,6))
		    // Cria็ใo de Linhas
			aLinExcel := {}
			For nCnt := 1 To Len(aStruQry)
				If Alltrim(aStruQry[nCnt,1]) == "DATA"
					aAdd(aLinExcel,DToC(SToD((cQry)->&(aStruQry[nCnt,1]))) )
				Else
					aAdd(aLinExcel,(cQry)->&(aStruQry[nCnt,1]))
				Endif
			Next nCnt

			oExcel:AddRow(cAba1,cTabela1, aLinExcel)
			(cQry)->(dbSkip())

		End

		If !Empty(oExcel:aWorkSheet)
 		 	cArquivo  := "ORCDESPR_"+ cMes +".xls"
		 	fErase(cPath +"\" +cArquivo)

			oExcel:Activate()
			oExcel:GetXMLFile(cArquivo)
		 
			CpyS2T("\SYSTEM\"+cArquivo, cPath)
		
			oExcelApp := MsExcel():New()
			oExcelApp:WorkBooks:Open(cPath+cArquivo) // Abre a planilha
			oExcelApp:SetVisible(.T.)
			oExcelApp:Destroy() //Encerra o processo do gerenciador de tarefas
		
		EndIf
		(cQry)->(DbCloseArea())
	Endif
Return  
************************************************************************************
Static Function MyPerg(_aResp,_lSup)

Local _aPergs   := {}

  aAdd( _aPergs,{ 1, "Ano: "           , Space(04),                          ,'naovazio()'      , ,'.T.', 50, .F.})
  aAdd( _aPergs,{ 1, "Mes: "           , Space(02),                          ,'naovazio()', ,'.T.', 50, .T.})
//aAdd( _aPergs,{ 1, "Emissใo de : "  , dDataBase, PesqPict("SE1", "E1_EMISSAO"),'naovazio()',"" ,'.T.', 50, .T.})
//aAdd( _aPergs,{ 1, "Emissใo at้: "  , dDataBase, PesqPict("SE1", "E1_EMISSAO"),'naovazio()',"" ,'.T.', 50, .T.})
//aAdd( _aPergs,{ 1, "Parceiro : "  , Space(06), PesqPict("SA1", "A1_COD"),'.T.',"SA1" ,'.T.', 50, .F.})

_lRet := ParamBox( _aPergs ,"Parโmetros", @_aResp, {|| .t. },,,,,,,.T.,.T.)

Return _lRet
