#Include 'Protheus.ch'
#Include 'TbiConn.ch'
#Include 'TryException.ch'
#DEFINE CRLF (Chr(13)+Chr(10))

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ RECDESPR บAutor  ณEvandro Cleto       บ Data ณ  04/05/21   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Relatorio Excel para Despesa e Receita   .                 บฑฑ
ฑฑบ          ณ                                                            บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Meu mesmo                                                  บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function RECDESPR() 
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
	Local aPDesp := {"PERC_DESPESA","N",5,2}//Array para montar Colunas de Percentual de despesa
	Local aPRec  := {"PERC_RECEITA","N",5,2}//Array para montar Colunas de Percentual de receita
	Local aRec   := {"RECEITA","N",15,2}//Array para montar Colunas de Receita
	Local nTPerdesp := 0
	Local nTPerRec  := 0
	Local nPerdesp  := 0
	Local nPerRec   := 0
	Private nTotReg := 0
	Private nPoslin := 0 
	Private _aResp  := {}	
    Private _cQuery   := ""
	Private  cQry     := ""
	Private _cQRec    := ""
	Private cQRc      := "" 
	Private cMes := ""


	if MyPerg(@_aResp,.T.)
	//CALCULO DE DESPESA
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
	_cQuery += "ED_DESCRIC DESC_NATUREZA,E5_NATUREZ COD_NATUREZA, SUM( E5_VALOR) AS DESPESA " + CRLF
    _cQuery += "FROM       "+RetSqlName("SE5")+" SE5                                        " + CRLF 
    _cQuery += "INNER JOIN "+RetSqlName("SED")+" SED                                        " + CRLF
	_cQuery += "ON E5_NATUREZ = ED_CODIGO AND SE5.D_E_L_E_T_ = ' '                          " + CRLF
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
	_cQuery += "GROUP BY YEAR(E5_DATA), MONTH(E5_DATA), ED_DESCRIC, E5_NATUREZ              " + CRLF
	_cQuery += "ORDER BY SUM( E5_VALOR) DESC    
   
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
				//Adiciono cabe็alho Coluna de Percentual Despesa
				aAdd(aStruQry,aPDesp)
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

		//CALCULO DE RECEITA	

    _cQRec := "SELECT YEAR(E5_DATA) as ANO,                                                " + CRLF
	_cQRec += "CASE MONTH(E5_DATA)                                                         " + CRLF
	_cQRec += " WHEN '1' THEN 'JANEIRO'                                                    " + CRLF
	_cQRec += " WHEN '2' THEN 'FEVEREIRO'                                                  " + CRLF
	_cQRec += " WHEN '3' THEN 'MARCO'                                                      " + CRLF
	_cQRec += " WHEN '4' THEN 'ABRIL'                                                      " + CRLF
	_cQRec += " WHEN '5' THEN 'MAIO'                                                       " + CRLF
	_cQRec += " WHEN '6' THEN 'JUNHO'                                                      " + CRLF
	_cQRec += " WHEN '7' THEN 'JULHO'                                                      " + CRLF
	_cQRec += " WHEN '8' THEN 'AGOSTO'                                                     " + CRLF
	_cQRec += " WHEN '9' THEN 'SETEMBRO'                                                   " + CRLF
	_cQRec += " WHEN '10' THEN 'OUTUBRO'                                                   " + CRLF
	_cQRec += " WHEN '11' THEN 'NOVEMBRO'                                                  " + CRLF
	_cQRec += " WHEN '12' THEN 'DEZEMBRO'                                                  " + CRLF
	_cQRec += "END AS MES,                                                                 " + CRLF
	_cQRec += "ED_DESCRIC DESC_NATUREZA,E5_NATUREZ COD_NATUREZA, SUM( E5_VALOR) AS RECEITA " + CRLF
    _cQRec += "FROM       "+RetSqlName("SE5")+" SE5                                        " + CRLF 
    _cQRec += "INNER JOIN "+RetSqlName("SED")+" SED                                        " + CRLF
	_cQRec += "ON E5_NATUREZ = ED_CODIGO AND SE5.D_E_L_E_T_ = ' '                          " + CRLF
	_cQRec += "WHERE SE5.E5_FILIAL = '"+xFilial("SE5")+"'                                  " + CRLF
	_cQRec += "AND YEAR(E5_DATA) = '"+Alltrim(_aResp[1])+"'                                " + CRLF
	_cQRec += "AND MONTH(E5_DATA) = '"+Alltrim(_aResp[2])+"'                               " + CRLF
	_cQRec += "AND E5_MOEDA NOT IN ('TB')                                                  " + CRLF
	_cQRec += "AND E5_RECPAG = 'R'                                                         " + CRLF
	_cQRec += "AND E5_SITUACA = ' '                                                        " + CRLF
	_cQRec += "AND E5_TIPODOC NOT IN ('BA','DA','DC')                                      " + CRLF
	_cQRec += "AND E5_BANCO = '077'                                                        " + CRLF
	_cQRec += "AND E5_NATUREZ < '09'                                                       " + CRLF
	_cQRec += "AND E5_LA NOT IN('S')                                                       " + CRLF
	_cQRec += "AND SE5.D_E_L_E_T_ = ' '                                                    " + CRLF
	_cQRec += "GROUP BY YEAR(E5_DATA), MONTH(E5_DATA), ED_DESCRIC, E5_NATUREZ              " + CRLF
	_cQRec += "ORDER BY SUM( E5_VALOR) DESC    
   
		If !Empty(_cQRec)
		
			TryException

				cQRc      := getNextAlias()		
				DbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQRec),cQRc,.F.,.F.)
				// Processando as linhas
				(cQRc)->(DbEval({|| nCount++}))   			
				nTotReg += nCount
				Procregua(nTotReg)
				(cQRc)->(DbGoTop())
				//Adiciono cabe็alho Coluna de RECEITA
				aAdd(aStruQry,aRec)
				//Adiciono cabe็alho Coluna de Percentual RECEITA
				aAdd(aStruQry,aPRec)
				While !(cQRc)->(Eof())
					//Calculo Totalizador RECEITA
					nValRec += (cQRc)->RECEITA
					(cQRc)->(dbSkip())
				End
   		       (cQRc)->(DBGOTOP())
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
		cAba1    := "Rel. Receita Despesa"
		cTabela1 := "Relat๓rio Receita Despesa"

		
		// Cria็ใo de nova aba 
		oExcel:AddworkSheet(cAba1)//Adiciona uma Worksheet ( Planilha )

		
		// Cria็ใo de tabela
		oExcel:AddTable (cAba1,cTabela1)//Adiciona uma tabela na Worksheet. Uma WorkSheet pode ter apenas uma tabela
		
		// Cria็ใo de colunas 
		For nCnt := 1 To Len(aStruQry)                                                                                     //Totalizo se for coluna DESPESA
			oExcel:AddColumn(cAba1,cTabela1,aStruQry[nCnt,1],IIF(aStruQry[nCnt,2]=="N",3,1),IIF(aStruQry[nCnt,2]=="N",1,1),.F.)//Adiciona uma coluna a tabela de uma Worksheet
		Next nCnt
	   // Adiciona Linha das Receitas
		While !(cQRc)->(Eof())
		    cMes := alltrim((cQRc)->MES)
			nPoslin++
 		 	Incproc("Processando Registro " + StrZero(nPoslin,6) + " de " + StrZero(nTotReg,6))
		    // Cria็ใo de Linhas
			aLinExcel := {}
			For nY := 1 To Len(aStruQry)
				If Alltrim(aStruQry[nY,1]) == "DATA"
					aAdd(aLinExcel,DToC(SToD((cQRc)->&(aStruQry[nY,1]))) )
				Elseif Alltrim(aStruQry[nY,1]) == "DESPESA"
						aAdd(aLinExcel,0)
				Elseif Alltrim(aStruQry[nY,1]) == "PERC_DESPESA"
						aAdd(aLinExcel,nPerdesp)
				Elseif Alltrim(aStruQry[nY,1]) == "PERC_RECEITA"
					aAdd(aLinExcel,nPerRec)
				Else
					aAdd(aLinExcel,(cQRc)->&(aStruQry[nY,1]))
				Endif
			Next nY

			oExcel:AddRow(cAba1,cTabela1, aLinExcel)
			(cQRc)->(dbSkip())

		End
      // Adiciona Linha das Despesas
		While !(cQry)->(Eof())
			nPoslin++
 		 	Incproc("Processando Registro " + StrZero(nPoslin,6) + " de " + StrZero(nTotReg,6))
		    // Cria็ใo de Linhas
			aLinExcel := {}
			For nCnt := 1 To Len(aStruQry)
				If Alltrim(aStruQry[nCnt,1]) == "DATA"
					aAdd(aLinExcel,DToC(SToD((cQry)->&(aStruQry[nCnt,1]))) )
				Elseif Alltrim(aStruQry[nCnt,1]) == "RECEITA"
						aAdd(aLinExcel,0)
				Elseif Alltrim(aStruQry[nCnt,1]) == "PERC_DESPESA"
					nPerdesp := (((cQry)->DESPESA/nValDesp)*100)
					aAdd(aLinExcel,nPerdesp)
					nTPerdesp += nPerdesp
				Elseif Alltrim(aStruQry[nCnt,1]) == "PERC_RECEITA"
					nPerRec := (((cQry)->DESPESA/nValRec)*100)
					aAdd(aLinExcel,nPerRec)
					nTPerRec += nPerRec
				Else
					aAdd(aLinExcel,(cQry)->&(aStruQry[nCnt,1]))
				Endif
			Next nCnt

			oExcel:AddRow(cAba1,cTabela1, aLinExcel)
			(cQry)->(dbSkip())

		End
		// Adiciono Totalizador das Despesas e Receitas
		aLinExcel := {}
		For nX := 1 To Len(aStruQry)
			If Alltrim(aStruQry[nX,1]) == "DESPESA"
				aAdd(aLinExcel,nValDesp)
			Elseif Alltrim(aStruQry[nX,1]) == "COD_NATUREZA"
				aAdd(aLinExcel,"Total: ")
			Elseif Alltrim(aStruQry[nX,1]) == "PERC_DESPESA"
				aAdd(aLinExcel,nTPerdesp)
			Elseif Alltrim(aStruQry[nX,1]) == "RECEITA"
				aAdd(aLinExcel,nValRec)
			Elseif Alltrim(aStruQry[nX,1]) == "PERC_RECEITA"
				aAdd(aLinExcel,nTPerRec)
			Else	
				aAdd(aLinExcel,"**")
			Endif
		Next nx

		oExcel:AddRow(cAba1,cTabela1, aLinExcel)

		// Adiciono Lucro ou Prejuizo
		aLinExcel := {}
		For nX := 1 To Len(aStruQry)
			If Alltrim(aStruQry[nX,1]) == "DESPESA"
				aAdd(aLinExcel,0)
			Elseif Alltrim(aStruQry[nX,1]) == "COD_NATUREZA"
				aAdd(aLinExcel,"Lucro/Prejuizo: ")
			Elseif Alltrim(aStruQry[nX,1]) == "PERC_DESPESA"
				aAdd(aLinExcel,0)
			Elseif Alltrim(aStruQry[nX,1]) == "RECEITA"
				aAdd(aLinExcel,(nValRec - nValDesp))
			Elseif Alltrim(aStruQry[nX,1]) == "PERC_RECEITA"
				aAdd(aLinExcel,0)
			Else	
				aAdd(aLinExcel,"**")
			Endif
		Next nx

		oExcel:AddRow(cAba1,cTabela1, aLinExcel)


		If !Empty(oExcel:aWorkSheet)
		    cArquivo  := "RECDESPR_"+cMes+".XLS"
			oExcel:Activate()
			oExcel:GetXMLFile(cArquivo)
		 
			CpyS2T("\SYSTEM\"+cArquivo, cPath)
		
			oExcelApp := MsExcel():New()
			oExcelApp:WorkBooks:Open(cPath+cArquivo) // Abre a planilha
			oExcelApp:SetVisible(.T.)
			oExcelApp:Destroy() //Encerra o processo do gerenciador de tarefas
		
		EndIf
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
