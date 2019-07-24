#INCLUDE "Protheus.CH"
#INCLUDE "TopConn.CH"
#INCLUDE "ap5mail.ch"

#DEFINE DS_MODALFRAME   128

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ WFW120P  º Autor ³ Dan Adrignoli        º Data ³ 02/07/19   º±±           
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Workflow de Aprovacao de Pedidos de Compra                  º±±
±±º          Ã-------------------------------------------------------------¶±±
±±º          ³ Na emissao do WorkFlow para aprovacao, manda e-mail         º±±
±±º          ³ Informativo para usuario definido no Parametro (MV_MAILAP2) º±±
±±º          ³ Esse e-mail nao consegue aprovar o pedido de compra.        º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Especifico TOPCAU                                           º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function WFW120P(nOpcao,oProcess)


If ValType(nOpcao) = "A"
	nOpcao := nOpcao[1]
Endif

If nOpcao == NIL
	nOpcao := 0
End

Do Case
	Case nOpcao == 0
		U_PC001E(oProcess)
	Case nOpcao == 1
		U_PC001R(oProcess)
	Case nOpcao == 2
		U_PC001T(oProcess)
EndCase

//oProcess:Free()

Return .T.
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ 	     Function PC001E               ³
//³Envio do workFlow de Pedido de Compra ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
User Function PC001E(oProcess,cPc,lRev)

Local lAprova02		:= .T.
Local oHTML			:= Nil
Local aVetProd		:= {}
Local cAprovador	:= ""
Local cEstrut		:= ""
Local cObsAnt		:= ""
Local cValMed		:= ""
Local cMotEnv		:= ""
Local nVlrTotal		:= 0
Local nSC7ItFrSg	:= 0
Local nVlrICMS		:= 0
Local nAliqIPI		:= 0
Local nVlrIPI		:= 0
Local nVlrLiqui		:= 0
Local nSaldoAtu		:= 0
Local nUltPreco		:= 0
Local nEnvios		:= 0
Local nC			:= 0
Local nI			:= 0
Local nH			:= 0
Local nD			:= 0
Local nU			:= 0
Local nOpcA			:= 0
Local nItemPed		:= 0
Local cMotivo 		:= Space(080)
Local cUsuarioProtheus := SubStr(__CUSERID,7,15)
Local aFiles		:= {}
Local nF			:= 0
Local cCodAprovador	:= ""
Local oMotivo
Local oDlg

Private cGrupoAprv	:= ""

oProcess := TWFProcess():New("WF0005", "Requisição de Aprovação de Pedido de Compras" )
oProcess:NewTask("WF0005","/WORKFLOW/WFW120.htm")

//oProcess:oWF:cMessengerDir := "/workflow/emp"+cEmpAnt+"wfw120/"

If cPc == Nil
	cPC := SC7->C7_NUM
EndIf

SC7->(DbSetorder(1))
SC7->(DbSeek(xFilial("SC7") + cPC))
While !SC7->(EOF()) .AND. SC7->C7_FILIAL == xFilial("SC7") .AND. SC7->C7_NUM == Alltrim(cPC)
	nItemPed ++
	SC7->(dbskip())
Enddo

SC7->(DbSetorder(1))
SC7->(DbSeek(xFilial("SC7") + cPC))

//cTexto := "Iniciando Processo..."
//cCodStatus := "100100"
//oProcess:Track(cCodStatus, cTexto, cUsuarioProtheus)

oHtml := oProcess:oHTML

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Cabecalho do Pedido               	  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
//oProcess:oHTML:ValByName('WFMailTo'	, "workflow@topcau.com.br")
oProcess:oHTML:ValByName('emissao'		, DtoC(SC7->C7_EMISSAO))
oProcess:oHTML:ValByName('fornecedor'	, SC7->C7_FORNECE + "/" + SC7->C7_LOJA)
oProcess:oHTML:ValByName('lb_nome'		, Posicione("SA2",1,xFilial("SA2") + SC7->C7_FORNECE + SC7->C7_LOJA, "A2_NOME"))
oProcess:oHTML:ValByName('lb_cond'		, Posicione("SE4",1,xFilial("SE4") + SC7->C7_COND, "E4_DESCRI"))
oProcess:oHTML:ValByName('pedido'		, SC7->C7_NUM)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Cabecalho (Natureza / Solicitante)	  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
//oProcess:oHTML:ValByName('natureza'	, SC7->C7_NATUREZ +" / " + Posicione("SED",1,xFilial("SED") + SC7->C7_NATUREZ, "SED->ED_DESCRIC"))
oProcess:oHTML:ValByName('solic'	, cUserName)
oProcess:oHTML:ValByName('ccusto'	, POSICIONE("CTT",1, xFilial("CTT") + SC7->C7_CC, "CTT_DESC01"))

nVlrTotal	:= 0
nVlrLiqui	:= 0
nSC7ItFrSg	:= 0
nVlrICMS	:= 0
nVlrIPI		:= 0

//RastreiaWF(oProcess:fProcessID + '.' + oProcess:fTask, "100001","100001")
oProcess:UserSiga := __CUSERID

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³  Itens do Pedido de Compra  		  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
SC7->(DbSetorder(1))
SC7->(DbSeek(xFilial("SC7") + Alltrim(cPC)))
While !SC7->(EOF()) .AND. SC7->C7_FILIAL == xFilial("SC7") .AND. SC7->C7_NUM == Alltrim(cPC)
	
	cNomeForn := Posicione("SA2",1,xFilial("SA2") + SC7->C7_FORNECE + SC7->C7_LOJA, "A2_NOME")
	cGrupoAprv:= SC7->C7_APROV
	
	Aadd((oProcess:oHTML:ValByName("produto.item"))		, SC7->C7_ITEM)
	//Aadd((oProcess:oHTML:ValByName("produto.obs"))		, SC7->C7_OBS)
	Aadd((oProcess:oHTML:ValByName("produto.codigo"))	, SC7->C7_PRODUTO)
	Aadd((oProcess:oHTML:ValByName("produto.descricao")), SC7->C7_DESCRI)
	Aadd((oProcess:oHTML:ValByName("produto.quant"))	, Transform(SC7->C7_QUANT, "@E 999999999.99"))
	Aadd((oProcess:oHTML:ValByName("produto.unid"))		, SC7->C7_UM)
	Aadd((oProcess:oHTML:ValByName("produto.preco"))	, Transform(SC7->C7_PRECO, "@E 999,999,999.9999"))
	Aadd((oProcess:oHTML:ValByName("produto.ctconta"))	, POSICIONE("CT1",1,xFilial("CT1") + SC7->C7_CONTA , "CT1_DESC01"))
	Aadd((oProcess:oHTML:ValByName("produto.total"))	, Transform(SC7->C7_TOTAL + SC7->C7_VALIPI, "@E 999,999,999.99"))
	Aadd((oProcess:oHTML:ValByName("produto.entrega"))	, DtoC(SC7->C7_DATPRF))
	
	nVlrLiqui	+= SC7->C7_TOTAL
	nVlrTotal	+= (SC7->C7_TOTAL+SC7->C7_VALIPI)
	
	RecLock("SC7",.F.)
	SC7->C7_WFID := oProcess:fProcessID
	MsUnlock("SC7")
	
	SC7->(dbskip())
Enddo

SC7->(DbSetorder(1))
SC7->(DbSeek(xFilial("SC7") + Alltrim(cPC)))

If lRev
	cCodAprovador := U_Aprovador(1,Alltrim(cPC),.T.)    //Retornar Código do Aprovador
Else
	cCodAprovador := U_Aprovador(1,Alltrim(cPC),.F.)    //Retornar Código do Aprovador
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³  Valor Total                		  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
//oProcess:oHTML:ValByName('lbvalor'	, Transform(nVlrLiqui	, "@E 999,999,999.99"))
oProcess:oHTML:ValByName('lbvalor'	, Transform(nVlrTotal	, "@E 999,999,999.99"))

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Botoes para Aprovacao             	  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
oProcess:oHTML:ValByName( 'RBAPROVA' , "Sim" )
oProcess:oHTML:ValByName( 'lbmotivo' , "" )
oProcess:oHTML:ValByName( 'CodAprovador' , If(Empty(cCodAprovador),"",cCodAprovador) )

dbSelectArea("SAK")
SAK->(dbSetOrder(2))
SAK->(dbSeek(xFilial("SAK") + cCodAprovador))
oProcess:oHTML:ValByName('cLogin',Alltrim(SAK->AK_LOGIN))
oProcess:ClientName(cUserName)

oProcess:cTo:= "WFW120"
oProcess:cSubject := "Aprovação do pedido de compra"
oProcess:bReturn:= "U_WFW120P(1)"

cMailID := oProcess:Start()
oProcess:NewTask("WF0005", "/workflow/wf_linkped.htm")

//Conout("Inicando Processo de Retorno dos Email" + cMailID)
                                                  
nHandle := 0 
cVar := ""
cAux	:= ""           
cBuffer := ""
cFile := "web/workflow/compras/messenger/emp"+cEmpAnt+"/wfw120/"+cMailId+".htm"

If File(cFile)
	nHandle := FOpen("web/workflow/compras/messenger/emp"+cEmpAnt+"/wfw120/"+cMailId+".htm")
	If nHandle > 0
		fRead(nHandle,@cBuffer,16384)                
		
		While !Empty(Alltrim(cBuffer) )                                    
			If '<script src="cid:SIGAWF' $ SubStr(cBuffer, 1, At(Chr(13),cBuffer))
				cAux += '<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>' + CRLF
			Else                                                     
				cAux += SubStr(cBuffer, 1, At(Chr(13),cBuffer))				
			EndIf                    
			If At(Chr(13),cBuffer) > 0
				cBuffer := SubStr(cBuffer, At(Chr(13),cBuffer) + 1, Len(cBuffer) ) 
			Else
				cBuffer := ""
			EndIf
		End
		
		fClose(nHandle) 
		fERase(cFile)
		
		nHandle := 0
		nHandle := fCreate(cFile)
		If nHandle > 0
			fWrite(nHandle,cAux)
		EndIf
		
		fClose(nHandle)
	EndIf		
	             
EndIf  
    

dbSelectArea("SAK")
SAK->(dbSetOrder(2))
SAK->(dbSeek(xFilial("SAK") + cCodAprovador))

SC7->(DbSetorder(1))
SC7->(DbSeek(xFilial("SC7") + Alltrim(cPC)))

oProcess:oHTML:ValByName( 'cNumWF'		, SC7->C7_NUM )
oProcess:oHTML:ValByName( 'cForneceWF'	, cNomeForn )                                                     '
oProcess:oHTML:ValByName('cCustoLK'		, POSICIONE("CTT",1,xFilial("CTT") + SC7->C7_CC,"CTT_DESC01"))

While !SC7->(EOF()) .AND. SC7->C7_FILIAL == xFilial("SC7") .AND. SC7->C7_NUM == Alltrim(cPC)
	
	Aadd((oProcess:oHTML:ValByName("produto.itemLK"))		, SC7->C7_ITEM)
	//Aadd((oProcess:oHTML:ValByName("produto.obsLK"))		, SC7->C7_OBS)
	Aadd((oProcess:oHTML:ValByName("produto.codigoLK"))		, SC7->C7_PRODUTO + " - "  + SC7->C7_DESCRI)
	Aadd((oProcess:oHTML:ValByName("produto.quantLK"))		, Transform(SC7->C7_QUANT, "@E 999999999.99"))
	Aadd((oProcess:oHTML:ValByName("produto.unidLK"))		, SC7->C7_UM)
	Aadd((oProcess:oHTML:ValByName("produto.precoLK"))		, Transform(SC7->C7_PRECO, "@E 999,999,999.9999"))
	Aadd((oProcess:oHTML:ValByName("produto.ctcontaLK"))	, POSICIONE("CT1",1,xFilial("CT1") + SC7->C7_CONTA , "CT1_DESC01"))
	Aadd((oProcess:oHTML:ValByName("produto.totalLK"))		, Transform(SC7->C7_TOTAL + SC7->C7_VALIPI, "@E 999,999,999.99"))

	SC7->(dbskip())
Enddo

oProcess:oHTML:ValByName('lbvalorLK'	, Transform(nVlrTotal	, "@E 999,999,999.99"))

oProcess:ohtml:ValByName('proc_linkI','http://189.89.43.140:6100/compras/messenger/emp' + cEmpAnt + '/wfw120/'+ cMailId + '.htm' )

//cTexto := "Enviando solicitação..."            
//cCodStatus := "100300"
//oProcess:Track(cCodStatus, cTexto, cUsuarioProtheus)

SAK->(DbSeek(xFilial("SAK")))
If lRev
	oProcess:cTo := U_Aprovador(2,Alltrim(cPC),.T.)
Else
	oProcess:cTo := U_Aprovador(2,Alltrim(cPC),.F.)
EndIf

oProcess:Start()
//oProcess:Free()
Return .T.

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ 	     Function PC001R               ³
//³Retorno do workFlow de Pedido de Compra ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
User Function PC001R(oProcess)

Local cHtmlA 	:= ""
Local cHtmlR 	:= ""
Local cPedido 	:= ""
Local cCompra	:= ""
Local cMotivo	:= ""
Local lRet		:= ""
Local cIdAprov  := ""

Begin Sequence

//Conout("Inicando Processo de Retorno dos Email")

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Pedido de Compra Aprovado     ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If oProcess:oHtml:RetByName('RBAPROVA') == "Sim"

	
	//Conout("==>>PROCESSO DE APROVADO")
	cIdAprov := Substr(oProcess:oHtml:RetByName('CodAprovador'),1,6)
	
	cQuery := ""
	cQuery += "SELECT R_E_C_N_O_ RECNO1 FROM "+ RetSqlName("SCR") +" SCR "
	cQuery += "	WHERE SCR.D_E_L_E_T_='' "
	cQuery += "	  AND SCR.CR_NUM='"+Alltrim(oProcess:oHtml:RetByName('pedido'))+"'				"
	cQuery += "	  AND SCR.CR_USER='"+Substr(oProcess:oHtml:RetByName('CodAprovador'),1,6)+"'	"
	
	If Select("TRB") > 0
		TRB->(dbCloseArea())
	EndIf
	memowrite("SQLSCR.SQL",cQuery)
	
	TCQUERY cQuery NEW ALIAS "TRB"
	
	dbSelectArea("TRB")
	TRB->(dbGoTop())
	While TRB->(!Eof())

		SCR->(dbGoTo(TRB->RECNO1))
	
		RecLock("SCR",.F.)
		
		If SCR->CR_USER == cIdAprov
			SCR->CR_STATUS  := "03"
			SCR->CR_LIBAPRO := SCR->CR_APROV 
			SCR->CR_VALLIB  := SCR->CR_TOTAL
			SCR->CR_TIPOLIM := Posicione("SAK",2,xFilial("SAK")+cIdAprov,"AK_TIPO")
		Else
			SCR->CR_STATUS  := "05"
		EndIf

		SCR->CR_DATALIB := dDataBase
		SCR->CR_OBS     := ""
		SCR->CR_USERLIB := cIdAprov

		SCR->(MsUnLock())
	
		TRB->(dbSkip())
	EndDo
	
	cRet := U_Aprovador(2,oProcess:oHtml:RetByName('pedido'),.T.)
	
	IF Empty(cRet)
		//Monta HTML de Pedido Aprovado
		dbSelectArea("SC7")
		SC7->(dbGoTop(1))
		SC7->(dbSeek(xFilial("SC7") + oProcess:oHtml:RetByName('pedido')))
		While SC7->(!Eof()) .And. SC7->C7_FILIAL == xFilial("SC7") ;
			.And. SC7->C7_NUM == oProcess:oHtml:RetByName('pedido')
			RecLock("SC7",.F.)
			SC7->C7_CONAPRO := "L"
			SC7->(MsUnLock())
			SC7->(dbSkip())
		EndDo
		//Conout("==>>RESPONDENDO APROVACAO PARA COMPRADOR")
		cHtmlA := MONTAHTML(oProcess:oHtml:RetByName('pedido'),"A",cMotivo)
		//Conout("Para : "+ GetMv("HE_APROVPE"))
		SendMAIL(cHtmlA,GetMv("HE_APROVPE"),"Pedido de compra aprovado...") //Manda Email para Comprador
	Else
		Conout("==>>ENVIADO E-MAIL PARA PROXIMO APROVADOR. ==<<")
		PC001E(oProcess,oProcess:oHtml:RetByName('pedido'),.T.) //Envia e-mail para os demais aprovadores.
	EndIf
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Pedido de Compra Reprovado    ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
ElseIf oProcess:oHtml:RetByName('RBAPROVA') <> "Sim"
	
	Conout("==>>PROCESSO DE REPROVADO")
	
	cMotivo := oProcess:oHtml:RetByName('lbmotivo')
	
	cQuery := ""
	cQuery += "SELECT R_E_C_N_O_ RECNO1,* FROM "+ RetSqlName("SCR") +" SCR "
	cQuery += "	WHERE SCR.D_E_L_E_T_='' "
	cQuery += "	  AND SCR.CR_NUM='"+Alltrim(oProcess:oHtml:RetByName('pedido'))+"'				"
	
	If Select("TRB") > 0
		TRB->(dbCloseArea())
	EndIf
	
	memowrite("REPROVACAO.SQL",cQuery)
	
	TCQUERY cQuery NEW ALIAS "TRB"
	
	dbSelectArea("TRB")
	TRB->(dbGoTop())
	While TRB->(!Eof())
		dbSelectArea("SCR")
		SCR->(dbGoTo(TRB->RECNO1))
		If SCR->(!Eof())
			SCR->(RecLock("SCR", .F.))
			SCR->CR_DATALIB := dDataBase
			SCR->CR_OBS     := oProcess:oHtml:RetByName('lbmotivo')
			SCR->CR_STATUS  := "04"
			SCR->(MsUnLock())
		EndIf
		Conout("==>>ALTERANDO STATUS SCR PARA REPROVADO")
		TRB->(dbSkip())
	EndDo
	
	Conout("==>>ALTERADO STATUS DO PEDIDO DE COMPRA - REPROVADO" + oProcess:oHtml:RetByName('pedido') )
	SC7->(DbSetOrder(1))
	SC7->(DbSeek(xFilial("SC7") + Alltrim(oProcess:oHtml:RetByName('pedido'))))
	
	//cPedido := SC7->C7_NUM
	//cCompra	:= UsrRetMail(SC7->C7_USER)
	
	While !SC7->(Eof()) .And. SC7->C7_FILIAL == xFilial("SC7") ;
		.And. Alltrim(SC7->C7_NUM) == Alltrim(oProcess:oHtml:RetByName('pedido'))
			Conout("==>>ENTROU NO WHILE DO PEDIDO DE COMPRA")
			SC7->(RecLock("SC7", .F.))
			SC7->C7_CONAPRO := "B"
			SC7->C7_RESIDUO := "S"
			SC7->(MsUnLock())
		SC7->(DBSkip())
	Enddo
	
	//Monta HTML de Pedido Aprovado
	//Conout("==>>RESPONDENDO REPROVACAO PARA COMPRADOR")
	cHtmlA := MONTAHTML(oProcess:oHtml:RetByName('pedido'),"A",cMotivo)
	Conout("Para : "+ GetMv("HE_APROVPE"))
	SendMAIL(cHtmlA,GetMv("HE_APROVPE"),"Pedido de compra reprovado...") //Manda Email para Comprador

Endif

oProcess:Finish()

//oProcess:Free()
End Sequence

Conout("Finalizando Processo de Retorno dos Email")

Return .T.

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ 	     Function PC001T               ³
//³TimeOut do workFlow de Pedido de Compra ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
User Function PC001T(oProcess)
Begin Sequence
RastreiaWF(oProcess:fProcessID + '.' + oProcess:fTaskID, "100001", "100004")
oProcess:Finish()
oProcess:Free()
End Sequence
Return .T.

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ 	    Function Aprovador             ³
//³		Retorna o Nome do Aprovador        ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
User Function Aprovador(nOpc,cPedido,lRev)

Local cRetorno	:=""
Local cQuery	:=""
Local cNivel	:="01"
Local cAprovador:=""

Begin Sequence

cQuery := "SELECT * FROM " + RetSqlName("SCR") +" WHERE D_E_L_E_T_='' "
cQuery += " AND CR_NUM ='"+Alltrim(cPedido)+"'		"
cQuery += " AND CR_FILIAL = '"+xFilial("SCR")+"'	"
cQuery += " AND D_E_L_E_T_ = ''						"

If !lRev
	cQuery += " AND CR_STATUS='02'"
Else
	cQuery += " AND CR_STATUS='01'"
Endif

cQuery += " ORDER BY CR_NIVEL	"

If Select("TRB") > 0
	TRB->(dbCloseArea())
EndIf

memowrite("BUSCAAPROVADOR.SQL",cQuery)

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"TRB",.T.,.T.)

cNivel	:= TRB->CR_NIVEL
cRetorno:=""
TRB->(dbGoTop())
While TRB->(!EOF())
	Conout("==>>BUSCANDO NOVOS APROVADORES")
	If cNivel == TRB->CR_NIVEL
		If nOpc == 2
			cRetorno += UsrRetmail(TRB->CR_USER) + ";"
			//cRetorno += "danadrignoli@gmail.com;"
		Else
			cRetorno += TRB->CR_USER
		EndIf
		IF TRB->CR_STATUS == "01"
		   Conout("==>>ALTERANDO STATUS SCR 02")
		   dbSelectArea("SCR")
		   dbGoTo(TRB->R_E_C_N_O_)
		   RecLock("SCR",.F.)
		   TRB->CR_STATUS := "02"
		   SCR-(MsUnLock())
		EndIf
	EndIf
	TRB->(dbSkip())
EndDo

If nOpc == 2
	cRetorno := Substr(cRetorno, 1, Len(cRetorno) - 1) //retirar o ultimo;
EndIf

Conout("==>>RETORNO DA VARIAVEL cRetorno" + "|" + cRetorno)

End Sequence

Return(cRetorno)
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³   	       Function HTML               ³
//³   Monta HTML de Pedidos de Compras     ³
//³  Aprovados / Reprovados para Enviar p/ ³
//³              Comprador                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
User Function MONTAHTML(cPedido,cTp,cMotivo)
Local cHTML 	:= ""
Local aVetAp	:= {}
Local nTotal	:= 0
Local nX		:= 0
Local cNumPC	:= ""
Local dEmissao	:= CtoD("")
Local cCompra	:= ""
Local cAprova	:= ""
Local cFornece	:= ""
Local cCondP	:= ""

SC7->(DbSetOrder(1))
SC7->(DbSeek(xFilial("SC7") + cPedido ))
While !SC7->(EOF()) .AND. SC7->C7_NUM == cPedido
	
	nTotal 		:= nTotal + SC7->C7_TOTAL
	cNumPC		:= SC7->C7_NUM
	dEmissao	:= SC7->C7_EMISSAO
	//cCompra		:= SC7->C7_USUARIO
	cAprova		:= SC7->C7_APROV
	cFornece	:= Posicione("SA2",1,xFilial("SA2")+SC7->C7_FORNECE + SC7->C7_LOJA,"SA2->A2_NOME")
	cCondP		:= Posicione("SE4",1,xFilial("SE4")+SC7->C7_COND,"SE4->E4_DESCRI")
	
	Aadd(aVetAp,{	SC7->C7_ITEM,;		//01 - Item
	SC7->C7_PRODUTO,;	//02 - Produto
	SC7->C7_DESCRI,;	//03 - Descricao
	SC7->C7_UM,;		//04 - Un. Medida
	SC7->C7_QUANT,;		//05 - Quantidade
	SC7->C7_PRECO,;		//06 - Vl. Unit
	SC7->C7_VALIPI,;	//07 - Vlr. IPI
	SC7->C7_TOTAL})		//08 - Vlr. Total
	SC7->(DbSkip())
Enddo

cHTML += "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>"
cHTML += "	<html xmlns='http://www.w3.org/1999/xhtml'>"
cHTML += "	<head>"
cHTML += "	<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' />"
cHTML += "	<title>Untitled Document</title>"
cHTML += "	<style type='text/css'>"
cHTML += "	<!--"
cHTML += "	.style1 {font-size: 18px}"
cHTML += "	-->"
cHTML += "	</style>"
cHTML += "	</head>"

cHTML += "	<body>"
cHTML += "	<table width='845' border='1' cellspacing='1' bgcolor='#99CCFF'>"
cHTML += "	  <tr>"
cHTML += "	    <th width='388' class='style1' scope='col'>Pedido de Compra </th>"

If cTp == "A" //Aprovado / Reprovado
	cHTML += "	    <th width='444' scope='col'><span class='style1'>Aprovado</span></th>"
Else
	cHTML += "	    <th width='444' scope='col'><span class='style1'>Reprovado</span></th>"
EndIf
cHTML += "	  </tr>"

cHTML += "	  <tr>"
cHTML += "	    <th class='style1' scope='col'><div align='left'>Numero : "+cNumPc+"</div></th>"
cHTML += "	    <th class='style1' scope='col'><div align='left'>Motivo : "+cMotivo+"</div></th>"
cHTML += "	  </tr>"
cHTML += "	</table>"

cHTML += "	<table width='845' border='0' cellspacing='0'>"
cHTML += "	  <tr>"
cHTML += "	    <td>Emiss&atilde;o : </td>"
cHTML += "	    <td>"+DTOC(dEmissao)+"</td>"
cHTML += "	    <td>Valor Total : </td>"
cHTML += "	    <td>"+Transform(nTotal,"@E 999,999,999.99")+"</td>"
cHTML += "	  </tr>"

cHTML += "	  <tr>"
//cHTML += "	    <td>Comprador : </td>"
//cHTML += "	    <td>"+cCompra+"</td>"
cHTML += "	    <td>Aprovador : </td>"
cHTML += "	    <td>"+cAprova+"</td>"
cHTML += "	  </tr>"

cHTML += "	  <tr>"
cHTML += "	    <td>Fornecedor : </td>"
cHTML += "	    <td>"+cFornece+"</td>"
cHTML += "	    <td>Cond. Pagamento </td>"
cHTML += "	    <td>"+cCondP+"</td>"
cHTML += "	  </tr>"
cHTML += "	</table>"

cHTML += "	<table width='845' border='1' cellspacing='1'>"
cHTML += "	  <tr bgcolor='#99CCFF'>"
cHTML += "	    <th scope='col'>Item</th>"
cHTML += "	    <th scope='col'>C&oacute;digo</th>"
cHTML += "	    <th scope='col'>Descri&ccedil;&atilde;o</th>"
cHTML += "	    <th scope='col'>Um</th>"
cHTML += "	    <th scope='col'>Quant</th>"
cHTML += "	    <th scope='col'>Vl. Unit&aacute;rio </th>"
cHTML += "	    <th scope='col'>IPI</th>"
cHTML += "	    <th scope='col'>Vl. Total </th>"
cHTML += "	  </tr>"

For nX := 1 To Len(aVetAp)
	cHTML += "	  <tr>"
	cHTML += "	    <td>"+aVetAp[nX][1]+"</td>"
	cHTML += "	    <td>"+aVetAp[nX][2]+"</td>"
	cHTML += "	    <td>"+aVetAp[nX][3]+"</td>"
	cHTML += "	    <td>"+aVetAp[nX][4]+"</td>"
	cHTML += "	    <td>"+Transform(aVetAp[nX][5],"@E 999999999.99")+"</td>"
	cHTML += "	    <td>"+Transform(aVetAp[nX][6],"@E 999,999,999.9999")+"</td>"
	cHTML += "	    <td>"+Transform(aVetAp[nX][7],"@E 999,999,999.99")+"</td>"
	cHTML += "	    <td>"+Transform(aVetAp[nX][8],"@E 999,999,999.99")+"</td>"
	cHTML += "	  </tr>"
Next nX

cHTML += "	</table>"

cHTML += "	</body>"
cHTML += "	</html>"
Return cHTML


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³   	     Function SENDMAIL             ³
//³   Envia e-mail Aprovado / Reprovado    ³
//³           para Comprador               ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

User Function SENDMAIL(cCorpo,cPara,cAssunto)
Local lOk 		:= .F.
Local cErro 	:= ""
Local cServer   := GetMV("MV_RELSERV")
Local cAccount  := GetMV("MV_RELACNT")
Local cEnvia    := GetMV("MV_RELACNT")
Local cPassword := GetMV("MV_RELPSW")

Local cCC		:= ""

CONNECT SMTP SERVER cServer ACCOUNT cAccount PASSWORD cPassword Result lOK

If GetNewPar("MV_RELAUTH",.F.)
	lRetAuth := MailAuth(cAccount,cPassword)
Else
	lRetAuth := .T.
EndIf
If !lRetAuth
	lOK := .F.
	ConOut("Nao foi possivel autenticar o usuario e senha para envio de e-mail!")
EndIf

If (lOk)
	
	ConOut("CONECTADO AO SERVIDOR SMTP")
	ConOut("Enviando e-mail de :   "+cEnvia)
	ConOut("Enviando e-mail para : "+cPara)
	ConOut("Com Copia : "+cCC)
	ConOut("Assunto : "+cAssunto)
	
	SEND MAIL FROM cEnvia TO cPara CC cCC SUBJECT cAssunto BODY cCorpo RESULT lOk
	
	If (!lOk)
		GET MAIL ERROR cErro
		ConOut("==== ERRO AO ENVIAR E-MAIL (SENDMAIL)====")
		ConOut(cErro)
	Else
		Conout("E-MAIL DE RETORNO ENVIADO")
	EndIf
	
	DISCONNECT SMTP SERVER RESULT lOK
	
Else
	GET MAIL ERROR cErro
	ConOut("==== ERRO AO ENVIAR E-MAIL (CONEXAO COM SMTP)====")
	ConOut(cErro)
EndIf

Return