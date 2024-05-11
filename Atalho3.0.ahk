#SingleInstance Force
#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode RegEx

/*
Alt		!
Ctrl	^
Shift	+
Win		#
*/

^+r::Reload  ; Ctrl+Alt+R
^+e::Edit  ; Ctrl+Alt+R
F10::Suspend
^SPACE::
{
	Winset, Alwaysontop, , A
	return
}

ObterDataEHora() {
	; URL da API que fornece a data/hora
	url := "http://worldtimeapi.org/api/timezone/America/Sao_Paulo"

    ; Realiza uma solicitação HTTP GET para a API
    httpRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    httpRequest.Open("GET", url)
    httpRequest.Send()

    ; Verifica se a solicitação foi bem-sucedida
    if (httpRequest.Status = 200) {
        ; Analisa a resposta JSON para obter a data e hora
        resposta := httpRequest.ResponseText
        dataHora := RegExReplace(resposta, ".*""datetime"":"".*T(.*?)\..*", "$1")
        return dataHora
    } else {
        MsgBox Erro ao obter a data e hora externa.
        return ""
    }
}

; Definindo uma variável para controlar o estado do script: ON ou OFF
estadoScript := true

; Função para alternar o estado do script
alternarEstadoScript() {
	global estadoScript
	estadoScript := !estadoScript

	; Se o script estiver ligado, será suspenso; caso contrário, será ativado
	if (estadoScript)
		;~ Suspend
		MsgBox, Script Desligado!
	else
		;~ Suspend
		MsgBox, Script Ligado!
}

; Defina uma variável para controlar o estado do ícone na bandeja do sistema
iconeVisivel := true

; Função para alternar a visibilidade do ícone na bandeja do sistema
AlternarIconeTray() {
    global iconeVisivel
    iconeVisivel := !iconeVisivel  ; Inverte o estado atual

    ; Se o ícone estiver visível, oculta-o; caso contrário, exibe-o
    if (iconeVisivel)
        Menu, Tray, NoIcon
    else
        Menu, Tray, Icon
}

; Defina uma hotkey para chamar a função AlternarIconeTray
^+t::AlternarIconeTray()  ; Pressione Ctrl + Alt + T para alternar a visibilidade do ícone na bandeja do sistema

global numeroAtual := ""
global hematocrito, hmCao, hmGato, hemoglobina, volumeGlobular

chamarUserName() {
	global nomePC := A_UserName
	global caminhoImagens := "C:\Users\" . nomePC . "\Pictures\"
}

definirNumero() {
    InputBox, numeroAtual, Digite um número, Digite o número atual:
}

calcularValores() {

	Random, hbRan, 0.31, 0.35
	Random, hmDogRan, 6.5, 7.7
	Random, hmCatRan, 4.0, 5.5
	global hematocrito := Round(volumeGlobular, 0)
	global hemoglobina := StrReplace(Round(hematocrito*hbRan, 1), ".", ",")
	global hmCao := StrReplace(Round(hematocrito/hmDogRan, 2), ".", ",")
	global hmGato := StrReplace(Round(hematocrito/hmCatRan, 2), ".", ",")

	if (A_ThisHotkey == "End") {
		ControlGetPos,, ctrlY,,, OKttbx2, i)hemograma|reticulócitos
		if (ctrlY = 184){
		ControlGetText, volumeGlobular, OKttbx2, i)hemograma|reticulócitos
			if (volumeGlobular > 65)
				MsgBox Hematócrito Muito Alto!
		Send {Up 2}
	} else if (ctrlY = 164){
		ControlSetText, OKttbx2,, i)hemograma|reticulócitos
		Send, %hemoglobina%
		Send, {Enter}
	} else if (ctrlY = 144){
		ControlSetText, OKttbx2,, i)hemograma|reticulócitos
		Send, %hmCao%
		Send {Enter}
	}
	} else {
		ControlGetPos,, ctrlY,,, OKttbx2, i)hemograma|reticulócitos
		if (ctrlY = 184){
		ControlGetText, volumeGlobular, OKttbx2, i)hemograma|reticulócitos
		Send {Up 2}
	} else if (ctrlY = 164){
		ControlSetText, OKttbx2,, i)hemograma|reticulócitos
		Send, %hemoglobina%
		Send, {Enter}
	} else if (ctrlY = 144){
		ControlSetText, OKttbx2,, i)hemograma|reticulócitos
		Send, %hmGato%
		Send {Enter}
	}
	}
}

#+s::
	chamarUserName()
	if (nomePC == "LABORATÓRIO 02") {
		Send {PrintScreen}
	} else {
		Send ^{Insert}
	}
return

+#a::
	MouseClickDrag, L, 0, 0, 150, 0, 100, R
return


Esc:: ;Calcular posição do mouse global hematocrito, hmCao, hmGato, hemoglobina, volumeGlobular
{
	dataHoraExterna := ObterDataEHora()

	while GetKeyState("Esc")
	{
		MouseGetPos, xPos, yPos, WhichWindow, WhichControl, 1
		ControlGetPos, x, y, w, h, %WhichControl%, ahk_id %WhichWindow%
		;~ ToolTip, %WhichControl%`nX%X%`tY%Y%`nW%W%`t%H%
		;~ MouseGetPos xPos, yPos,, classControl, 1
		WinGetActiveTitle winTitle
		ToolTip X(%xPos%)`tY(%yPos%)`n%winTitle%`n%WhichWindow%`nNome do Controle: %WhichControl%`nX%X%`tY%Y%`nW%W%`t%H%`n%dataHoraExterna%`n%hbRan%
	}
	ToolTip
	return
}

procurarImagem() {
	global nomePC, caminhoImagens

    ; Lista de arquivos de imagem possíveis
    imageFiles := ["hemograma.png", "hemograma2.png", "hemograma3.png", "hemograma4.png", "hemograma5.png", "hemograma6.png"]

	; Loop para iterar sobre cada arquivo de imagem
	Loop, % imageFiles.MaxIndex()
	{
		chamarUserName()
        ; Obtém o caminho do arquivo de imagem atual
        currentImage := imageFiles[A_Index]
		currentPath := caminhoImagens . currentImage

        ; Realiza a pesquisa da imagem atual
        ;~ ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Users\LABORATÓRIO 02\Pictures\%currentImage%
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %caminhoImagens%%currentImage%
		;~ MsgBox % currentPath

        ; Verifica o resultado da pesquisa
        if (ErrorLevel = 0)
        {
            ; Se a imagem for encontrada, clique na posição encontrada
            MouseClick,, FoundX+485, FoundY+11
            ; Encerra o loop
            break
        }
	}

    ; Se nenhuma imagem for encontrada, exibe uma mensagem
    if (ErrorLevel = 1)
        MsgBox, Nenhuma imagem encontrada na tela.
return
}

^+!k::
	chamarUserName()
    Run, explorer.exe "%caminhoImagens%Screenshots"
return

Esconder := False

F12::
	if WinExist ("ahk_exe brave.exe") {
		if (Esconder) {
			WinShow, i)Brave ;Substitua "NomeDaJanela" pelo título da janela que você deseja manipular
			Esconder := False
		} else {
			WinHide, i)Brave
			Esconder := True
		}
	}
return

F3::
{
	global hbRan, hmDogRan, hmCatRanm, hematocrito, hemoglobina, hmCao, hmGato, volumeGlobular
	while GetKeyState("F3") {
		ToolTip,
	(Ltrim
	Hemoglobina Random: %hbRan%
	Hemácia Cão Random: %hmDogRan%
	Hemácia Gato Random: %hmCatRan%
	VolumeGlobular: %volumeGlobular%
	Hematócrito: %hematocrito%
	Hemoglobina: %hemoglobina%
	HmCão: %hmCao%
	HmGato: %hmGato%
	)
	}
	ToolTip
	return
}

;~ F4::
	;~ FileGetTime, dataArquivo, C:\Users\LABORATÓRIO 02\Pictures\Enviados\1205\Captura de Tela (378).png, C
	;~ FormatTime, dataDoArquivo, %dataArquivo%, ddMM
	;~ while GetKeyState(A_ThisHotkey)
		;~ ToolTip %dataDoArquivo%
	;~ ToolTip
;~ return

;~ F4::
	;~ origem := "C:\Users\LABORATÓRIO 02\Pictures\Enviados\*.*"
	;~ destino := "C:\Users\LABORATÓRIO 02\Pictures\Screenshots\"
	;~ Loop, Files, %origem%, R
	;~ {
		;~ arquivo := A_LoopFileFullPath
		;~ FileMove, %arquivo%, %destino%, 1
	;~ }
;~ return

F4::
	WinGet, active_id, ControlList, A
	ControlGetFocus, currentControl, A

	ControlGet, nextControl, Hwnd,, %currentControl%, A

	while GetKeyState(A_ThisHotkey)
		ToolTip, %TabCount%
	ToolTip
return

enviarLaudo() {
	if WinActive("i)hemograma|reticulócitos") && !WinActive("Solicitação de Nova Amostra"){
		Send #{PrintScreen}
		Sleep 500
		MouseClick,, 600, 1020 ;normal
	} else if WinActive("Solicitação de Nova Amostra"){
		MouseClick,, 600, 1020
	} else if WinActive("i)eas|fezes|microfilárias|giardia|parvovirose|hematozoários|sarna|fungo"){
		MouseClick,, 600, 1020 ;normal
	} else if WinActive("i)esporotricose"){
		MouseClick,, 620, 650 ;envia
	}
}

contarLaudos() {
	global nomePC, caminhoImagens
	chamarUserName()

	caminhoLaudos := "C:\Users\" . nomePC . "\Pictures\Screenshots\*.*"
	contador := 0
	Loop, Files, %caminhoLaudos%
	{
		contador++
	}
	MsgBox, %contador% laudos digitados até agora.
}

laudosConferidos() {
	global nomePC, caminhoImagens
	chamarUserName()
	origem := "C:\Users\" . nomePC . "\Pictures\Screenshots\*.*"

	if (FileExist(origem)) {
		contador := 0
		Loop, Files, %origem%
		{
			arquivo := A_LoopFileFullPath
			FileGetTime, dataArquivo, %arquivo%, C
			FormatTime, dataArquivo, %dataArquivo%, MMdd
			destino := "C:\Users\" . nomePC . "\Pictures\Enviados\" . dataArquivo
			if (!FileExist(destino))
				FileCreateDir, %destino%

			if (FileExist(destino))
				FileMove, %arquivo%, %destino%\, 1
				contador++
		}

		MsgBox %contador% arquivos foram movidos com sucesso.
	} else {
		MsgBox O caminho de origem é inválido.
	}
return
}

ativarJanela() {
	WinActivate i)informações
	WinActivate i)resultado
	WinActivate i)atenção
	WinActivate i)HF-LAB
	return
}

^g::laudosConferidos()
^+g::contarLaudos()

#IfWinActive i)identificação
{
	^l::
	{
		MouseClick,, 320, 360
		Send, luiz
		Send, {TAB}
		Send, 438398
		Send, {ENTER}{ENTER}
		return
	}
	^k::
	{
		MouseClick,, 320, 360
		Send, anacarolina
		Send, {TAB}
		Send, 120288
		Send, {ENTER}{ENTER}
		return
	}
}

#IfWinActive ahk_exe MSACCESS.EXE
Delete::
{
	if WinActive("ahk_exe MSACCESS.EXE"){
		ativarJanela()
		}
}

#IfWinActive i)número|realizados|liberar
	global numState := true
	NumpadAdd::
	{
		if (numeroAtual = "") {
			definirNumero()
			numeroFixo := numeroAtual
		} else {
			ControlGetFocus, controleFocado
			ControlSetText, %controleFocado%,,
			if (numState = true) {
				numeroAtual := numeroFixo
				Send %numeroAtual%
				numState := false
			} else {
				Send % --numeroAtual
				numState := true
				numeroAtual := ++numeroAtual
			}
		}
	return
	}
	NumpadSub::
		ControlGetFocus, controleFocado
		ControlSetText, %controleFocado%,,
		Send % --numeroAtual
	return
	NumpadMult::
		ControlGetFocus, controleFocado
		ControlSetText, %controleFocado%,,
		Send % ++numeroAtual
	return
	NumpadEnter::
	{
		numState := true
		send, {ENTER}
		return
	}

	return

#IfWinActive i)digitação
{
	WheelDown::MouseClick,, 180, 70
	XButton1::MouseClick,, 180, 70
	WheelUp::
	Up::
		ControlGetFocus, controlAtivo
		ControlGetPos, posX, posY, , , %controlAtivo%
	    if (posX = 17 && posY = 500) || (posX = 367 && posY = 61) {
			SendInput, {Up}
		} else {
			MouseClick,, 300, 520
		}
	return
	Left::
		ControlGetFocus, controlAtivo
		ControlGetPos, posX, posY, , , %controlAtivo%
	    if (posX = 17 && posY = 500) || (posX = 367 && posY = 61) {
			SendInput, {Left}
		} else {
			MouseClick,, 700, 70
		}
	return
	NumpadDiv::Send ICTERÍCIA
	NumpadMult::Send LIPEMIA
	NumpadSub::Send ?
	:*:CE::Confirmar Espécie`n
	^n::MouseClick,, 50, 1020
	^a::Send, +{Home}
	F5::
		FormatTime, dataHora, %A_Now%, dd/MM/yyyy
		Send, (%dataHora%)`n
		return
	^+v::MouseClick,, 700, 70
	Right::
		procurarImagem()
	return
	Insert::
		Send ICTERÍCIA{+}`n
		return
	Home::
		Send ICTERÍCIA{+}{+}`n
		return
	PgUp::
		Send ICTERÍCIA{+}{+}{+}`n
		return
	Delete::
		Send LIPEMIA{+}`n
		return
	End::
		Send LIPEMIA{+}{+}`n
		return
	PgDn::
		Send LIPEMIA{+}{+}{+}`n
		return


}

#IfWinActive i)informações
{
	NumpadEnter::MouseClick,, 480, 250
}

#IfWinActive i)hemograma|reticulócitos
{
	End::
	Home::calcularValores()
	Alt:: MsgBox, Hemograma!
	Loop
	{
		MsgBox, Caxias!
	}
	NumpadDiv::
	{
		WinGetActiveTitle, winTitle
		If InStr(winTitle, "hemograma 4") {
			;~ MsgBox, Hemograma 4
			MouseClick,, 600, 660
			MouseClick,, 60, 660
		} else {
			;~ MsgBox, Não é Hemograma 4
			MouseClick,, 600, 660
		}
		return
	}

	^NumpadDiv::
		MouseClick,, 975, 660
	return

	NumpadMult::MouseClick,, 815, 1025
	NumpadSub::MouseClick,, 440, 555
	;~ NumpadAdd::enviarLaudo()
	NumpadAdd::
	{
		WinGetActiveTitle, winTitle
		If InStr(winTitle, "Veterinária Jardim Guanabara") or InStr(winTitle, "Cat Para Gatos") or InStr(winTitle, "Tatiane Ferraro"){
			MsgBox, 4,, JARDIM GUANABARA ou CAT! `n`nTEM CAPILAR???
				IfMsgBox Yes
				{
					Sleep 1000
					enviarLaudo()
				}
				else
				{
				}
		} else {
			;~ MsgBox, NÃO É JARDIM NEM CAT
			enviarLaudo()
		}
		return
	}

	^L::
	{
		MouseClick,, 460, 535
		ControlSetText, OKtRichTbx2,, i)hemograma
		Send, {backspace}
		Sleep 500
		return
	}

	PgDn::
	{
		MouseClick,, 950, 630
		ControlSetText, OKtRichTbx2,, i)hemograma,
		Sleep 500
		Send, Presença de Microfilárias na amostra enviada.
		return
	}
	:*:atip::
		InputBox, celulasAtipicas, Insira a quantidade de células atípicas,, Quantidade de Células Atípicas
			if ErrorLevel
				return
			Send % "Presença de manchas de Gumprecht. Presença de " . celulasAtipicas . "% de células atípicas, com relação núcleo:citoplasma alta, anisocitose e anisocariose moderadas, citoplasma discretamente basofílico, cromatina frouxa e ocasionais nucléolos evidentes.{space}"
	return
	:*:gum::Presença de manchas de Gumprecht.{space}
	:*:esf::Presença de ocasionais esferócitos.

	NumpadEnter::

	ControlGetPos, posX, posY,,, OKttbx2, i)hemograma

	if (posY = 436){
		Send, 000
		Send, {Enter}
	} else if (posX = 745) && (posY = 144){
		ControlGetText, leucometriaTotal, OKttbx2, i)hemograma
		if (leucometriaTotal > 50000)
			MsgBox, Atenção!!! Leucometria Total Muito Alta!
		Send {Tab 2}
	} else if (posX = 557) && (posY = 168){
		ControlGetText, basofilo, OKttbx2, i)hemograma
		if (basofilo > 0)
			MsgBox, ATENÇÃO!!! INSERIU BASÓFILO!
			ControlSetText, OKttbx2,, i)hemograma,
		Send {Tab}
	}else {
		Send, {Tab}
	}

    ;~ if (posY = 436){
        ;~ Send, 000
		;~ Send, {Enter}
	;~ } else if {

	;~ } else {
        ;~ Send, {Tab}
    ;~ }
	return

	Right::
		ControlGetPos, posX, posY,,, OKttbx2, i)hemograma
		if (posX = 152)
			MouseClick,, 745, 144, 2
		return

	Left::
		ControlGetPos, posX, posY,,, OKttbx2, i)hemograma
		if (posX > 500) && (posY <= 456)
			MouseClick,, 152, 144, 1
		if (posX = 152)
			MouseClick,, 745, 144, 2
		return

	Up::
	{
		ControlGetPos, posX, posY,,, OKttbx2, i)hemograma
		if (posY = 416) && (posX = 152){
			MouseClick,, 152, 184, 2
	} 	else {
			Send, {Up}
	}
		return
	}
}

;~ #IfWinActive i)Selecionar os Hemoparasitos
	;~ Numpad1::
		;~ MsgBox, Você está na Janela Certa!
	;~ Numpad2::
	;~ Numpad3::
	;~ Numpad4::

#IfWinActive i)resultado
{
	NumpadMult::
	{
		MouseClick,, 100, 520
		Sleep 100
		WinActivate i)hemograma
		return
	}
	NumpadSub::
	{
		MouseClick,, 600, 520
		Sleep 100
		WinActivate i)hemograma
		return
	}
}

#IfWinActive i)giardia|parvovirose|microfilárias
{
	Numpad1::
		if WinActive("i)giardia|parvovirose"){
			MouseClick,, 500, 320
			ControlSetText, OKtRichTbx2,, i)giardia|parvovirose,
			Send, Negativo.
			enviarLaudo()
		} else if WinActive("i)microfilárias"){
			MouseClick,, 500, 320
			ControlSetText, OKtRichTbx2,, i)microfilárias,
			Send, Negativo. Não foram observadas microfilárias na amostra analisada.
			enviarLaudo()
		}
	return

	Numpad2::
		if WinActive("i)giardia|parvovirose"){
			MouseClick,, 500, 320
			ControlSetText, OKtRichTbx2,, i)giardia|parvovirose,
			Send, POSITIVO.
			enviarLaudo()
		} else if WinActive("i)microfilárias"){
			MouseClick,, 500, 320
			ControlSetText, OKtRichTbx2,, i)microfilárias,
			Send, Positivo. Presença de Microfilárias na amostra analisada.
			enviarLaudo()
		}
	return

	Numpad0::enviarLaudo()
}

#IfWinActive i)SISANCLIV
	Numpad1:: MouseClick,, 30, 300

#IfWinActive i)solicitação ;COAGULADOS
{
	Numpad1::
	{
		MouseClick,, 500, 320
		Send, Não foi possível realizar a análise solicitada pois a amostra enviada encontra-se inviável ( Amostra coagulada)`nFavor enviar nova amostra.
		;~ Sleep 500
		enviarLaudo()
		return
	}
	Numpad2::
	{
		MouseClick,, 500, 320
		Send, Não foi possível realizar a análise solicitada pois a amostra enviada encontra-se insuficiente.`nFavor enviar nova amostra.
		;~ Sleep 500
		enviarLaudo()
		return
	}
	Numpad3::
	{
		MouseClick,, 500, 320
		Send, Não foi possível realizar a análise solicitada pois a amostra enviada encontra-se inadequada.`n
		;~ Sleep 500
		enviarLaudo()
		return
	}
	Numpad4::
	{
		MouseClick,, 500, 320
		Send, Não foi possível realizar a análise solicitada pois a amostra enviada encontra-se inadequada.`nRecomenda-se coletar nova amostra por raspado de pele (pelos e crostas).
		;~ Sleep 500
		enviarLaudo()
		return
	}
	Numpad5::
	{
		MouseClick,, 500, 320
		Send, Não foi possível realizar a análise solicitada pois a amostra enviada encontra-se inadequada.`nRecomenda-se coletar nova amostra por swab (rolamento).
		;~ Sleep 500
		enviarLaudo()
		return
	}
	Numpad6::
	{
		MouseClick,, 500, 320
		Send, Não foi possível realizar a análise solicitada pois a amostra enviada encontra-se inadequada.`nRecomenda-se coletar nova amostra por imprint direto.
		;~ Sleep 500
		enviarLaudo()
		return
	}
	Esc::
	{
		while GetKeyState("Esc")
		{
			ToolTip 1=Coagulado`n2=Amostra Insuficiente`n3=Amostra Inadequada`n4=Sarfu Inadequado`n5=Oto na Fita`n6=Esporo na Fita
		}
		ToolTip
		return
	}
}

#IfWinActive i)eas
{
	;~ Tab::
	;~ {
		;~ ControlGetPos, posiX, posiY,,, OKttbx2, i)eas
		;~ if (posiX = 650) && (posiY = 251) {
			;~ MouseClick,, 10, 526, 3
			;~ MsgBox, Está no Controle Certo! %posiX% %posiY%
	;~ } 	else {
			;~ Send, {Down}
			;~ MsgBox, NÃO está no Controle Certo! %posiX% %posiY%
	;~ }
		;~ return
	;~ }

	;~ Tab::
		;~ ControlGetPos, posiX, posiY,,, OKttbx2, i)eas
		;~ MsgBox, X(%posiX%)`tY(%posiY%)
	;~ return

	NumpadSub::enviarLaudo()

	:*:e4::Estruvita (raros)`n
	:*:e1::Estruvita ({+})`n
	:*:e2::Estruvita ({+}{+})`n
	:*:e3::Estruvita ({+}{+}{+})`n
	:*:e5::Estruvita (presentes)`n
	:*:e6::Estruvita (frequentes)`n

	:*:o4::Oxalato de Cálcio (raros)`n
	:*:o1::Oxalato de Cálcio ({+})`n
	:*:o2::Oxalato de Cálcio ({+}{+})`n
	:*:o3::Oxalato de Cálcio ({+}{+}{+})`n
	:*:o5::Oxalato de Cálcio (presentes)`n
	:*:o6::Oxalato de Cálcio (frequentes)`n

	:*:b4::Bilirrubina (raros)`n
	:*:b1::Bilirrubina ({+})`n
	:*:b2::Bilirrubina ({+}{+})`n
	:*:b3::Bilirrubina ({+}{+}{+})`n
	:*:b5::Bilirrubina (presentes)`n
	:*:b6::Bilirrubina (frequentes)`n

	:*:u4::Uretrais (raras)`n
	:*:u1::Uretrais ({+})`n
	:*:u2::Uretrais ({+}{+})`n
	:*:u3::Uretrais ({+}{+}{+})`n
	:*:u5::Uretrais (presentes)`n
	:*:u6::Uretrais (frequentes)`n

	:*:v4::Vesicais (raras)`n
	:*:v1::Vesicais ({+})`n
	:*:v2::Vesicais ({+}{+})`n
	:*:v3::Vesicais ({+}{+}{+})`n
	:*:v5::Vesicais (presentes)`n
	:*:v6::Vesicais (frequentes)`n

	:*:p4::Pélvicas (raras)`n
	:*:p1::Pélvicas ({+})`n
	:*:p2::Pélvicas ({+}{+})`n
	:*:p3::Pélvicas ({+}{+}{+})`n
	:*:p5::Pélvicas (presentes)`n
	:*:p6::Pélvicas (frequentes)`n

	:*:g4::Granulosos (raros)`n
	:*:g1::Granulosos ({+})`n
	:*:g2::Granulosos ({+}{+})`n
	:*:g3::Granulosos ({+}{+}{+})`n
	:*:g5::Granulosos (presentes)`n
	:*:g6::Granulosos (frequentes)`n

	:*:h4::Hialinos (raros)`n
	:*:h1::Hialinos ({+})`n
	:*:h2::Hialinos ({+}{+})`n
	:*:h3::Hialinos ({+}{+}{+})`n
	:*:h5::Hialinos (presentes)`n
	:*:h6::Hialinos (frequentes)`n

	:*:lp::Sedimentoscopia prejudicada devido à intensa hematúria.`n
	:*:im::Impregnação por bilirrubina.`n
	:*:gg::Presença de gotículas de gordura.`n
	:*:aw::Ausentes`n
	:*:acv::Presença de aglomerados de células vesicais.`n
	:*:=::{+}

VerificarPosicaoEEnviar(teclas, hotkey) {
    ControlGetPos, posX, posY,,, OKttbx2, i)eas
    if (posY >= 293 && posY <= 419) {
        SendInput, %teclas%
		Sleep 100
        SendInput, {Enter}
    } else {
        ;~ SendInput, {NumpadEnter}
		SendInput, {%hotkey%}
    }
}

NumpadEnter::
	VerificarPosicaoEEnviar("N", "NumpadEnter")
	return

Numpad1::
	VerificarPosicaoEEnviar("{+}", "Numpad1")
	return

Numpad2::
	VerificarPosicaoEEnviar("{+}{+}", "Numpad2")
	return

Numpad3::
	VerificarPosicaoEEnviar("{+}{+}{+}", "Numpad3")
	return

Numpad4::
	VerificarPosicaoEEnviar("t", "Numpad4")
	return

}

#IfWinActive i)fezes
{
	Numpad1::
	{
		MouseClick,, 970, 190
		Send, Não foram observados ovos ou cistos/oocistos na amostra examinada.
		enviarLaudo()
		return
	}
	Numpad2::
	{
		MouseClick,, 970, 190
		Send, Cistos de Giardia sp.
		enviarLaudo()
		return
	}
	Numpad3::
	{
		MouseClick,, 970, 190
		Send, Coccídios.
		enviarLaudo()
		return
	}
	Numpad4::
	{
		MouseClick,, 970, 190
		Send, Cystoisospora spp.
		enviarLaudo()
		return
	}
	Numpad5::
	{
		MouseClick,, 970, 190
		Send, Ovos de Ancylostoma spp.
		enviarLaudo()
		return
	}
	Numpad6::
	{
		MouseClick,, 970, 190
		Send, Ovos de Toxocara sp.
		enviarLaudo()
		return
	}
	Esc::
	{
		while GetKeyState("Esc")
		{
			ToolTip 1=Negativo`n2=Giardia`n3=Coccídio`n4=Isospora`n5=Ancylostoma`n6=Toxocara
		}
		ToolTip
		return
	}
}

#IfWinActive i)esporotricose
{
	Numpad1::
	{
		MouseClick,, 700, 300
		Send, Positivo.`nForam observadas estruturas leveduriformes compatíveis com o fungo Sporothrix schenckii.
		enviarLaudo()
		return
	}
	Numpad2::
	{
		MouseClick,, 700, 300
		Send, Negativo. Não foram observadas estruturas fúngicas na amostra enviada.
		enviarLaudo()
		return
	}
}

#IfWinActive i)sarna|fungo
{
	Numpad1::
		MouseClick,, 240, 325
		Send, Não foram observadas estruturas fúngicas na amostra enviada.
		return
	Numpad2::
		MouseClick,, 710, 325
		Send, Não foram observados ácaros na amostra examinada.
		return
	Numpad0::
	{
		;~ MouseClick,, 600, 1015
		enviarLaudo()
		return
	}
}

#IfWinActive i)hematozoários
{
	Numpad1::
	{
		MouseClick,, 600, 300
		Send, Não foram observados hemoparasitos na amostra analisada.
		enviarLaudo()
		return
	}
}

#If WinActive("i)citologia") && !WinActive("i)fecal")
	Numpad1::MouseClick,, 750, 350
	Numpad2::MouseClick,, 600, 42
	NumpadAdd::
		;~ if WinExist("i)citologia") {
		WinActivate i)citologia
		MouseClick,, 600, 650
return

#IfWinActive i)HF-LAB
	NumpadAdd::
		if WinExist("i)citologia") && !WinActive("i)citologia"){
			;~ MsgBox Existe!
			MouseClick,, 30, 520
			WinActivate i)citologia
		} else {
			SendInput {NumpadAdd}
		}
return