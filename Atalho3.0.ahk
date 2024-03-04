#SingleInstance Force
#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode RegEx

F10::Suspend
^SPACE::
{
	Winset, Alwaysontop, , A
	return
}

definirNumero() {
	global numSup := 753
}

#+s::Send ^{Insert} ;NÃO PRECISA NO SEGUNDO PC

Esc:: ;Calcular posição do mouse
{
	while GetKeyState("Esc")
	{
		MouseGetPos, xPos, yPos, WhichWindow, WhichControl, 1
		ControlGetPos, x, y, w, h, %WhichControl%, ahk_id %WhichWindow%
		;~ ToolTip, %WhichControl%`nX%X%`tY%Y%`nW%W%`t%H%
		;~ MouseGetPos xPos, yPos,, classControl, 1
		WinGetActiveTitle winTitle
		ToolTip X(%xPos%)`tY(%yPos%)`n%winTitle%`n%WhichWindow%`nNome do Controle: %WhichControl%`nX%X%`tY%Y%`nW%W%`t%H%
	}
	ToolTip
	return
}

nomeUsuario := A_UserName

^+!k::
    Run, explorer.exe "C:\Users\%nomeUsuario%\Pictures\Screenshots"
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


;~ /*
valoresHemato:
{
	;~ global hematocrito := Round(++volumeGlobular, 0)
	if (volumeGlobular > 15) {
        global hematocrito := Round(++volumeGlobular, 0)
    } else {
        global hematocrito := Round(volumeGlobular, 0)
    }
	global hemaciaCao := StrReplace(Round(hematocrito/7, 2), ".", ",")
	global hemaciaGato := StrReplace(Round(hematocrito/5, 2), ".", ",")
	global hemoglobinaAll := StrReplace(Round(hematocrito/3, 2), ".", ",")
	return
}

F3::
{
	;~ valoresHemato()
	while GetKeyState("F3") {
		ToolTip,
	(Ltrim
	O hematócrito é %hematocrito%`%
	A hematimetria de CÃO é %hemaciaCao%
	A hematimetria de GATO é %hemaciaGato%
	A hemoglobina é %hemoglobinaAll%
	)
	}
	ToolTip
	return
}


End::
{
	ControlGetPos,, ctrlY,,, OKttbx2, i)hemograma
	if (ctrlY = 184){
		ControlGetText, volumeGlobular, OKttbx2, i)hemograma
		;~ valoresHemato()
		Send {Up 2}
		goto valoresHemato
		;~ MsgBox, Você está no Hematócrito
	} else if (ctrlY = 164){
		;~ MsgBox, Você está na Hemoglobina
		ControlSetText, OKttbx2,, i)hemograma
		;~ Sleep 100
		Send, %hemoglobinaAll%
		Send, {Enter}
	} else if (ctrlY = 144){
		;~ MsgBox, Você está na Hematimetria
		ControlSetText, OKttbx2,, i)hemograma
		;~ Sleep 100
		Send, %hemaciaCao%
		Send {Enter}
	}
	return
}

Home::
{
	ControlGetPos,, ctrlY,,, OKttbx2, i)hemograma
	if (ctrlY = 184){
		ControlGetText, volumeGlobular, OKttbx2, i)hemograma
		;~ valoresHemato()
		Send {Up 2}
		goto valoresHemato
		;~ MsgBox, Você está no Hematócrito
	} else if (ctrlY = 164){
		;~ MsgBox, Você está na Hemoglobina
		ControlSetText, OKttbx2,, i)hemograma
		;~ Sleep 100
		Send, %hemoglobinaAll%
		Send, {Enter}
	} else if (ctrlY = 144){
		;~ MsgBox, Você está na Hematimetria
		ControlSetText, OKttbx2,, i)hemograma
		;~ Sleep 100
		Send, %hemaciaGato%
		Send {Enter}
	}
	return
}

enviarLaudo() {
	if WinActive("i)hemograma") && !WinActive("Solicitação de Nova Amostra"){
		;~ Send #{PrintScreen}
		;~ Sleep 500
		;~ MouseClick,, 600, 1020
		MouseClick,, 960, 1020 ;só digita sem enviar
		;~ MouseClick,, 600, 1020 ;normal
	} else if WinActive("Solicitação de Nova Amostra"){
		;~ Sleep 500
		;~ MouseClick,, 960, 1020
		MouseClick,, 600, 1020
	} else if WinActive("i)eas") or WinActive("i)fezes") or WinActive("i)microfilárias") or WinActive("i)giardia") or WinActive("i)hematozoários") or WinActive("i)sarna|fungo"){
		;~ Sleep 500
		MouseClick,, 600, 1020 ;normal
		;~ MouseClick,, 960, 1020 ;só digita
	} else if WinActive("i)esporotricose"){
		MouseClick,, 620, 650 ;envia
		;~ MouseClick,, 1000, 650 ;digita
	}
}

laudosConferidos() {
    ; Especifique os caminhos dos diretórios
    origem := "C:\Users\%nomeUsuario%\Pictures\Screenshots\*.*" ; Coloque o caminho da pasta de origem
    ;~ destino := "C:\Users\LABORATÓRIO 01\Pictures\Enviados\1107" ; Coloque o caminho da pasta de destino

	;Obtém a data atual no formato MMDD
    FormatTime, data, %A_Now%, MMdd

	; Cria o caminho da pasta de destino com a data atual
	destino := "C:\Users\%nomeUsuario%\Pictures\Enviados\" . data

	; Verifica se a pasta de destino existe, se não, cria ela
    if (!FileExist(destino))
        FileCreateDir, %destino%

    ; Garante que os caminhos sejam válidos
    if (FileExist(origem) && FileExist(destino)) {
        ; Move os arquivos
        FileMove, %origem%, %destino%, 1
        MsgBox, Os arquivos foram movidos com sucesso.
    } else {
        MsgBox, O caminho de origem ou destino é inválido.
    }
}

ativarJanela() {
	WinActivate i)informações
	WinActivate i)resultado
	WinActivate i)atenção
	WinActivate i)HF-LAB
	return
}

^g::laudosConferidos()

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

#IfWinActive i)número|realizados
	global numState := true
	NumpadAdd::
	{
		ControlGetFocus, controleFocado
		ControlSetText, %controleFocado%,,
		if (numState = true) {
			definirNumero()
			Send %numSup%
			numState := false
		} else {
			Send % --numSup
			numState := true
		}
		return
	}
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
	WheelUp::MouseClick,, 300, 520
	NumpadDiv::Send ICTERÍCIA
	NumpadMult::Send LIPEMIA
	NumpadSub::Send ?
	!n::MouseClick,, 50, 1020
	^a::Send, +{Home}
	F5::
		FormatTime, dataHora, %A_Now%, dd/MM/yyyy HH:mm
		Send, (%dataHora%)
		return
}

#IfWinActive i)informações
{
	NumpadEnter::MouseClick,, 480, 250
}

#IfWinActive i)hemograma
{
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
	:*:atip::Presença de manchas de Gumprecht. Presença de 19% de células atípicas, com relação núcleo:citoplasma alta, anisocitose e anisocariose moderadas, citoplasma discretamente basofílico, cromatina frouxa e ocasionais nucléolos evidentes.{space}
	:*:gum::Presença de manchas de Gumprecht.{space}
	:*:esf::Presença de ocasionais esferócitos.

	NumpadEnter::
	{
	ControlGetPos, posX, posY,,, OKttbx2, i)hemograma
    if (posY = 436){
        Send, 000
		Send, {Enter}
    ;~ } else if (posY = 184){ ;do hematócrito para metas
		;~ MouseClick,, 152, 416, 2
		;~ Send, 0
	;~ } else if (posY = 144) && (posX = 745){ ;da leucometria para eosinofilo
		;~ MouseClick,, 557, 188, 2
	;~ } else if (posY = 188) && (posX = 557){ ;do eosinofilo para neutrofilo
		;~ MouseClick,, 557, 268, 2
	} else {
        Send, {Tab}
    }
	return
	}

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

#IfWinActive i)microfilárias
{
	Numpad1::
		MouseClick,, 500, 320
		ControlSetText, OKtRichTbx2,, i)microfilárias,
		;~ Sleep 500
		Send, Negativo. Não foram observadas microfilárias na amostra analisada.
		enviarLaudo()
		return
	Numpad2::
		MouseClick,, 500, 320
		ControlSetText, OKtRichTbx2,, i)microfilárias,
		;~ Sleep 500
		Send, Positivo. Presença de Microfilárias na amostra analisada.
		enviarLaudo()
		return
	Numpad0::enviarLaudo()
}

#IfWinActive i)giardia
{
	Numpad1::
		MouseClick,, 500, 320
		ControlSetText, OKtRichTbx2,, i)giardia,
		;~ Sleep 500
		Send, Negativo.
		enviarLaudo()
		return
	Numpad2::
		MouseClick,, 500, 320
		ControlSetText, OKtRichTbx2,, i)giardia,
		;~ Sleep 500
		Send, POSITIVO.
		enviarLaudo()
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
		Send, Não foi possível realizar a análise solicitada pois a amostra enviada encontra-se inadequada.`nRecomenda-se coletar nova amostra por método de imprint direto.
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
			ToolTip 1=Coagulado`n2=Amostra Insuficiente`n3=Dermato Na Fita`n4=Sarfu Inadequado`n5=Oto na Fita`n6=Esporo na Fita
		}
		ToolTip
		return
	}
}

#IfWinActive i)eas
{
	F2::
	{
		ControlFocus,, ahk_id 0x13407b2
		return
	}

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

#IfWinActive i)otológica
{
	NumpadSub::MouseClick,, 960, 430
}
