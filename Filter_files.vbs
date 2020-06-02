Dim pastaOrigem,words,pastaDestino,keywords,tt,wordsTemp,pastaDestinoTemp,wordsTemp2,pastaDestinoTemp2 'declarando as variaveis
Set oFS = createObject("Scripting.FileSystemObject") 'criando o objeto que vai navegar nas pastas
wordsTemp=""
pastaDestinoTemp=""

Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory

pastaOrigem = strCurDir

'pastaOrigem=InputBox("Escolha a pasta de origem (caminho absoluto):","Filter Files") 'armazenando o caminho de origem

dim noStop
noStop = true

if(pastaOrigem="")then
	dim junk
	junk = msgBox("pasta de origem não pode ser vazio",0,"Error")
else
	while tt<>7 and noStop 'loop que permite o usuario colocar quantos filtros quiser
		wordsTemp2 = InputBox("Digite as palavras-chaves separadas por virgula:","Filter Files") 'as variaveis wordsTemp2 e pastaDestinoTemp2 são usadas
		pastaDestinoTemp2 = BrowseFolder(strCurDir,false)
		'pastaDestinoTemp2 = InputBox("Escolha a pasta de destino (caminho absoluto):","Filter Files") 'como exception
		wordsTemp = wordsTemp2 & chr(0) & wordsTemp 
		pastaDestinoTemp = pastaDestinoTemp2 & chr(0) & pastaDestinoTemp
		if(wordsTemp2="" or pastaDestinoTemp2="")then
			noStop = false
			'msgbox("nostop=false")
		else
			tt = msgBox("Deseja adicionar mais filtros e destinos?",4)	'a opção 4 do msgbox, disponibiliza os botoes yes(6) e no(7)				
		end if
	wend
	if(noStop=false)then		
		junk = msgBox("as palavras-chaves ou pasta de destino não podem ser vazios",0,"Error")
	else
		wordsTemp = split(wordstemp,chr(0),-1,1) 'quebra a string em array
		pastaDestino = split(pastaDestinoTemp,chr(0),-1,1)
		dim a,b
		a=0 	'indice das pastas de destino
		b=true	'faz a função do break no laÃ§o abaixo

		For each words in wordsTemp
			keywords = Split(words,",",-1,1)
			For Each File in oFS.GetFolder(pastaOrigem).Files
				'msgBox(oFS.GetAbsolutePathName(File))
				Dim word
				For Each word in keywords
					if(b)then
						'msgBox(word)
						word = Trim(word)
						'msgBox(pastaDestino(a) & "\" & oFS.GetFileName(File))		
						if(Instr(1,oFS.GetFileName(File),word,1)>0)then
							if(Not(Instr(1,oFS.GetFileName(File),".vbs",1)>0))then
								oFS.MoveFile File, pastaDestino(a) & "\" & oFS.GetFileName(File)
								b=false
							end if
						end if
					end if
				Next
				b=true
			Next
			a=a+1
		Next
		MsgBox("Fim da execução")
	end if
end if
	
Function BrowseFolder( myStartLocation, blnSimpleDialog )
' This function generates a Browse Folder dialog
' and returns the selected folder as a string.
'
' Arguments:
' myStartLocation   [string]  start folder for dialog, or "My Computer", or
'                             empty string to open in "Desktop\My Documents"
' blnSimpleDialog   [boolean] if False, an additional text field will be
'                             displayed where the folder can be selected
'                             by typing the fully qualified path
'
' Returns:          [string]  the fully qualified path to the selected folder
'
' Based on the Hey Scripting Guys article
' "How Can I Show Users a Dialog Box That Only Lets Them Select Folders?"
' http://www.microsoft.com/technet/scriptcenter/resources/qanda/jun05/hey0617.mspx
'
' Function written by Rob van der Woude
' http://www.robvanderwoude.com
    Const MY_COMPUTER   = &H11&
    Const WINDOW_HANDLE = 0 ' Must ALWAYS be 0

    Dim numOptions, objFolder, objFolderItem
    Dim objPath, objShell, strPath, strPrompt

    ' Set the options for the dialog window
    strPrompt = "Select a destination folder:"
    If blnSimpleDialog = True Then
        numOptions = 0      ' Simple dialog
    Else
        numOptions = &H10&  ' Additional text field to type folder path
    End If
    
    ' Create a Windows Shell object
    Set objShell = CreateObject( "Shell.Application" )

    ' If specified, convert "My Computer" to a valid
    ' path for the Windows Shell's BrowseFolder method
    If UCase( myStartLocation ) = "MY COMPUTER" Then
        Set objFolder = objShell.Namespace( MY_COMPUTER )
        Set objFolderItem = objFolder.Self
        strPath = objFolderItem.Path
    Else
        strPath = myStartLocation
    End If

    Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, _
                                              numOptions, strPath )

    ' Quit if no folder was selected
    If objFolder Is Nothing Then
        BrowseFolder = ""
        Exit Function
    End If

    ' Retrieve the path of the selected folder
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path

    ' Return the path of the selected folder
    BrowseFolder = objPath
End Function
