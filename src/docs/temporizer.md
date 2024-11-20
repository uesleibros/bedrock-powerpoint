# Temporizer

Facilita o trabalho de contabilizar o tempo para executar tarefas.

## Considerações

- É necessário que você esteja utilizando o **Temporizer** dentro de um ciclo (looping), para que assim as checagens dos intervalos funcione apropriadamente.
- Sempre antes de começar a usar, é recomendado utilizar o método `ClearIntervals` para limpar todos os intervalos anteriores. Útil caso esteja usando ele em mais de 1 slide, para não haver conflito com os nomes que podem ser usados igualmente em ambos os slides.

## Exemplos

### Temporizador Contínuo

```vb
Public Sub Exemplo()
	ClearIntervals ' limpamos todos os intervalos para sumir com o cache.

	' Enquanto o slide atual estiver visível, o temporizador acionará uma mensagem a cada 5 segundos.
	Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = Me.SlideNumber
		If Wait(5, "Mostrar mensagem na tela") Then
			MsgBox "Olá."
		End If
		DoEvents
	Loop
End Sub
```

### Temporizador Único

```vb
Public Sub Exemplo()
	ClearIntervals ' limpamos todos os intervalos para sumir com o cache.

	' Enquanto o slide atual estiver visível, o temporizador acionará uma mensagem após 5 segundos apenas uma vez.
	Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = Me.SlideNumber
		If Wait(5, "Mostrar mensagem na tela", True) Then
			MsgBox "Olá."
		End If
		DoEvents
	Loop
End Sub
```

## Diferenças entre os dois tipos

- **Temporizador Contínuo**
- - Pode ser acionado repetidamente em intervalos especificados durante a execução do código.
- - Ideal para ações que necessitam ocorrer continuamente.

- **Temporizador Único**
- - Aciona a ação uma única vez após o intervalo especificado e não é repetido.
- - Útil para execuções pontuais, como configurações de inicialização.