# BetterDLL

## Vantagens

- Oculta a complexidade da interação direta com as APIs do sistema operacional. Isso é especialmente útil para desenvolvedores menos experientes ou aqueles que não estão familiarizados com os detalhes internos das APIs.

- Oferece opções de configuração flexíveis, como a capacidade de definir o diretório onde a DLL está localizada. Isso permite trabalhar com várias DLLs de terceiros e personalizar o comportamento conforme necessário.

- Gerencia automaticamente a alocação e liberação de memória ao carregar e descarregar bibliotecas dinâmicas (DLLs), evitando vazamentos de memória e outros problemas relacionados.

## Create

> Método para iniciar a dll.

### Exemplo de uso

```vb
Dim exemplo As New DLL, exemplo2 As New DLL
exemplo.Create("user32.dll") ' DLL nativa.
exemplo2.Create "math.dll", "C:\Users\manteiga32\Desktop" ' DLL externa.
```

## Add

> Método para adicionar funções.

### Exemplo de uso

```vb
Dim exemplo As New DLL
exemplo.Create("user32")
exemplo.Add "MessageBoxA", vbLong ' Adicionamos a função "MessageBoxA" e colocamos que ela retorna o tipo "Long".
```

## Run

> Função que executa uma função da DLL adicionada e existente.

### Exemplo de uso

```vb
exemplo.Run "MessageBoxA", Array(0, "Texto", "Título", vbOKOnly) ' Função com parâmetros.
exemplo2.Run("GetTickCount") ' Função sem parâmetros.
```

## Código de exemplo em um loop

```vb
Private Sub Main()
    Dim user32 As New DLL
    user32.Create ("user32.dll")
    user32.Add "MessageBoxA", vbLong
    user32.Add "GetAsyncKeyState", vbInteger
    
    Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1
        If user32("GetAsyncKeyState", Array(vbKeyZ)) Then
            ' Você pode usar tanto o .Run ou apenas no formato usado nesse if.
            user32.Run "MessageBoxA", Array(0, "Você apertou a tecla Z", "Título", vbOKOnly)
        End If
        Shapes("timer").TextEffect.text = Timer
        DoEvents
    Loop
End Sub
```