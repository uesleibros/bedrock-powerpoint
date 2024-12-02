# Discord RPC

O **Discord RPC** permite integrar seu aplicativo VBA com o Discord, atualizando o status de presença do usuário no Discord, como "Jogando", "Assistindo" e outros. A comunicação é feita por meio de pipes do Windows, e o status pode ser personalizado.

## Considerações

- Você precisa usar o [**JSONVBA**](https://pptgamespt.wixsite.com/pptg-coding/json) da PPTGames ou adaptar o código para compatibilidade com outras bibliotecas.
- O método de autorização é automático ao chamar o método `Connect`.
- O Discord deve estar em execução para o RPC funcionar corretamente.

## Exemplos

### Exemplo Básico de Conexão e Atualização de Status

```vb
Dim rpc As New DiscordRPC

Public Sub configurar()
    rpc.Connect "1298810058746105866", application
    rpc.Update "in-game", state:="Jogando", details:="No Mapa 1"
End Sub
```

### Exemplo de Alteração de Status

```vb
Public Sub alterar_presenca()
    rpc.Update "in-game", state:="Jogando", details:="No Mundo 1", party_size:=5
End Sub
```

> Detalhe que o único parâmetro obrigatório dentre esses usados é primeiro (activity). Os outros são opcionais e você pode usar para mudar os valores dinamicamente sem precisar criar um milhão de arquivos.json pra cada caso.

### Exemplo de Desconexão

```vb
Public Sub desconectar()
    rpc.Disconnect
End Sub
```

---

## Diferenças Entre os Métodos

- **`Connect`**: Conecta o seu aplicativo ao Discord, passando o `client_id` e uma referência do aplicativo.
- **`Update`**: Atualiza o status de presença com base nas atividades (e.g., jogo ou música), e pode incluir informações como estado, detalhes, imagens e números de pessoas em uma "party".
- **`Disconnect`**: Desconecta o aplicativo do Discord e limpa o status.

---

## Estrutura do Arquivo de Atividade

A atividade é definida em um arquivo JSON, localizado em uma pasta chamada **rpc/activities**. Exemplo de uma atividade:

```json
{
  "state": "Jogando",
  "details": "No Mapa 1",
  "assets": {
    "large_image": "game_logo",
    "large_text": "Jogando o jogo incrível",
    "small_image": "avatar",
    "small_text": "Jogando com amigos"
  },
  "party": {
    "size": [5, 10]
  }
}
```

---

### Considerações Finais

- **Autorização e `client_id`**: O `client_id` deve ser obtido no [Portal de Desenvolvedores do Discord](https://discord.com/developers/applications).
- **Requisitos de Comunicação**: O Discord deve estar em execução no sistema para que a comunicação funcione.
- **Biblioteca JSON**: Use o JSONVBA da PPTGames ou modifique o código para outras bibliotecas.

---
