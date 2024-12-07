# Documentação do Parser XON (eXtensible Object Notation)

> Fase Experimental

XON (eXtensible Object Notation) é uma linguagem projetada para estruturar dados hierárquicos em um formato legível e flexível. Esta linguagem é ideal para representar objetos complexos como dicionários, listas, e valores primitivos, frequentemente usados em aplicações de configuração, jogos, e modelagem de dados. 

O módulo apresentado implementa um **parser** e um **serializer** para manipular dados XON em VBA. Ele converte texto XON em estruturas manipuláveis no VBA (listas e dicionários) e permite exportar novamente para o formato XON.

> Antes de usar, é necessário utilizar o módulo [List](https://pptgamespt.wixsite.com/pptg-coding/better-arrays) da PPTGames, sem ele, dará erros indesejados.

## Características do XON

1. **Chaves e Valores:**
   - Atribuições são feitas com `->`.
   - Chaves podem ser strings entre aspas (`"key"`) ou diretamente (`key`).

2. **Blocos e Listas:**
   - Blocos são definidos por colchetes `[ ]`.
   - Listas são definidas por parênteses `( )`.

3. **Dados Hierárquicos:**
   - Suporta estruturação recursiva e complexa.

4. **Suporte a Tipos Primitivos:**
   - Strings (`"texto"` ou `'texto'`).
   - Números (`inteiros` e `decimais`).
   - Valores booleanos (`true`, `false`).
   - Valor nulo (`null`).

## Exemplo de XON

```perl
game -> [
  title -> "Aventura de Makoto"
  author -> "FoxyBR_123"
  version -> "1.0.0"

  player -> [
    name -> "Makoto Hoshino"
    class -> "Guerreiro"
    level -> 5
    health -> 320
    stats -> [
      strength -> 18
      agility -> 12
      intelligence -> 9
    ]
    inventory -> (
      [ name -> "Espada de Aço" type -> "Arma" damage -> 35 ]
      [ name -> "Poção de Cura" type -> "Consumível" effect -> "cura" amount -> 50 ]
    )
  ]

  map -> [
    name -> "Floresta da Amazônia Imperial"
    difficulty -> "Média"
    regions -> (
      [ name -> "Clareira" type -> "Ponto Seguro" ]
      [ name -> "Caverna Escura" type -> "Zona de Perigo" ]
    )
  ]

  enemies -> (
    [ name -> "Goblin" type -> "Inimigo" level -> 3 health -> 120 attack -> 10 ]
    [ name -> "Dragão" type -> "Chefe" level -> 10 health -> 2000 attack -> 150 ]
  )
]
```

## Descrição do Parser

### Objetivo

O parser XON converte o texto da linguagem XON em objetos VBA como **listas** e **dicionários**, possibilitando manipulação direta em código.

### Componentes do Parser

1. **Função `Parse(Code As String) As Object`**
   - Ponto de entrada para o parser.
   - Recebe uma string XON e devolve a estrutura correspondente.
   - **Exemplo:**
     ```vb
     Dim data As Object
     Set data = Parse(XONString)
     Debug.Print data("game")("title")  ' Aventura de Makoto
     ```

2. **Função `Stringify(Value As Object) As String`**
   - Serializa uma estrutura VBA (listas ou dicionários) de volta ao formato XON.
   - **Exemplo:**
     ```vb
     Dim XONCode As String
     XONCode = Stringify(data)
     Debug.Print XONCode
     ```

3. **Suporte a Tokens Especiais**
   - Definidos na constante `SYMBOLS`: `[ ] ( ) ->`.

---

## Erros e Tratamento

O parser é projetado para lidar com erros de sintaxe e fornecer mensagens detalhadas, incluindo a linha e coluna onde o problema foi detectado.

### Erros Comuns

1. **Fim Inesperado do Input:**
   - Exemplo: 
     ```
     Error at line 3, column 15: Unterminated string
     ```

2. **Falta de Fechamento:**
   - Exemplo: 
     ```
     Error at line 8, column 1: Unterminated block, missing ']'
     ```

## Exemplo de Uso

### Parser

```vb
Dim XONString As String
XONString = "game -> [ title -> ""Makoto Adventure"" version -> ""1.0.0"" ]"

Dim Data As Object
Set Data = Parse(XONString)

Debug.Print Data("game")("title")  ' Output: Makoto Adventure
```

### Serializer

```vb
Dim XONCode As String
XONCode = Stringify(Data)
Debug.Print XONCode
```

---

## Considerações

O parser XON combina a simplicidade de uso com suporte a estruturas avançadas, tornando-o ideal para projetos que necessitam de manipulação e configuração de dados complexos no VBA. Com suporte extensivo a validações, ele oferece segurança e confiabilidade na análise e na geração de código XON.
