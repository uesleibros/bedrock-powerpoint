# JSONF (FAST)

## Visão Geral

O módulo `JSONF` foi desenvolvido com pesquisas e bases em repositórios do Github de parsers para a linguagem de marcação JSON. Com esse estudo, conseguimos fazer o parser mais rápido para VBA com todas as funcionalidades que você tem direito.

## Funções Públicas

### `parse(json_string As String) As Variant`

A função `parse` converte uma string JSON em um objeto VBA (seja um `Dictionary` para objetos JSON ou um `List` para arrays JSON).

#### Parâmetros:

- **json_string** (String): A string que contém os dados JSON a serem convertidos.

#### Retorno:

- Retorna um objeto VBA: pode ser um `Dictionary` (para objetos JSON) ou um `List` (para arrays JSON). Se a string JSON estiver vazia ou mal formada, a função gera um erro.

#### Detalhes:

A função remove caracteres de espaço, quebras de linha e tabulação da string JSON antes de processá-la. Ela analisa cada caractere da string, tratando corretamente os valores entre aspas (strings), valores numéricos, booleanos (`true`/`false`), `null`, objetos (`{}`) e arrays (`[]`). A função retorna um `Dictionary` para objetos JSON e uma `List` para arrays JSON.

#### Exemplo de Uso:

```vb
Dim jsonString As String
Dim parsedObject As Variant

jsonString = "{""nome"": ""João"", ""idade"": 30, ""cidade"": ""São Paulo""}"
Set parsedObject = parse(jsonString)

' Acessando os dados parseados
Debug.Print parsedObject("nome")  ' Saída: João
Debug.Print parsedObject("idade") ' Saída: 30
```

### `stringify(json_object As Object) As String`

A função `stringify` converte um objeto VBA (como `Dictionary` ou `List`) em uma string JSON.

#### Parâmetros:

- **json_object** (Object): O objeto VBA (pode ser um `Dictionary` ou `List`) a ser convertido em JSON.

#### Retorno:

- Retorna uma string representando o objeto como um JSON.

#### Detalhes:

Esta função verifica o tipo do objeto VBA fornecido e gera uma string JSON correspondente. Se o objeto for um `Dictionary`, ele é convertido em um objeto JSON (`{}`), e se for um `List`, ele é convertido em um array JSON (`[]`). A função lida corretamente com diferentes tipos de dados, como strings, números, booleanos e objetos aninhados.

#### Exemplo de Uso:

```vb
Dim jsonObject As Dictionary
Dim jsonString As String

Set jsonObject = New Dictionary
jsonObject.Add "nome", "João"
jsonObject.Add "idade", 30
jsonObject.Add "cidade", "São Paulo"

jsonString = stringify(jsonObject)

' Saída: {"nome": "João", "idade": 30, "cidade": "São Paulo"}
Debug.Print jsonString
```

---

## Observações Importantes:

- A função `parse` gera um erro se a string JSON fornecida estiver vazia ou mal formada.
- A função `stringify` pode ser utilizada para gerar uma string JSON a partir de qualquer objeto que implemente a interface `Dictionary` ou `List`.
- O módulo trata automaticamente a formatação de strings JSON, incluindo escape de caracteres especiais e controle de indentação para facilitar a leitura.
- Não é necessário opções pra habilitar palavras sem aspas e etc, nosso JSON faz tudo e reconhece os problemas por si só e arruma.
