# Temporizer

Facilita o controle de tempo e execução de tarefas com temporizadores avançados em VBA.

## Exemplo de Código

> Dica, use sempre a função `ClearIntervals` antes de fazer o Wait, para garantir que todos os intervalos estejam devidamente resetados.

```vb
If Wait(1, "A cada 1 segundo faça x coisa") Then
	' coiso
End If
```