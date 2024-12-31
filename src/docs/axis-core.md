# AxisCore

O módulo AxisCore é uma biblioteca para VBA que permite a interação com controles de jogos (joysticks, gamepads) conectados ao computador. Ele oferece funcionalidades para detectar dispositivos conectados, ler estados dos botões, direcionais e analógicos.

## Inicialização

Para começar a usar o módulo, é necessário inicializá-lo:

```vba
Public Sub Initialize()
```
Esta função deve ser chamada antes de qualquer outra operação com o módulo. Ela detecta e inicializa todos os joysticks conectados.

## Estrutura Principal

### Joysticks

```vba
Public Joysticks As JoystickCollection
```
Variável global que contém informações sobre todos os joysticks conectados. A estrutura possui:

- `Devices(MAX_JOYSTICKID)`: Array de dispositivos com as seguintes propriedades para cada joystick:
  - `ID`: Identificador único do dispositivo
  - `Name`: Nome do dispositivo
  - `Deadzone`: Zona morta para os analógicos (padrão: 2000)
  - `Connected`: Status de conexão
  - `State`: Estado atual do dispositivo

- `TotalConnected`: Número total de joysticks atualmente conectados

> Veja um exemplo de uma das formas de se utilizar o Joysticks
> 
```vb
' Para saber quantos joysticks estão conectados
Debug.Print Joysticks.TotalConnected

' Para obter o nome do primeiro joystick
Debug.Print Joysticks.Devices(0).Name

' Para verificar se um joystick específico está conectado
Debug.Print Joysticks.Devices(1).Connected

' Para ajustar a zona morta de um joystick
Joysticks.Devices(0).Deadzone = 3000
```

> Agora um exemplo prático de como configurar para obter o ID do primeiro controle conectado.
>  Lembre-se de utilizar o `Initialize` antes de qualquer coisa.

```vb
Dim joystickID As Long

joystickID = Joysticks.Devices(0).ID ' Por padrão vem desconectado caso você não tenha plugado ele
```

## Enumerações

### JoystickAxis

Define os eixos disponíveis no controle:
- `LEFT_ANALOG_X` (0): Eixo X do analógico esquerdo
- `LEFT_ANALOG_Y` (1): Eixo Y do analógico esquerdo
- `RIGHT_ANALOG_X` (2): Eixo X do analógico direito
- `RIGHT_ANALOG_Y` (3): Eixo Y do analógico direito
- `TRIGGER_LT` (4): Gatilho esquerdo
- `TRIGGER_RT` (5): Gatilho direito

### JoystickButton

Define os botões disponíveis:
- `BUTTON_X` (0): Botão X
- `BUTTON_A` (1): Botão A
- `BUTTON_B` (2): Botão B
- `BUTTON_Y` (3): Botão Y
- `BUTTON_LB` (4): Bumper esquerdo
- `BUTTON_RB` (5): Bumper direito
- `BUTTON_LT` (6): Gatilho esquerdo
- `BUTTON_RT` (7): Gatilho direito
- `BUTTON_SELECT` (8): Botão Select
- `BUTTON_START` (9): Botão Start
- `BUTTON_LS` (10): Botão do analógico esquerdo
- `BUTTON_RS` (11): Botão do analógico direito

### DPadDirection e JoystickAxisDirection

Definem as direções possíveis para o D-Pad e analógicos:
- `CENTERED/AXIS_DIRECTION_CENTERED` (-1): Centralizado
- `DIRECTION_UP/AXIS_DIRECTION_UP` (0): Para cima
- `DIRECTION_RIGHT/AXIS_DIRECTION_RIGHT` (9000): Para direita
- `DIRECTION_DOWN/AXIS_DIRECTION_DOWN` (18000): Para baixo
- `DIRECTION_LEFT/AXIS_DIRECTION_LEFT` (27000): Para esquerda
- `DIRECTION_UP_RIGHT/AXIS_DIRECTION_UP_RIGHT` (4500): Diagonal superior direita
- `DIRECTION_UP_LEFT/AXIS_DIRECTION_UP_LEFT` (31500): Diagonal superior esquerda
- `DIRECTION_DOWN_RIGHT/AXIS_DIRECTION_DOWN_RIGHT` (13500): Diagonal inferior direita
- `DIRECTION_DOWN_LEFT/AXIS_DIRECTION_DOWN_LEFT` (22500): Diagonal inferior esquerda

## Métodos Principais

### Atualização de Estado

```vba
Public Sub UpdateInput()
```
Atualiza o estado de todos os controles conectados. Deve ser chamada em cada frame ou ciclo de atualização.

### Verificação de Conexão

```vba
Public Function IsConnected(ByVal joyID As Long) As Boolean
```
Verifica se um controle específico está conectado.
- **Parâmetro**: `joyID` - ID do controle
- **Retorno**: `True` se conectado, `False` caso contrário

### Estados dos Botões

```vba
Public Function IsButtonDown(ByVal joyID As Long, ByVal button As JoystickButton) As Boolean
```
Verifica se um botão está pressionado.
- **Parâmetros**: 
  - `joyID` - ID do controle
  - `button` - Botão a ser verificado
- **Retorno**: `True` se pressionado, `False` caso contrário

```vba
Public Function IsButtonPressed(ByVal joyID As Long, ByVal button As JoystickButton) As Boolean
```
Verifica se um botão acabou de ser pressionado (apenas no frame atual).
- **Retorno**: `True` se foi pressionado neste frame, `False` caso contrário

```vba
Public Function IsButtonReleased(ByVal joyID As Long, ByVal button As JoystickButton) As Boolean
```
Verifica se um botão acabou de ser solto.
- **Retorno**: `True` se foi solto neste frame, `False` caso contrário

### Leitura de Eixos Analógicos

```vba
Public Function GetAxisValue(ByVal joyID As Long, ByVal axis As JoystickAxis, Optional ByVal normalized As Boolean = False) As Double
```
Obtém o valor de um eixo analógico.
- **Parâmetros**:
  - `joyID` - ID do controle
  - `axis` - Eixo a ser lido
  - `normalized` - Se `True`, retorna valores entre -1 e 1 (ou 0 e 1 para gatilhos). Se `False`, retorna valores brutos
- **Retorno**: Valor do eixo

```vba
Public Function IsMoving(ByVal joyID As Long, ByVal axis As JoystickAxis) As Boolean
```
Verifica se um eixo analógico está em movimento (fora da zona morta).
- **Retorno**: `True` se em movimento, `False` caso contrário

### D-Pad e Direções

```vba
Public Function IsDPadPressed(ByVal joyID As Long) As Boolean
```
Verifica se houve mudança na direção do D-Pad.
- **Retorno**: `True` se houve mudança, `False` caso contrário

```vba
Public Function GetDPadDirection(ByVal joyID As Long) As DPadDirection
```
Obtém a direção atual do D-Pad.
- **Retorno**: Enum `DPadDirection` indicando a direção

```vba
Public Function GetAnalogStickDirection(ByVal joyID As Long, ByVal axisX As JoystickAxis, ByVal axisY As JoystickAxis) As JoystickAxisDirection
```
Obtém a direção de um analógico baseado em seus eixos X e Y.
- **Parâmetros**:
  - `joyID` - ID do controle
  - `axisX` - Eixo X do analógico
  - `axisY` - Eixo Y do analógico
- **Retorno**: Enum `JoystickAxisDirection` indicando a direção
