# Math

## Visão Geral

Esse módulo adiciona funcionalidades essenciais de cálculos matemáticos à linguagem VBA (Visual Basic for Applications), focado em áreas como desenvolvimento de jogos e simulações físicas. Ele oferece uma ampla gama de funções que lidam com vetores, cálculos de distâncias, ângulos, rotações 2D, colisões, forças gravitacionais, e progressões matemáticas.

Este módulo foi projetado para ser eficiente e rápido, proporcionando cálculos otimizados que atendem a necessidades avançadas de manipulação geométrica e simulações numéricas.

### Funcionalidades

A seguir, as principais funções e seus objetivos:

## Funções Básicas

- **PI**  
  A constante pi (`π`), usada para cálculos envolvendo ângulos, rotações e trigonometria.

## Cálculos de Vetores e Geometria

- **DotProduct2D(v As Variant, axis As Variant) As Double**  
  Calcula o produto escalar (dot product) entre dois vetores 2D.

- **NormalizeVector2D(v As Variant) As Variant**  
  Normaliza um vetor 2D, tornando-o unitário.

- **AngleBetweenVectors2D(vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double) As Double**  
  Calcula o ângulo entre dois vetores 2D em graus.

- **RotatePoint2D(x As Double, y As Double, angle As Double) As Variant**  
  Rotaciona um ponto 2D em torno da origem por um ângulo especificado.

- **Distance2D(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double**  
  Calcula a distância Euclidiana entre dois pontos 2D.

## Funções de Física

- **CollisionForce(mass1 As Double, mass2 As Double, velocity1 As Double, velocity2 As Double) As Double**  
  Calcula a força de colisão entre dois objetos com base em suas massas e velocidades.

- **ElasticCollisionVelocity(mass1 As Double, mass2 As Double, velocity1 As Double, velocity2 As Double) As Variant**  
  Calcula as velocidades de dois objetos após uma colisão elástica.

- **GravitationalForce(mass1 As Double, mass2 As Double, distance As Double) As Double**  
  Calcula a força gravitacional entre dois corpos com base em suas massas e distância.

- **Friction(force As Double, coefficient As Double) As Double**  
  Calcula a força de atrito com base na força aplicada e no coeficiente de atrito.

## Matrizes e Transformações

- **RotationMatrix2D(angle As Double) As Variant**  
  Gera uma matriz de rotação 2D para um ângulo dado.

- **MultiplyMatrix2D(m As Variant, v As Variant) As Variant**  
  Multiplica uma matriz 2D por um vetor 2D.

- **Determinant2x2(m As Variant) As Double**  
  Calcula o determinante de uma matriz 2x2.

## Cálculos de Distância

- **EuclideanDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double**  
  Calcula a distância Euclidiana entre dois pontos.

- **ManhattanDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double**  
  Calcula a distância Manhattan (distância absoluta) entre dois pontos.

- **ChebyshevDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double**  
  Calcula a distância Chebyshev (distância máxima entre as coordenadas de dois pontos).

## Funções Trigonométricas e Exponenciais

- **ArcSine(value As Double) As Double**  
  Calcula o arco seno de um valor (inverso de `sin`).

- **ArcCosine(value As Double) As Double**  
  Calcula o arco cosseno de um valor (inverso de `cos`).

- **ArcTangent(value As Double) As Double**  
  Calcula o arco tangente de um valor (inverso de `tan`).

- **Sinh(x As Double) As Double**  
  Calcula o seno hiperbólico de `x`.

- **Cosh(x As Double) As Double**  
  Calcula o cosseno hiperbólico de `x`.

- **Tanh(x As Double) As Double**  
  Calcula a tangente hiperbólica de `x`.

## Progressões Matemáticas

- **ArithmeticProgression(a1 As Double, d As Double, n As Long) As Double**  
  Calcula o n-ésimo termo de uma progressão aritmética.

- **SumArithmeticProgression(a1 As Double, d As Double, n As Long) As Double**  
  Calcula a soma dos n primeiros termos de uma progressão aritmética.

- **GeometricProgression(a1 As Double, r As Double, n As Long) As Double**  
  Calcula o n-ésimo termo de uma progressão geométrica.

- **SumGeometricProgression(a1 As Double, r As Double, n As Long) As Double**  
  Calcula a soma dos n primeiros termos de uma progressão geométrica.

## Funções Auxiliares

- **Clamp(value As Double, Min As Double, Max As Double) As Double**  
  Limita um valor a um intervalo específico (mínimo e máximo).

- **RandomFloat(Min As Double, Max As Double) As Double**  
  Gera um número aleatório de ponto flutuante dentro de um intervalo.

- **Map(value As Double, inMin As Double, inMax As Double, outMin As Double, outMax As Double) As Double**  
  Mapeia um valor de um intervalo para outro intervalo.

- **Mean(values As Variant) As Double**  
  Calcula a média de um conjunto de valores.

- **StandardDeviation(values As Variant) As Double**  
  Calcula o desvio padrão de um conjunto de valores.

## Utilização

Este módulo pode ser integrado diretamente ao VBA e utilizado em qualquer tipo de aplicação que necessite de cálculos matemáticos avançados, como:

- Simulações físicas (ex.: motores de física em jogos).
- Cálculos de movimento, colisões e trajetórias.
- Manipulações geométricas e transformações de vetores/matrizes.
  E ETC MANO.

---

## Referências

- [Khan Academy - Álgebra Linear & Vetores](https://pt.khanacademy.org/math/linear-algebra)
- [MathWorld - Cálculos Vetoriais](https://mathworld.wolfram.com/Vector.html)
- [OpenGL Mathematics (GLM) Library para C++](https://github.com/g-truc/glm)
- [Rosetta Code - Exemplos de Geometria](https://rosettacode.org/wiki/Category:Collision_detection)
- [Game Physics Engine Development - Ian Millington](https://github.com/matheusportela/Poiesis/blob/master/references/Game%20Physics%20Engine%20Development%20-%20Ian%20Millington.pdf)
