# Planejador de Experimentos Discretos

## Visão geral

O **Planejador de Experimentos Discretos** é uma aplicação desktop em **Python + Tkinter** para construir planos experimentais em espaços combinatórios discretos, quando o fatorial completo é grande demais para o orçamento disponível.
O programa recebe um conjunto de hiperparâmetros e seus valores possíveis, calcula o tamanho do espaço completo, e então escolhe um subconjunto de configurações com base em três objetivos simultâneos:

1. **Cobertura t-wise** entre hiperparâmetros, com `t = 2` ou `t = 3`.
2. **Balanceamento marginal** dos valores de cada hiperparâmetro.
3. **Diversidade** entre configurações, medida por **distância de Hamming**.

Se o espaço completo couber no orçamento, a solução é exata, isto é, o programa retorna o **fatorial completo**.
Se não couber, o programa usa uma **heurística gulosa iterativa** para gerar um plano aproximado.

O código implementa:
- definição e normalização de hiperparâmetros,
- geração do espaço completo ou de um pool de candidatos,
- avaliação de cobertura marginal e t-wise,
- seleção gulosa multiobjetivo,
- modo de **screening**,
- modo de **refinement** com hiperparâmetros fixados,
- exportação para **XLSX** com resumo e plano experimental. fileciteturn1file0

---

## Objetivo do programa

O problema tratado pela aplicação é o seguinte.

Dado um conjunto de hiperparâmetros discretos

\[
I_1, I_2, \ldots, I_N
\]

o espaço completo de configurações é o produto cartesiano:

\[
E = I_1 \times I_2 \times \cdots \times I_N
\]

e o número total de experimentos possíveis é:

\[
|E| = \prod_{n=1}^{N} |I_n|
\]

Quando \(|E|\) é muito grande, torna-se inviável executar todas as combinações.
Nesse caso, o objetivo passa a ser selecionar um subconjunto \(S \subset E\), com cardinalidade limitada pelo orçamento, de forma que ele:
- cubra bem o espaço,
- represente adequadamente os valores de cada hiperparâmetro,
- e evite redundância excessiva entre configurações.

Esse tipo de problema pertence ao campo de **Design of Experiments (DoE)**, **combinatorial testing**, **covering arrays** e **space-filling design** em espaços discretos.

---

## Fundamentação teórica

### 1. Espaço combinatório discreto

Se cada hiperparâmetro assume um número finito de valores, o conjunto de experimentos possíveis é discreto e finito.
A enumeração completa é ideal apenas quando o custo computacional ou experimental permite.

O programa calcula o tamanho do espaço completo por:

\[
|E| = \prod_{n=1}^{N} |I_n|
\]

No código, isso é feito pela função `full_factorial_count`, que multiplica o número de valores de cada hiperparâmetro. fileciteturn1file0

### 2. Orçamento experimental

O usuário define:
- o número máximo total de operações (`max_operations`),
- e o número de repetições por configuração (`repetitions`).

Então o número máximo de configurações únicas é:

\[
U = \left\lfloor \frac{\text{max\_operations}}{\text{repetitions}} \right\rfloor
\]

Se \(U \ge |E|\), a aplicação executa o fatorial completo.
Caso contrário, ela precisa aproximar uma solução ótima.

### 3. Cobertura marginal

A cobertura marginal procura equilibrar a frequência de cada valor individual de cada hiperparâmetro.
Se um hiperparâmetro \(H_n\) possui \(|I_n|\) valores e o plano contém \(U\) configurações únicas, o alvo ideal por valor é:

\[
T_{n,v} = \frac{U}{|I_n|}
\]

No código, isso é calculado pela função `compute_targets`.
Durante a seleção, o programa premia candidatos que ajudam a aproximar esses alvos.

A ideia é evitar viés, por exemplo:
- testar excessivamente um valor,
- e quase ignorar os demais.

### 4. Cobertura t-wise

A cobertura **t-wise** procura garantir que combinações entre grupos de \(t\) hiperparâmetros sejam observadas no plano.

#### Cobertura 2-wise
Busca cobrir pares de hiperparâmetros.

#### Cobertura 3-wise
Busca cobrir trincas de hiperparâmetros.

A motivação teórica vem da literatura de **combinatorial interaction testing**: muitos defeitos, comportamentos emergentes ou interações relevantes aparecem em combinações de baixa ordem, principalmente pares e trincas.

No código:
- `build_twise_universe` constrói o universo de combinações possíveis,
- `covered_twise_of_config` calcula quais combinações t-wise uma configuração cobre,
- a heurística premia candidatos que cobrem combinações ainda não vistas. fileciteturn1file0

### 5. Diversidade por distância de Hamming

A aplicação mede diversidade entre duas configurações \(x\) e \(y\) por meio da **distância de Hamming**:

\[
d(x,y) = \sum_{i=1}^{N} \mathbf{1}[x_i \neq y_i]
\]

Quanto maior a distância, mais diferentes são as configurações.
Na prática, isso reduz redundância e melhora o espalhamento do plano no espaço discreto.

O programa usa:
- distância mínima entre uma nova configuração e as já selecionadas,
- distância mínima e média do conjunto final como indicadores de qualidade.

### 6. Heurística gulosa multiobjetivo

Quando o espaço não cabe no orçamento, a seleção é feita por uma heurística gulosa com múltiplos critérios.

Para cada candidato, o programa calcula uma pontuação composta por:
- ganho em cobertura t-wise,
- ganho em balanceamento marginal,
- ganho em diversidade.

Em termos conceituais, a função de score é:

\[
\text{score}(x) =
w_t \cdot G_t(x) +
w_m \cdot G_m(x) +
w_d \cdot G_d(x)
\]

onde:
- \(G_t(x)\) é o ganho de cobertura t-wise,
- \(G_m(x)\) é o ganho marginal,
- \(G_d(x)\) é o ganho de diversidade,
- \(w_t, w_m, w_d\) são os pesos definidos pelo usuário.

No código, isso é implementado em `score_candidate`. fileciteturn1file0

### 7. Screening e refinement

O programa implementa duas fases conceitualmente distintas:

#### Screening
Fase ampla, exploratória.
O objetivo é amostrar bem o espaço global e identificar regiões promissoras.

#### Refinement
Fase focada.
Alguns hiperparâmetros são fixados em valores escolhidos, reduzindo o espaço de busca e permitindo gerar um plano mais denso e específico naquela sub-região.

Esse fluxo é coerente com práticas comuns em:
- otimização de hiperparâmetros,
- DOE sequencial,
- exploração seguida de aprofundamento local.

---

## Arquitetura do programa

O sistema é composto por:

### Núcleo algorítmico
Funções para:
- normalização de valores,
- cálculo do fatorial completo,
- geração de candidatos,
- cálculo de métricas,
- planejamento heurístico,
- exportação.

### Interface gráfica
Construída com **Tkinter** e `ttk`, organizada em abas:
- **Hiperparâmetros**
- **Planejamento**
- **Fases**
- **Resultado**

### Execução assíncrona
A heurística roda em **thread separada**, com:
- log incremental,
- barra de progresso,
- cancelamento,
- estimativa de tempo.

### Exportação
O resultado pode ser exportado para **XLSX** com:
- planilha de experimentos,
- aba de resumo do planejamento. fileciteturn1file0

---

## Requisitos

- Python 3.10 ou superior, recomendado
- Tkinter
- openpyxl

### Instalação

