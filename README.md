# Tratamento de Catálogo de Produtos - OSGT

## Sobre o Projeto

Este projeto contém um script desenvolvido para automatizar o processo de tratamento e validação das respostas fornecidas por analistas de importadores no catálogo de produtos da plataforma OSGT.

O objetivo principal é garantir a integridade, a padronização e a qualidade dos dados inseridos no sistema, reduzindo a carga de trabalho manual e minimizando erros.

### 🎯 O Problema

O preenchimento manual do catálogo de produtos por diferentes analistas pode gerar uma série de inconsistências, como:

* Erros de digitação.
* Falta de padronização (maiúsculas/minúsculas, espaços extras).
* Valores preenchidos em formatos incorretos (texto em campo numérico, etc.).
* Dados estruturais que precisam ser reorganizados.

Este script atua como uma camada de controle de qualidade, aplicando um conjunto de regras de negócio para limpar e validar os dados antes que eles sejam consumidos por outros sistemas.

### ✨ Principais Funcionalidades

O script realiza as seguintes operações:

* **Padronização Geral:** Converte textos para maiúsculas, remove espaços desnecessários e trata valores nulos (`#N/A`, `NAN`).
* **Validação de Tipos:** Verifica se os valores em colunas específicas correspondem ao tipo de dado esperado (ex: `NÚMERO_INTEIRO`, `NÚMERO_REAL`, `BOOLEANO`).
* **Correção Estrutural:** Reorganiza linhas onde a especificação de um item é informada em uma linha separada (casos "99" e "999").
* **Limpeza de Unidades de Medida:** Identifica e remove unidades de medida de campos numéricos (ex: "150 KG/L" → "150"), validando a unidade contra uma lista pré-definida.
* **Geração de Relatório de Erros:** Cria um arquivo Excel de saída onde as células com dados inválidos (que não puderam ser corrigidos automaticamente) são destacadas em vermelho para facilitar a revisão manual.

### 🛠️ Tecnologias Utilizadas

* **Python**
* **Pandas:** Para manipulação e análise dos dados em formato de DataFrame.
* **NumPy:** Para operações numéricas e tratamento de valores nulos.
* **OpenPyXL:** Como motor para ler e escrever arquivos `.xlsx`.
