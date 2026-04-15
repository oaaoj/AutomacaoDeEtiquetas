# Automação de Geração de Etiquetas sem EAN

## Descrição
Este projeto tem como objetivo automatizar o processo de geração de arquivos de etiquetas para produtos sem código EAN, eliminando a necessidade de operações manuais realizadas diretamente no sistema.

Anteriormente, o processo era executado por meio do cadastro manual, loja a loja, para geração dos arquivos no formato necessário — uma atividade repetitiva, suscetível a erros e com alto custo operacional. A automação reduz esse esforço de horas para segundos, garantindo padronização e eficiência.

## Funcionamento
A aplicação realiza um pipeline automatizado de processamento de dados:

1. Leitura de uma planilha base localizada em diretório padrão (ambiente corporativo)
2. Ingestão de tabela auxiliar contendo referências de produtos (ex: códigos, descrições, etc.)
3. Tratamento e transformação dos dados
4. Distribuição lógica dos itens por loja
5. Geração de múltiplos arquivos Excel estruturados, prontos para uso operacional

## Arquitetura / Uso
- Execução via aplicação empacotada (.exe), dispensando conhecimento técnico por parte do usuário final
- Input e output baseados em diretórios pré-configurados (padrão corporativo)
- Fluxo orientado a processamento em lote (batch)

## Benefícios
- Redução significativa de tempo operacional (horas → segundos)
- Eliminação de etapas manuais no sistema
- Padronização da saída de dados
- Minimização de erros humanos
- Facilidade de uso para times não técnicos

## Observações
O projeto está versionado e disponível neste repositório, permitindo evolução contínua conforme novas regras de negócio e necessidades operacionais.
