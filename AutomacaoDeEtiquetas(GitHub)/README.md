# Publicacao GitHub - Automacao de Etiquetas

Esta é uma aplicação desenvolvida para Otimizar o tempo da equipe dos Analistas de Compras.

Ela faz o trabalho de, automaticamente, ler o número de pares em cada uma das Lojas (que estão dispostos em uma planilha Excel) e
escrever o número de linhas que equivale ao números de sapatos que irão ser enviados para um determinado local.

Ela foi desenvolvida para substituir o trabalho manual de cadastro de itens que não possuem EAN (European Article Number), o que reduziu
um trabalho de algumas horas em segundos.

## Como configurar

1. Defina a variavel de ambiente `AUTOMACAO_ETIQUETAS_BASE_DIR` com a pasta base da automacao.
2. Garanta que dentro da pasta base existam:
- `ARQUIVOS_BASE`
- `ARQUIVOS_ETIQUETA`

Exemplo de caminho base (interno):
- `\\servidor\pasta\AutomacaoEtiquetas`

## Arquivos principais

- `src/AutomacaoDeEtiquetas(Beta).py` (versao com interface)
- `src/automacaoetiquetas.py` (versao script)

## Build local (opcional)

```bash
pyinstaller src/AutomacaoDeEtiquetas(Beta).spec
```

