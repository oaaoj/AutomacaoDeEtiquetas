# Publicacao GitHub - Automacao de Etiquetas

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

