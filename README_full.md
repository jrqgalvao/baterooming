# {{PRODUCT_NAME}} v1.4.4

Aplicativo desktop local para conferência operacional de planilhas, com dois fluxos principais:

- **Bate-Rooming:** compara listas de rooming e gera um relatório operacional.
- **Match de Nomes:** padroniza nomes entre duas planilhas preservando o modelo de exportação.

Esta pasta é uma cópia pública sanitizada. Substitua `{{PRODUCT_NAME}}` e `{{COMPANY_NAME}}` antes de distribuir uma versão identificada.

## Recursos

- Interface desktop local com `pywebview`.
- Leitura e exportação de arquivos Excel.
- Normalização e matching de nomes com `rapidfuzz`.
- Detecção de duplicatas e placeholders.
- Processamento 1:1 sem reutilizar o mesmo registro de referência.
- Exportação formatada para revisão operacional.
- Assets genéricos para uso em repositório público.

## Bate-Rooming

O fluxo compara a Planilha 1, usada como referência, com a Planilha 2, usada para conferência.

O arquivo exportado contém as abas:

- `RESUMO`
- `RESULTADO COMPLETO`
- `DIVERGÊNCIAS`
- `SEM CORRESPONDÊNCIA`
- `REPETIDOS`
- `LOG`

A aba `LOG` registra, para cada linha, a decisão final, etapa do matching, score, threshold, candidatos considerados, candidatos bloqueados, conflitos 1:1 e motivo da decisão.

### Regras de matching

- Nomes são normalizados antes da comparação, removendo acentos, caixa, pontuação e espaços duplicados.
- Matches exatos após a normalização têm prioridade.
- Nomes repetidos e placeholders são excluídos do fuzzy matching e classificados como `REPETIDO`.
- Entre candidatos com o mesmo primeiro e último token significativo, vence o maior score.
- O mesmo registro de referência pode ser usado apenas uma vez.
- Fuzzy entre identidades diferentes exige score conservador e primeiro/último token compatíveis.
- Conflitos de referência são mantidos no resultado e detalhados no `LOG`.

## Match de Nomes

- Detecta automaticamente a coluna de nomes quando possível.
- Mantém a ordem e as linhas vazias da planilha de destino.
- Permite ajustar o percentual mínimo de similaridade.
- Exporta usando a planilha modelo como base.

## Estrutura

```text
.
|-- app.py
|-- app.spec
|-- app_version_info.txt
|-- app.exe.config
|-- assets/
|-- core/
|-- tests/
|-- ui/
|-- requirements.txt
|-- requirements-dev.txt
`-- README_full.md
```

## Requisitos

- Windows 10 ou superior.
- Python 3.11 recomendado.
- WebView2 Runtime.

Dependências de runtime:

```bash
pip install -r requirements.txt
```

Dependências de desenvolvimento:

```bash
pip install -r requirements-dev.txt
```

## Execução local

```bash
python app.py
```

## Testes e qualidade

```bash
pytest -q
ruff check . --exclude build --exclude dist
```

Também é recomendado validar a sintaxe JavaScript das três telas e executar um fluxo controlado de cada ferramenta.

## Build do executável

```bash
pyinstaller --noconfirm --clean app.spec
```

O resultado será criado em `dist/app/app.exe`.

## Publicação

Antes de publicar:

1. Substitua os placeholders de produto e organização.
2. Revise os assets genéricos.
3. Confirme que não existem planilhas reais ou dados sensíveis.
4. Não inclua `dist/`, `build/`, caches ou arquivos temporários.
5. Rode os testes e valide o executável final.

O processamento principal é local e não depende de API externa.
