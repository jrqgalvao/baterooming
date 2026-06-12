# {{PRODUCT_NAME}} v1.4.4

Ferramenta desktop em Python para apoiar rotinas de conferência operacional com planilhas. O projeto inclui duas funções principais:

- **Bate-Rooming:** compara bases de rooming/listas e gera planilha final com resultado completo, divergências, sem correspondência e repetidos.
- **Match de Nomes:** identifica correspondências prováveis entre listas de nomes, preservando a estrutura da planilha modelo para exportação.

Este repositório está preparado como versão pública. Substitua os placeholders abaixo antes de distribuir uma versão final:

- `{{PRODUCT_NAME}}`: nome público do aplicativo.
- `{{COMPANY_NAME}}`: nome da organização, autor ou marca.

## Principais Recursos

- Interface desktop com `pywebview`.
- Leitura e exportação de planilhas Excel.
- Matching textual com normalização de nomes.
- Exportação com abas organizadas e formatação aplicada.
- Preservação de layout visual via arquivos HTML/CSS locais.
- Empacotamento Windows com PyInstaller.
- Assets genéricos para uso inicial em repositório público.

## Estrutura do Projeto

```text
.
├── app.py                     # Entrada principal do aplicativo desktop
├── app.spec                   # Configuração do PyInstaller
├── app_version_info.txt       # Metadados do executável Windows
├── app.exe.config             # Configuração auxiliar do executável
├── assets/                    # Ícone e logos genéricos
├── core/                      # Regras de negócio e exportações
├── tests/                     # Testes automatizados
├── ui/                        # Telas HTML/CSS/JS
├── requirements.txt           # Dependências de runtime
├── requirements-dev.txt       # Dependências de desenvolvimento
└── README_full.md             # Documentação detalhada
```

## Requisitos

- Windows 10 ou superior.
- Python 3.11 recomendado.
- WebView2 Runtime instalado no Windows.

Dependências de runtime:

```bash
pip install -r requirements.txt
```

Dependências de desenvolvimento:

```bash
pip install -r requirements-dev.txt
```

## Como Rodar em Desenvolvimento

Na raiz do projeto:

```bash
python app.py
```

O aplicativo abre uma janela desktop com o menu principal e navegação para as ferramentas disponíveis.

## Como Gerar o Executável

Na raiz do projeto:

```bash
pyinstaller --noconfirm --clean app.spec
```

O resultado será criado em:

```text
dist/app/app.exe
```

## Testes e Qualidade

Execute a suíte completa:

```bash
pytest -q
```

Execute lint:

```bash
ruff check . --exclude build --exclude dist
```

Valide a sintaxe JavaScript das telas extraindo os scripts dos HTMLs ou usando o fluxo já coberto pelos testes estruturais.

## Regras de Negócio

### Match de Nomes

- Detecta a coluna de nomes automaticamente quando possível.
- Normaliza nomes antes de comparar.
- Mantém linhas vazias como casos identificáveis.
- Retorna status de encontrado, não encontrado ou nome vazio.
- Preserva as linhas da planilha modelo ao exportar.

### Bate-Rooming

- Compara registros entre bases de origem.
- Gera abas separadas para resultado completo, divergências, sem correspondência e repetidos.
- Mantém mensagens e status orientados ao usuário final.
- Exporta planilha formatada para facilitar revisão operacional.

## Customização Antes de Publicar

Antes de disponibilizar o projeto publicamente, revise:

- Trocar `{{PRODUCT_NAME}}` pelo nome real do produto.
- Trocar `{{COMPANY_NAME}}` pelo nome público desejado.
- Substituir os logos genéricos em `assets/`, se necessário.
- Atualizar `app_version_info.txt` com metadados finais.
- Atualizar o título da janela em `app.py`.
- Rodar busca textual para garantir que nenhum placeholder indevido ficou no pacote final.

Comando sugerido para localizar placeholders:

```bash
rg "\{\{PRODUCT_NAME\}\}|\{\{COMPANY_NAME\}\}" .
```

## Segurança e Privacidade

- O processamento é local.
- O projeto não depende de API externa para executar as regras principais.
- Não inclua planilhas reais, dados sensíveis ou artefatos gerados no repositório.
- Mantenha `dist/`, `build/`, caches e arquivos temporários fora do versionamento.

## Arquivos que Não Devem Ir para o Repositório

A configuração `.gitignore` já cobre os principais casos:

- `dist/`
- `build/`
- caches de Python e testes
- arquivos temporários
- ambientes virtuais

## Checklist de Release

1. Atualizar versão em `app.py`, UI e `app_version_info.txt`.
2. Rodar `pytest -q`.
3. Rodar `ruff check . --exclude build --exclude dist`.
4. Gerar executável com PyInstaller.
5. Abrir o executável e validar a navegação principal.
6. Testar um fluxo real de cada ferramenta com dados de exemplo.
7. Confirmar que o pacote final não contém marcas, dados sensíveis ou arquivos temporários.

## Licença

Defina a licença antes de publicar o repositório. Se o projeto for privado ou interno, deixe isso explícito em um arquivo `LICENSE` ou na descrição do repositório.
