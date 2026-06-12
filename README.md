# BateRooming

Aplicativo desktop para comparar, validar e organizar listas de hospedagem em planilhas Excel.

O BateRooming automatiza conferências entre listas, identificando correspondências, divergências, registros repetidos e nomes não encontrados. Todo o processamento acontece localmente, sem envio de dados para serviços externos.

## Funcionalidades

### Bate-Rooming

Compara uma planilha do sistema com uma lista fornecida pelo hotel e gera um relatório estruturado.

- Detecta colunas automaticamente em diferentes layouts.
- Aceita planilhas com ou sem cabeçalho.
- Identifica divergências entre nomes e quartos.
- Separa registros repetidos e sem correspondência.
- Permite ignorar divergências de quarto quando necessário.
- Exporta resultados completos ou somente registros filtrados.
- Gera abas específicas para facilitar a revisão operacional.

### Match de Nomes

Compara duas listas e identifica as correspondências mais prováveis entre nomes.

- Normaliza nomes antes da comparação.
- Usa similaridade textual para encontrar correspondências.
- Permite configurar o percentual mínimo de similaridade.
- Classifica registros encontrados, não encontrados e vazios.
- Preserva a estrutura e a formatação da planilha modelo na exportação.

## Tecnologias

- Python 3.11
- pywebview
- openpyxl
- RapidFuzz
- HTML, CSS e JavaScript
- PyInstaller
- pytest e Ruff

## Estrutura do Projeto

```text
.
├── app.py                 # Aplicação desktop e integração com a interface
├── core/                  # Regras de negócio e exportação de planilhas
├── ui/                    # Interfaces HTML, CSS e JavaScript
├── assets/                # Ícones e imagens genéricas
├── tests/                 # Testes automatizados
├── app.spec               # Configuração de build do PyInstaller
├── requirements.txt       # Dependências de execução
└── requirements-dev.txt   # Dependências de desenvolvimento
```

## Requisitos

- Windows 10 ou superior
- Python 3.11 recomendado
- Microsoft Edge WebView2 Runtime

## Instalação

```bash
git clone https://github.com/urukrehn/baterooming.git
cd baterooming
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Executando o Projeto

```bash
python app.py
```

A aplicação abrirá uma janela desktop com acesso ao Bate-Rooming e ao Match de Nomes.

## Gerando o Executável

```bash
pip install -r requirements-dev.txt
pyinstaller --noconfirm --clean app.spec
```

O executável será criado em `dist/app/app.exe`.

## Testes e Qualidade

```bash
pytest -q
ruff check . --exclude build --exclude dist
```

A suíte cobre regras de matching, comparação de rooming, exportação, API da interface e estrutura do pacote executável.

## Privacidade

- As planilhas são processadas localmente.
- O aplicativo não utiliza APIs externas para executar as comparações.
- Nenhum dado é enviado automaticamente para servidores externos.
- Planilhas reais e arquivos exportados não devem ser adicionados ao repositório.

## Personalização

O repositório utiliza logos genéricos e placeholders de produto/empresa. Antes de distribuir uma versão personalizada, revise:

- textos exibidos na interface;
- metadados em `app_version_info.txt`;
- ícones e imagens em `assets/`;
- título da aplicação em `app.py`.

## Contribuição

1. Crie uma branch para a alteração.
2. Mantenha mudanças pequenas e focadas.
3. Adicione ou atualize testes ao alterar regras de negócio.
4. Execute os testes e o lint antes de abrir um pull request.

## Licença

Este projeto ainda não possui uma licença definida. Adicione um arquivo `LICENSE` antes de reutilizar ou distribuir o código publicamente.
