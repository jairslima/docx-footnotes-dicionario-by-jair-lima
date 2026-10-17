# Footnote Injector by Jair Lima

Este utilitário automatiza a injeção de notas de rodapé explicativas em documentos técnicos e livros tecnológicos no formato Microsoft Word (.docx). Ele varre o arquivo DOCX, identifica termos técnicos pré-cadastrados em um dicionário (JSON ou CSV) e insere as notas de rodapé nativas preservando 100% da formatação do documento.

---

## Recursos Principais

* **Injeção Nativa de Notas de Rodapé**: As notas adicionadas são tratadas pelo Word como notas de rodapé nativas, com numeração e posicionamento automático ao pé da página.
* **Preservação Absoluta de Formatação**: Utiliza clonagem profunda de elementos XML (`w:r`) do OpenXML, permitindo inserir a marcação sem alterar estilos de fonte, cores, negritos ou itálicos originais.
* **Auto-infraestrutura de Notas**: Se o documento DOCX original não contiver notas de rodapé, o utilitário cria automaticamente toda a estrutura XML necessária (arquivos de notas, relacionamentos e tipos de conteúdo).
* **Controle de Ocorrências**: Permite definir se as notas devem ser geradas apenas na primeira ocorrência de cada termo no texto (comportamento padrão recomendado para livros e relatórios técnicos) ou em todas as ocorrências.
* **Busca Segura por Regex**: Utiliza limites de palavras (`\btermo\b`) em modo case-insensitive para evitar que termos semelhantes disparem em correspondências parciais indesejadas (exemplo: "porta" dentro de "importar").
* **Suporte a JSON e CSV**: Aceita bases de dicionário estruturadas de forma simples.

---

## Estrutura do Projeto

O projeto possui os seguintes arquivos principais:

* `docx_footnote_injector.py`: Script do core e executável em linha de comando (CLI).
* `docx_footnote_gui.py`: Interface Gráfica (GUI) moderna desenvolvida em CustomTkinter.
* `build_executables.ps1`: Script PowerShell automatizado para compilar a aplicação com PyInstaller.
* `dicionario_padrao.json`: Dicionário básico de conceitos tecnológicos.
* `test_injector.py`: Testes unitários de validação XML.
* `dist/`: Diretório que armazena os executáveis compilados (.exe) gerados.

---

## Como Utilizar os Executáveis (.exe)

Você pode rodar a aplicação diretamente através dos executáveis prontos contidos no diretório [dist](dist/):

### 1. Interface Gráfica (GUI)
Abra a pasta `dist` e execute o arquivo `docx-footnote-gui-by-jair-lima.exe` com dois cliques. A interface gráfica permite:
* Escolher visualmente o documento Word de entrada (.docx).
* Selecionar o dicionário (JSON ou CSV) ou usar o padrão.
* Definir onde salvar o arquivo de saída.
* Configurar o controle de ocorrências através de caixas de seleção.
* Monitorar logs de progresso em tempo real e receber alertas de conclusão.

### 2. Executável CLI
Para executar via prompt de comando ou PowerShell sem necessitar do interpretador Python local:
```bash
./dist/docx-footnote-injector-by-jair-lima.exe caminho/do/documento_entrada.docx caminho/do/documento_saida.docx
```

---

## Como Utilizar os Scripts em Python

### Requisitos e Instalação
Caso queira executar os scripts originais diretamente no Python (versão 3.7 ou superior), é necessário instalar as dependências requeridas:
```bash
pip install python-docx lxml customtkinter pyinstaller
```

### 1. Executando a Interface Gráfica
```bash
python docx_footnote_gui.py
```

### 2. Executando via Script CLI
Para processar utilizando o dicionário de tecnologia padrão:
```bash
python docx_footnote_injector.py caminho/do/documento_entrada.docx caminho/do/documento_saida.docx
```

Para usar um dicionário JSON ou CSV customizado:
```bash
python docx_footnote_injector.py entrada.docx saida.docx --dict meu_dicionario.json
```

Para forçar a inserção de nota em todas as ocorrências de um termo no texto (em vez de apenas na primeira):
```bash
python docx_footnote_injector.py entrada.docx saida.docx --all
```

---

## Compilando os Executáveis Novamente
Se você fizer alterações no código fonte e desejar rebuildar os executáveis, execute o script auxiliar de compilação em um terminal PowerShell:
```powershell
./build_executables.ps1
```
Os novos binários serão re-gerados e substituídos na pasta `dist/`.

---

## Validação e Testes
Para executar os testes unitários e validar as rotinas de injeção XML do injetor:
```bash
python test_injector.py
```

