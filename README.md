
# Gerador de Certificados em Word e PDF

Este projeto automatiza a geração de certificados personalizados em **Word** e **PDF**. Ele utiliza um modelo Word como base, substitui o texto do modelo por nomes de participantes fornecidos em um arquivo **Excel** e organiza os arquivos gerados em pastas específicas.

## Funcionalidades

- Substitui a string `NNnomeNN` no modelo Word pelo nome do participante.
- Gera arquivos **Word** para todos os nomes fornecidos.
- Converte os arquivos **Word** para **PDF**.
- Organiza os PDFs em subpastas separadas, nomeadas de acordo com as planilhas do arquivo Excel.

---

## Estrutura do Projeto

```
/projeto_certificados
│
├── mod_certificado_word.docx       # Modelo do certificado em Word (template)
├── xls_para_certificados.xlsx      # Arquivo Excel com os nomes organizados em planilhas
├── main.py                         # Script principal para gerar os certificados
├── Certificados_word/              # Pasta com os arquivos Word gerados
└── Certificados_pdf/               # Pasta principal com subpastas para PDFs organizados
    ├── Planilha1/                  # Subpasta com PDFs da primeira planilha
    ├── Planilha2/                  # Subpasta com PDFs da segunda planilha
    └── ...
```

---

## Pré-requisitos

Antes de executar o projeto, certifique-se de que as seguintes bibliotecas Python estejam instaladas:

```bash
pip install python-docx pandas docx2pdf
```

### Requisitos do Sistema

- **Python** 3.8 ou superior
- **Sistema operacional Windows** (a biblioteca `docx2pdf` só funciona no Windows).

---

## Como Usar

### 1. Prepare os arquivos de entrada:

- **`mod_certificado_word.docx`**: Modelo do certificado com a string `NNnomeNN` que será substituída.
- **`xls_para_certificados.xlsx`**: Arquivo Excel com os nomes organizados em colunas chamadas **NOME** e planilhas separadas.

### 2. Execute o Script

No terminal ou prompt de comando, execute o script:

```bash
python main.py
```

### 3. Resultados

- Os arquivos **Word** serão salvos na pasta `Certificados_word`.
- Os arquivos **PDF** serão organizados em subpastas dentro de `Certificados_pdf`, conforme os nomes das planilhas do Excel.

---

## Exemplo de Estrutura do Excel

O arquivo Excel deve conter planilhas com a seguinte estrutura:

| NOME           |
|----------------|
| João da Silva  |
| Maria Santos   |
| Pedro Oliveira |

---

## Exemplo de Execução

Ao executar o script, o terminal exibirá mensagens como:

```
Certificados Word gerados em: Certificados_word
Certificados PDF organizados em: Certificados_pdf
```

---

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir um **Pull Request** ou relatar problemas na aba **Issues**.

---

## Licença

Este projeto é licenciado sob a **MIT License**. Consulte o arquivo `LICENSE` para mais detalhes.

---

## Autor

Desenvolvido por **Joao Paulo Tot - emailparaojp@gmail.com**.
Se esse sistema te ajudou, sinta-se a vontade para fazer uma **contribuição via PIX - pixdojp@gmail.com**
