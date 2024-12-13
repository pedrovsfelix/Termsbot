# TermsBot

**TermsBot** é uma aplicação em Python desenvolvida para gerar documentos de termos de forma automática e massiva. Ele utiliza uma lista em Excel como fonte de dados e um modelo de documento Word (DOCX) como base para criar documentos personalizados. O projeto surgiu a partir de um problema real onde o processo manual exigia boas horas de trabalho e muitas vezes erros na formatação dos documentos.

## Funcionalidades

- Importar uma lista de dados de um arquivo Excel.
- Importar um modelo de documento Word.
- Substituir marcadores no modelo com dados da lista.
- Gerar documentos personalizados para cada entrada na lista.

## Tecnologias Utilizadas

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter): Para a interface gráfica.
- [Tkinter](https://docs.python.org/3/library/tkinter.html): Para caixas de diálogo e mensagens.
- [Pandas](https://pandas.pydata.org/): Para manipulação de arquivos Excel.
- [python-docx](https://python-docx.readthedocs.io/en/latest/): Para manipulação de documentos Word.

## Requisitos

- Python 3.8 ou superior.
- Bibliotecas adicionais (instale usando o comando abaixo):

```bash
pip install customtkinter pandas python-docx
```

## Instalação

1. Clone este repositório ou baixe os arquivos.
2. Certifique-se de que as bibliotecas necessárias estão instaladas (veja a seção **Requisitos**).
3. Execute o arquivo principal `app.py`.

```bash
python app.py
```

## Como Usar

1. **Buscar Lista**:
   - Clique no botão "Buscar Lista" e selecione um arquivo Excel que contenha os dados.
   - O arquivo deve ter as seguintes colunas:
     - `NOME`
     - `RG`
     - `CPF`
     - `CARGO`
     - `EMPRESA`

2. **Buscar Termo**:
   - Clique no botão "Buscar Termo" e selecione o modelo de documento Word (DOCX).
   - Use placeholders no modelo com os nomes das colunas (ex.: `NOME`, `RG`, etc.).

3. **Gerar Termos**:
   - Clique no botão "Gerar Termos". Os documentos gerados serão salvos na pasta `termos` no mesmo diretório do programa.

## Estrutura do Projeto

```
TermsBot/
├── app.py             # Arquivo principal do programa
├── termos/            # Diretório onde os documentos gerados serão salvos 
└── README.md          # Documentação do projeto
```

## Possíveis Erros

- **Colunas ausentes no arquivo Excel**:
  - Certifique-se de que o arquivo contém todas as colunas necessárias.

- **Erro ao acessar arquivos**:
  - Verifique os caminhos dos arquivos selecionados.
  - Confirme que os arquivos não estão sendo usados por outro programa.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir uma _issue_ ou enviar um _pull request_.

## 🤝 Desenvolvedor

<table>
  <tr>
    <td align="center">
      <a href="https://github.com/pedrovsfelix" title="Aluno">
        <img src="https://avatars.githubusercontent.com/u/93714667?v=4" width="100px;" alt="Foto do Aluno Pedro Félix"/><br>
        <sub>
          <b>Pedro Félix</b>
        </sub>
      </a>
    </td>
  </tr>
</table>
