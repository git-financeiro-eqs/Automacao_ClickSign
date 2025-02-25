# Automação de Termo de Recebimento

Este projeto é uma automação para gerar e enviar termos de recebimento de cartões corporativos para os funcionários da empresa. O programa é composto por três módulos principais:

- **geradorFormularios.py**: Gera PDFs personalizados com os termos de recebimento para cada funcionário.
- **clicksign.py**: Envia os PDFs gerados para assinatura via e-mail utilizando a plataforma ClickSign.
- **gui.py**: Interface gráfica para interação com o usuário, permitindo a seleção de arquivos e execução das funcionalidades.
<br/>

## Requisitos

- **Python 3.x**
- **Bibliotecas Python**:
  - `selenium`
  - `webdriver_manager`
  - `pandas`
  - `pyperclip`
  - `pyautogui`
  - `reportlab`
  - `tkinter`
<br/>
<br/>

## Instalação

1. Clone o repositório:
 ```bash
 git clone https://github.com/seu-usuario/automacao-termo-recebimento.git
 cd automacao-termo-recebimento
```
<br/>

Instale as dependências:

```bash
Copy
pip install -r requirements.txt
```

### Uso
Execute o programa:

```bash
Copy
python gui.py
```
<br/>
<br/>

### Na interface gráfica:

Insira a data no formato **DD/MM/AAAA**.

Selecione o arquivo Excel contendo os dados dos funcionários.

Clique em **"Gerar PDFs"** para criar os termos de recebimento.

Clique em **"Enviar E-mail"** para enviar os termos para assinatura via ClickSign.
<br/>
<br/>

### Estrutura do Projeto
**PDFs/**: Pasta onde os PDFs gerados são armazenados.

**Imagens/**: Contém as imagens utilizadas na interface gráfica.

**geradorFormularios.py**: Módulo responsável por gerar os PDFs.

**clicksign.py**: Módulo responsável por enviar os PDFs para assinatura.

**gui.py**: Módulo que contém a interface gráfica do programa.
<br/>
<br/>

### Exemplo de Arquivo Excel
O arquivo Excel deve conter as seguintes colunas:

<table> <thead> <tr> <th>Nome</th> <th>CPF</th> <th>Email</th> <th>Matrícula</th> </tr> </thead> <tbody> <tr> <td>João Silva</td> <td>12345678901</td> <td>joao.silva@email.com</td> <td>12345</td> </tr> <tr> <td>Maria Souza</td> <td>98765432109</td> <td>maria.souza@email.com</td> <td>67890</td> </tr> </tbody> </table> 
