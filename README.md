# Automação de Termo de Recebimento

Este projeto é uma automação para gerar e enviar termos de recebimento de cartões corporativos para os funcionários da empresa. O programa é composto por três módulos principais:

- **geradorFormularios.py**: Gera PDFs personalizados com os termos de recebimento para cada funcionário.
- **clicksign.py**: Envia os PDFs gerados para assinatura via e-mail utilizando a plataforma ClickSign.
- **gui.py**: Interface gráfica para interação com o usuário, permitindo a seleção de arquivos e execução das funcionalidades.
<br/>
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
Instale as dependências:

bash
Copy
pip install -r requirements.txt
Certifique-se de ter o ChromeDriver instalado e configurado para o Selenium.

Uso
Execute o programa:

bash
Copy
python gui.py
Na interface gráfica:

Insira a data no formato DD/MM/AAAA.

Selecione o arquivo Excel contendo os dados dos funcionários.

Clique em "Gerar PDFs" para criar os termos de recebimento.

Clique em "Enviar E-mail" para enviar os termos para assinatura via ClickSign.

Estrutura do Projeto
PDFs/: Pasta onde os PDFs gerados são armazenados.

Imagens/: Contém as imagens utilizadas na interface gráfica.

geradorFormularios.py: Módulo responsável por gerar os PDFs.

clicksign.py: Módulo responsável por enviar os PDFs para assinatura.

gui.py: Módulo que contém a interface gráfica do programa.

Exemplo de Arquivo Excel
O arquivo Excel deve conter as seguintes colunas:

Nome	CPF	Email	Matrícula
João Silva	12345678901	joao.silva@email.com	12345
Maria Souza	98765432109	maria.souza@email.com	67890
