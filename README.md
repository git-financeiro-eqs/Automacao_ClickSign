# Automação de Termo de Recebimento

Este projeto é uma automação para gerar e enviar termos de recebimento de cartões corporativos para os funcionários da empresa. O programa é composto por três módulos principais:

- **geradorFormularios.py**: Gera PDFs personalizados com os termos de recebimento para cada funcionário.
- **clicksign.py**: Envia os PDFs gerados para assinatura via e-mail utilizando a plataforma ClickSign.
- **gui.py**: Interface gráfica para interação com o usuário, permitindo a seleção de arquivos e execução das funcionalidades.

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

## Instalação

1. Clone o repositório:

   ```bash
   git clone https://github.com/seu-usuario/automacao-termo-recebimento.git
   cd automacao-termo-recebimento
