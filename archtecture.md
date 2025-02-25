# Arquitetura do Projeto

## Visão Geral

O programa é dividido em três módulos principais, cada um com uma responsabilidade específica:

### 1. Interface Gráfica (`gui.py`)

- **Responsabilidade**: Interagir com o usuário, coletando dados e acionando as funcionalidades dos outros módulos.
- **Tecnologia Utilizada**: Utiliza a biblioteca `tkinter` para criar a interface gráfica.

### 2. Geração de PDFs (`geradorFormularios.py`)

- **Responsabilidade**: Gerar os termos de recebimento em formato PDF para cada funcionário.
- **Tecnologia Utilizada**: Utiliza a biblioteca `reportlab` para criar os PDFs com base em um template pré-definido.
- **Armazenamento**: Os PDFs são armazenados na pasta `PDFs/`.

### 3. Envio de E-mails (`clicksign.py`)

- **Responsabilidade**: Enviar os PDFs gerados para assinatura via e-mail utilizando a plataforma ClickSign.
- **Tecnologia Utilizada**: Utiliza a biblioteca `selenium` para automatizar o processo de login e envio de documentos na plataforma ClickSign.

---

## Fluxo de Execução

1. O usuário insere a data e seleciona o arquivo Excel contendo os dados dos funcionários na interface gráfica.
2. Ao clicar em **"Gerar PDFs"**, o módulo `geradorFormularios.py` é acionado, gerando os PDFs e armazenando-os na pasta `PDFs/`.
3. Ao clicar em **"Enviar E-mail"**, o módulo `clicksign.py` é acionado, enviando os PDFs para assinatura via e-mail utilizando a plataforma ClickSign.

---

## Diagrama de Arquitetura

```plaintext
+-------------------+       +-------------------+       +-------------------+
|   Interface       |       |   Geração de      |       |   Envio de        |
|   Gráfica         |       |   PDFs            |       |   E-mails         |
|   (gui.py)        |       |   (geradorFormularios.py) |   (clicksign.py)  |
+-------------------+       +-------------------+       +-------------------+
        |                           |                           |
        |                           |                           |
        v                           v                           v
+-------------------+       +-------------------+       +-------------------+
|   Entrada de      |       |   Geração de      |       |   Envio de        |
|   Dados (Excel)   |       |   PDFs (PDFs/)    |       |   E-mails (ClickSign)|
+-------------------+       +-------------------+       +-------------------+
