# Arquitetura do Projeto
<br/>

O programa é dividido em três módulos principais, cada um com uma responsabilidade específica:
<br/>
<br/>
<table>
  <thead>
    <tr>
      <th>Módulo</th>
      <th>Descrição</th>
      <th>Principais Responsabilidades</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Interface Gráfica (gui.py)</strong></td>
      <td>Módulo responsável pela interação com o usuário, coletando dados e acionando as funcionalidades dos outros módulos.</td>
      <td>
        <ul>
          <li>Coletar dados do usuário (data e arquivo Excel).</li>
          <li>Acionar a geração de PDFs e o envio de e-mails.</li>
          <li>Fornecer uma interface amigável para o usuário.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td><strong>Geração de PDFs (geradorFormularios.py)</strong></td>
      <td>Módulo responsável por gerar os termos de recebimento em formato PDF para cada funcionário.</td>
      <td>
        <ul>
          <li>Ler dados do arquivo Excel.</li>
          <li>Gerar PDFs personalizados com base em um template.</li>
          <li>Armazenar os PDFs na pasta <code>PDFs/</code>.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td><strong>Envio de E-mails (clicksign.py)</strong></td>
      <td>Módulo responsável por enviar os PDFs gerados para assinatura via e-mail utilizando a plataforma ClickSign.</td>
      <td>
        <ul>
          <li>Automatizar o login na plataforma ClickSign.</li>
          <li>Enviar os PDFs para assinatura via e-mail.</li>
          <li>Gerenciar o fluxo de envio e confirmação.</li>
        </ul>
      </td>
    </tr>
  </tbody>
</table>

<br/>

## Fluxo de Execução

1. O usuário insere a data e seleciona o arquivo Excel contendo os dados dos funcionários na interface gráfica.
2. Ao clicar em **"Gerar PDFs"**, o módulo `geradorFormularios.py` é acionado, gerando os PDFs e armazenando-os na pasta `PDFs/`.
3. Ao clicar em **"Enviar E-mail"**, o módulo `clicksign.py` é acionado, enviando os PDFs para assinatura via e-mail utilizando a plataforma ClickSign.

---
