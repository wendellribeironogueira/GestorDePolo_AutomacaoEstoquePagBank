# GestorDePolo_AutomacaoEstoquePagBank

**Automatiza o controle de estoque logístico local, análise de dados operacionais e validação de reversas através de VBA no Excel**

![Banner do Projeto](banner.png)

> ⚠️ Este projeto não contém dados reais da empresa. Toda informação sensível foi removida antes da publicação.

---

## 🔖 Visão Geral

O **Gestor de Polo** é uma solução automatizada desenvolvida em **VBA para Excel**, criada durante o Projeto **PagResolve** na **PagBank**. Sua função é otimizar o processo manual de controle de estoque feito por auxiliares logísticos e gerar relatórios estatísticos com base em arquivos CSV exportados da plataforma **Workfinity (iSolution)**.

Este sistema eliminou horas de lançamentos manuais diários, padronizando processos e reduzindo erros.

---

## 🔹 Funcionalidades Principais

O sistema está dividido em **3 módulos** principais:

### ✅ Módulo IMPORTAR
- Importa automaticamente dados de um **arquivo CSV** exportado do Workfinity.
- Atualiza a planilha **ESTOQUE.xlsm**, alterando status de equipamentos para "Ativado" quando finalizados.
- Registra equipamentos substituídos na aba **REVERSA**, marcando como defeituosos.
- Evita duplicidade de registros através de controle com `Scripting.Dictionary`.
- Gera contadores com resumo da operação.

### ✅ Módulo REVERSA
- Permite **validação de seriais defeituosos** retirados do cliente via TextBox interativa.
- Compara o serial com a planilha REVERSA e exibe gráficos de confirmação (tique verde ou X vermelho).
- Ideal para uso com leitores de código de barras.

### ✅ Módulo RELATÓRIO
- Gera painel completo com dados operacionais a partir de um CSV.
- Métricas geradas:
  - Chamados por status, técnico, cidade, tipo de serviço
  - SLA (dentro e fora do prazo)
  - Equipamentos mais utilizados
  - Chamados vencendo hoje
  - Alertas de reabertura
  - Horários fora do expediente
- Organiza os dados automaticamente em tabelas e dashboards prontos para apresentação.

---

## 🛠️ Como Utilizar

1. Baixe os arquivos `.bas` e importe-os para seu projeto VBA.
2. Tenha em seu diretório o arquivo **ESTOQUE.xlsm** com as planilhas:
   - `ESTOQUE`
   - `REVERSA`
   - `RECEBIMENTO`
3. Exporte o CSV da plataforma **Workfinity** manualmente.
4. Execute o módulo `SelecionarArquivo` no IMPORTAR para carregar o CSV.
5. Rode `AtualizarPlanilhasDeOutroArquivo` para aplicar as atualizações.
6. Use o botão ou pressione Enter no TextBox para validar reversas.
7. Rode o `GerarPainelDeInformacoes` no módulo RELATÓRIO para gerar os dashboards.

---

## 📊 Impacto

- Redução de **mais de 90%** do tempo gasto com controle de estoque e devolução.
- Elimina retrabalho e inconsistências de registros manuais.
- Melhoria na gestão e visibilidade de dados operacionais em tempo real.

---

## 🧱 Tecnologias Utilizadas

- VBA para Excel (Visual Basic for Applications)
- Dicionários `Scripting.Dictionary`
- Gráficos e Tabelas Dinâmicas automatizadas
- Interface com `ActiveX TextBox` e `Shapes`
- Arquivos `.CSV` padronizados (Workfinity)

---

## 🔒 Considerações

- O sistema é 100% local, mas pode futuramente ser adaptado para **API do Workfinity**.
- Requer configuração de permissão de macros no Excel para execução.

---

## 📄 Licença

Este projeto é disponibilizado sob a licença **Creative Commons BY-NC-ND 4.0**.  
Foi desenvolvido por iniciativa própria durante meu tempo livre como colaborador da empresa PagBank, com o objetivo de otimizar rotinas operacionais.  
Seu uso é permitido apenas para fins educacionais, de demonstração ou portfólio técnico.  
Não é permitido uso comercial ou modificação sem autorização prévia.  

[Saiba mais sobre a licença](https://creativecommons.org/licenses/by-nc-nd/4.0/deed.pt-br)

---

## 🧾 Registro de Autoria

Este projeto foi originalmente publicado por mim no GitHub corporativo da PagBank, com o objetivo de registrar sua autoria e facilitar sua continuidade interna.  
**Versão original (acesso restrito a colaboradores):**  
[https://github.com/W3NR1/GestorDePolo](https://github.com/W3NR1/GestorDePolo)

---

## 👨‍💼 Autor

**Wendell Ribeiro Nogueira**  
Especialista em Suporte, Infraestrutura e Automação  
[GitHub pessoal](https://github.com/wendellribeironogueira)  
[GitHub corporativo (PagBank)](https://github.com/W3NR1)  
[LinkedIn](https://www.linkedin.com/in/wendell-ribeiro-nogueira)
