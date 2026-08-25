# 🧠 Sistema Clínico para Psicologia

Sistema web completo para gestão de um consultório de psicologia, desenvolvido em **Google Apps Script** com **Google Sheets** como banco de dados. Nasceu de uma necessidade real: organizar pacientes, agenda, prontuários, sessões e financeiro de um consultório em uma única ferramenta, sem custo de servidor.

> ⚠️ **Projeto real, anonimizado para portfólio.** Dados sensíveis (ID da planilha, nome da profissional e informações de pacientes) foram removidos ou substituídos por placeholders. É um sistema que está de fato em uso no dia a dia de um consultório.

---

## ✨ Funcionalidades

- 🔐 **Autenticação e perfis de acesso** — login, logout, redefinição de senha e gestão de usuários com diferentes níveis de permissão.
- 📊 **Dashboard** — visão geral com indicadores e gráficos (sessões, pacientes, financeiro) usando Chart.js.
- 👥 **Gestão de pacientes** — cadastro, edição, exclusão e consulta de histórico completo.
- 📅 **Agenda** — calendário de sessões interativo (FullCalendar), com agendamento e acompanhamento do status (agendado / realizado).
- 📝 **Prontuário clínico** — registro e histórico de evoluções por paciente, com busca e filtros.
- 🧾 **Módulo financeiro e fiscal** — controle de recebimentos e emissão de **recibos em PDF** gerados automaticamente.
- 📄 **Geração de documentos** — criação de PDFs (recibos e relatórios) via Google Docs/Drive.
- 🗂️ **Logs de auditoria** — registro das ações realizadas no sistema.

---

## 🛠️ Tecnologias utilizadas

| Camada | Tecnologias |
|---|---|
| **Back-end** | Google Apps Script (JavaScript / runtime V8) |
| **Front-end** | HTML5, CSS3, JavaScript |
| **Banco de dados** | Google Sheets |
| **Integrações Google** | Google Drive, Google Docs, Google Spreadsheets APIs |
| **Bibliotecas** | Chart.js (gráficos), FullCalendar (agenda), Font Awesome (ícones) |

---

## 🏗️ Como funciona a arquitetura

O sistema é uma **Single Page Application (SPA)** servida pelo Google Apps Script:

- **`Servidor.gs`** — todo o back-end: funções de login, CRUD de pacientes, sessões, prontuários, financeiro e geração de PDFs. É a camada que conversa com a planilha do Google Sheets (o "banco de dados").
- **`index.html`** — a página principal (tela de login e estrutura base da aplicação).
- **`styles.html`** — os estilos (CSS) da aplicação.
- **`tela_login.html`** — a interface de login.
- **`js_*.html`** — os módulos do front-end. 

> 💡 **Sobre os arquivos `.html`:** no Google Apps Script, o código JavaScript do lado do cliente é servido dentro de arquivos `.html` (dentro de tags `<script>`). Por isso os módulos de front-end (`js_pacientes`, `js_agenda`, `js_financeiro`, etc.) têm extensão `.html` — é a convenção da plataforma, não são páginas HTML tradicionais.

### Módulos do front-end

| Arquivo | Responsabilidade |
|---|---|
| `js_main.html` | Navegação principal e estrutura da interface |
| `js_dashboard.html` | Painel de indicadores e gráficos |
| `js_pacientes.html` | Cadastro e gestão de pacientes |
| `js_agenda.html` | Calendário e agendamento de sessões |
| `js_sessoes.html` | Registro e acompanhamento de sessões |
| `js_prontuario.html` | Prontuário e evoluções clínicas |
| `js_financeiro.html` | Controle financeiro e fiscal |
| `js_documentos.html` | Geração de documentos/recibos |
| `js_usuarios.html` | Gestão de usuários |
| `js_config.html` | Configurações do sistema |
| `js_logs.html` | Registro de logs/auditoria |

---

## 🚀 Como executar / implantar

Este projeto roda dentro do ecossistema Google. Para colocá-lo no ar:

1. Crie uma **planilha no Google Sheets** que servirá de banco de dados (com as abas de pacientes, usuários, sessões, financeiro, etc.).
2. Na planilha, vá em **Extensões → Apps Script** e cole os arquivos deste repositório no editor.
3. No arquivo `Servidor.gs`, substitua o placeholder `COLE_AQUI_O_ID_DA_SUA_PLANILHA` pelo **ID da sua planilha** (encontrado na URL da planilha).
4. Clique em **Implantar → Nova implantação → Aplicativo da Web**.
5. Acesse pela URL gerada e faça login.

---

## 📌 Destaques técnicos

- Desenvolvimento **full-stack** (back-end + front-end) em uma única plataforma.
- Integração com múltiplas APIs do Google (Sheets, Drive, Docs).
- **Geração dinâmica de PDFs** a partir de dados da planilha.
- Sistema de **autenticação com controle de permissões** por perfil.
- Interface responsiva e organizada em módulos.

---

## 👨‍💻 Autor

**Alisson Miranda** — Desenvolvedor de Software | Analista de Sistemas
Formado em Análise e Desenvolvimento de Sistemas.

🔗 [LinkedIn](https://www.linkedin.com/in/alisson-miranda-9742862a8)
