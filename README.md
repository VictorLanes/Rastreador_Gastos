# 💰 Rastreador_Gastos

Aplicação desenvolvida em **Python** com interface gráfica moderna para gerenciar suas **despesas**, **cartões de crédito** e **orçamentos mensais** de forma prática e eficiente.

---

## 🚀 Funcionalidades Principais

- ✅ Controle de despesas com categorias e formas de pagamento.
- ✅ Cadastro e gerenciamento de cartões de crédito.
- ✅ Previsão de faturas com base no ciclo de fechamento do cartão.
- ✅ Relatórios automáticos por categoria e total geral.
- ✅ Exportação dos dados para **Excel** com formatação profissional.
- ✅ Backup e restauração do banco de dados local.
- ✅ Indicadores visuais para controle do orçamento mensal.
- ✅ Alternância entre **Modo Claro** e **Modo Escuro**.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3**
- **Tkinter** (`ttkbootstrap`) para interface gráfica moderna.
- **SQLite3** para banco de dados local.
- **OpenPyXL** para geração e formatação de planilhas Excel.
- **Shutil** para operações de backup e restauração.
- **Datetime** para manipulação de datas e previsões.

---

## 📦 Instalação

### ✅ Pré-requisitos
- Python 3.8 ou superior instalado.
- 
###📂 Estrutura de Arquivos
  📁 Projeto
   ├── financeiro.db        # Banco de dados local (gerado automaticamente)
   ├── seu_arquivo.py       # Arquivo principal do projeto
   ├── backups/             # Pasta opcional para armazenar backups
   └── README.md            # Documentação do projeto

### ✅ Instalação das dependências
```bash
pip install ttkbootstrap openpyxl
