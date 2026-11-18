# 📄 Gerador de Propostas Comerciais Automatizado (PDF + WhatsApp)

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-0F172A?style=for-the-badge)
![ReportLab](https://img.shields.io/badge/PDF-ReportLab-2563EB?style=for-the-badge)
![Pandas](https://img.shields.io/badge/Data-Pandas%20%26%20Excel-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Windows](https://img.shields.io/badge/Build-PyInstaller%20(.exe)-0078D6?style=for-the-badge&logo=windows&logoColor=white)

Aplicação desktop desenvolvida em **Python** para automatizar a geração de propostas comerciais corporativas em PDF a partir de planilhas Excel, com numeração sequencial automática, geração de QR Code institucional e disparo direto via **WhatsApp Web**.

---

## 🎯 Problema Resolvido

Em operações de engenharia e manutenção industrial, a elaboração manual de orçamentos em editores de texto consome tempo, gera erros de cálculo de impostos/margens e dificulta a padronização visual. Este sistema reduz o processo inteiro a **um único clique**, transformando uma tabela de custos base em um documento PDF executivo pronto para envio ao cliente.

---

## ✨ Funcionalidades Principais

- **🖥️ Interface Gráfica Desktop (Tkinter):** Tela limpa e responsiva com validação de campos, seleção de planilhas `.xlsx`, atalho para edição de itens no Excel e processamento assíncrono (`threading`) sem travar a janela.
- **🎨 PDF Corporativo com Capa Executiva (ReportLab):**
  - **Página 1 (Capa):** Design escuro editorial (*Slate 900* & *Royal Blue*), logotipo vetorial, selo oficial da proposta e cartão de identificação exclusivo do cliente.
  - **Página 2 (Orçamento Técnico):** Tabela dinâmica com zebrado automático, alertas técnicos condicionais (ex.: notas específicas para limpeza de tanque), cálculo automático de margem e impostos, valor total por extenso em português (sem bibliotecas externas), cards de garantia/validade, assinatura vetorial e QR Code.
- **🔢 Controle Sequencial Automático:** Persistência do número da proposta (`Nº XXX/AAAA`) via arquivo local (`contador.txt`).
- **📱 Integração com WhatsApp Web:** Validação e padronização automática de números brasileiros (`+55 + DDD + Número`) e abertura direta do WhatsApp Web com mensagem comercial pré-formatada.
- **📦 Distribuição Standalone (`.exe`):** Empacotado com **PyInstaller** para rodar em qualquer máquina Windows sem necessidade de instalar Python.

---

## 🛠️ Tecnologias Utilizadas

| Categoria | Tecnologia / Biblioteca | Finalidade |
| :--- | :--- | :--- |
| **Linguagem** | Python 3.x | Lógica principal, regras de negócio e automação |
| **Interface Gráfica** | `tkinter` | Construção da janela desktop e controles interativos |
| **Manipulação de Dados** | `pandas` & `openpyxl` | Leitura, criação e cálculo de valores a partir do Excel |
| **Geração de PDF** | `reportlab` | Desenho vetorial da capa, tabelas, cards e tipografia |
| **QR Code** | `qrcode` & `Pillow` | Geração dinâmica do QR Code de validação/portfólio |
| **Automação Web** | `urllib.parse` & `webbrowser` | Codificação da mensagem e abertura do WhatsApp Web |
| **Empacotamento** | `PyInstaller` | Compilação em executável único (`.exe`) para Windows |

---

## 📂 Estrutura do Projeto

```text
gerador-propostas-comerciais/
├── gerador_proposta.py              # Código-fonte completo (GUI + Motor PDF + WhatsApp)
├── modelo_planilha_servicos.xlsx    # Planilha base de serviços e custos (gerada automaticamente)
├── contador.txt                     # Controle sequencial do número das propostas
├── requirements.txt                 # Dependências do projeto
└── README.md                        # Documentação do portfólio
```

---

## 🚀 Como Executar o Projeto

### 1. Clonar o repositório e instalar dependências
```bash
git clone https://github.com/seu-usuario/gerador-propostas-comerciais.git
cd gerador-propostas-comerciais
pip install pandas openpyxl reportlab qrcode pillow pyinstaller
```

### 2. Executar via código-fonte
```bash
python gerador_proposta.py
```

### 3. Gerar o Executável Windows (`.exe`)
Para compilar a aplicação desktop sem janela de console em segundo plano:
```bash
pyinstaller --noconfirm --onefile --windowed --name="Gerador_Propostas" --hidden-import="openpyxl" --hidden-import="PIL" --hidden-import="PIL._tkinter_finder" --collect-all="reportlab" gerador_proposta.py
```
O executável final será gerado na pasta `dist/Gerador_Propostas.exe`.

---

## 📊 Estrutura da Planilha Excel (`.xlsx`)

O sistema lê automaticamente qualquer planilha que contenha as duas colunas abaixo (caso não exista, o programa cria um modelo padrão na primeira execução):

| Serviço | Custo_Base |
| :--- | :---: |
| Troca do filtro diesel de alta eficiência | 150.00 |
| Troca do óleo lubrificante sintético (15W40) | 300.00 |
| Limpeza química e descontaminação de tanque de combustível | 500.00 |

---

## 📸 Demonstração Visual

> 💡 **Dica para o seu GitHub:** Crie uma pasta `assets/` no repositório e adicione capturas de tela da interface Tkinter, da Capa Executiva (Página 1) e da Tabela de Orçamento (Página 2) usando os links abaixo:
>
> - `![Interface Desktop](assets/tela_app.png)`
> - `![Capa da Proposta](assets/capa_pdf.png)`
> - `![Orçamento Detalhado](assets/orcamento_pdf.png)`

---

## 👨‍💻 Autor

Desenvolvido por **[Seu Nome]** — Automação de Processos & Desenvolvimento Python.  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Conectar-0A66C2?style=flat-square&logo=linkedin)](https://www.linkedin.com/in/seu-perfil)
[![GitHub](https://img.shields.io/badge/GitHub-Portfólio-181717?style=flat-square&logo=github)](https://github.com/seu-usuario)
