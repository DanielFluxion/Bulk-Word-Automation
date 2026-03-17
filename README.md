# 🚀 Automação de Avaliação de Treinamentos (Bulk Word Automation)

Este projeto é uma aplicação web interativa desenvolvida com **Python** e **Streamlit**. Ele resolve um problema comum de RH e Treinamento: a busca manual de dados em planilhas e o preenchimento repetitivo de formulários em Word.

## 📋 Funcionalidades
- **Busca Inteligente:** Localiza dados de colaboradores (Nome, Cargo, Setor) instantaneamente através da Matrícula.
- **Integração Excel:** Consome dados diretamente de arquivos `.xlsx` usando a biblioteca Pandas.
- **Automação de Documentos:** Gera documentos Word personalizados (`.docx`) preenchendo templates automaticamente com os dados encontrados.
- **Interface Intuitiva:** UI moderna e limpa desenvolvida para facilitar a experiência do usuário final.

## 🛠️ Tecnologias Utilizadas
- **Python:** Linguagem base do projeto.
- **Streamlit:** Framework para criação da interface web.
- **Pandas:** Manipulação e tratamento de dados da planilha Excel.
- **Docxtpl:** Automação de templates Word (Jinja2 para documentos).
- **Openpyxl:** Motor de leitura para arquivos Excel modernos.

## 📂 Estrutura do Projeto
- `app.py`: O código principal da aplicação.
- `base de treinamentos (1).xlsx`: Banco de dados dos colaboradores e treinamentos.
- `template.docx`: Modelo de documento com tags `{{campo}}` para preenchimento automático.
- `style.css`: Personalização visual da interface.
- `requirements.txt`: Lista de dependências para rodar o projeto.

## 🚀 Como Rodar o Projeto
1. Clone o repositório:
   ```bash
   git clone [https://github.com/DanielFluxion/Bulk-Word-Automation.git](https://github.com/DanielFluxion/Bulk-Word-Automation.git)

2. Crie e ative seu ambiente virtual (opcional, mas recomendado):
   ```bash
   python -m venv .venv
   .venv\Scripts\activate

3. Instale as bibliotecas necessárias:
   ```bash
   pip install -r requirements.txt

4. Inicie a aplicação:
   ```bash
   streamlit run app.py
