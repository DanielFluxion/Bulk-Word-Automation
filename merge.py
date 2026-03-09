import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import pandas as pd

# --- CONFIG ---
st.set_page_config(page_title="Gerador de Eficácia", layout="wide")

def local_css(file_name):
    with open(file_name, encoding="utf-8") as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

local_css("style.css")

st.markdown("""
<div class="custom-header">
    <div class="header-content-wrapper">
        <span class="header-icon">📋</span>
        <span class="header-title-text">Avaliação de Eficácia</span>
    </div>
</div>
""", unsafe_allow_html=True)

# --- CARREGA DADOS ---
@st.cache_data
def carregar_dados():
    try:
        tabs = pd.read_excel("base de treinamentos (1).xlsx", sheet_name=None)
        df_c = tabs["Colaboradores"]
        df_b = tabs["Base de Treinamentos"]
        df_c['Matrícula'] = df_c['Matrícula'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_b['Matrícula'] = df_b['Matrícula'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        return df_c, df_b
    except Exception as e:
        st.error(f"Erro ao carregar Excel: {e}")
        return None, None

df_colaboradores, df_base = carregar_dados()

# --- CARD: COLABORADOR / AVALIADOR ---
card = st.container()
with card:
    st.markdown('<div class="main-card-title">Dados do Colaborador/Avaliador</div>', unsafe_allow_html=True)

    col_m1, col_m2 = st.columns([1, 2])
    with col_m1:
        matricula_input = st.text_input("Digite a Matrícula", placeholder="Ex: 1001", key="mat_search")

    # FIX: Força atualização via session_state quando matrícula muda
    if matricula_input and df_colaboradores is not None:
        res_colab = df_colaboradores[df_colaboradores['Matrícula'] == matricula_input.strip()]
        if not res_colab.empty:
            nome_encontrado  = res_colab.iloc[0]['Nome do Colaborador']
            cargo_encontrado = res_colab.iloc[0]['Cargo']
            area_encontrada  = res_colab.iloc[0]['Setor']

            # Só atualiza se mudou (evita loop)
            if st.session_state.get("ultima_matricula") != matricula_input:
                st.session_state["ultima_matricula"] = matricula_input
                st.session_state["n1"] = nome_encontrado
                st.session_state["c1"] = cargo_encontrado
                st.session_state["s1"] = area_encontrada

            with col_m2:
                st.write("")
                st.write("")
                st.success(f"✅ {nome_encontrado}")
        else:
            # Limpa os campos se matrícula não encontrada
            if st.session_state.get("ultima_matricula") != matricula_input:
                st.session_state["ultima_matricula"] = matricula_input
                st.session_state["n1"] = ""
                st.session_state["c1"] = ""
                st.session_state["s1"] = ""

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sub-header">👤 Dados do Colaborador</div>', unsafe_allow_html=True)
        c_nome  = st.text_input("Nome", key="n1")
        c_cargo = st.text_input("Cargo", key="c1")
        c_area  = st.text_input("Área/Setor", key="s1")

    with c2:
        st.markdown('<div class="sub-header">📝 Dados do Avaliador</div>', unsafe_allow_html=True)
        a_nome  = st.text_input("Nome do Avaliador", key="n2")
        a_cargo = st.text_input("Cargo do Avaliador", key="c2")
        a_area  = st.text_input("Área do Avaliador", key="s2")

# --- CARD: TREINAMENTO ---
training_card = st.container()
with training_card:
    st.markdown('<div class="training-card-title">🎓 Detalhes do Treinamento</div>', unsafe_allow_html=True)

    t_cod_busca = st.text_input("Digite o Código para buscar o treinamento (Ex: PCG 0001)")

    t_nome = t_periodo = ""

    if t_cod_busca and matricula_input and df_base is not None:
        res_treino = df_base[
            (df_base['Matrícula'] == matricula_input.strip()) &
            (df_base['Código do Procedimento'].str.contains(t_cod_busca.strip(), case=False, na=False))
        ]
        if not res_treino.empty:
            t_nome    = res_treino.iloc[0]['Procedimento']
            data_br   = pd.to_datetime(res_treino.iloc[0]['Data do Treinamento'])
            t_periodo = data_br.strftime('%d/%m/%Y')
            st.success("✅ Treinamento encontrado!")
        else:
            st.info("💡 Este código não foi encontrado para esta matrícula.")

    t_nome_final    = st.text_input("Ação de Desenvolvimento e Treinamento", value=t_nome)
    col_t1, col_t2, col_t3 = st.columns(3)
    t_periodo_final = col_t1.text_input("Período (Data)", value=t_periodo)
    t_inst          = col_t2.text_input("Instituição", value="DESSMA")
    t_local         = col_t3.text_input("Local", value="SALA DE REUNIÃO DESSMA")

# --- GERA O DOCUMENTO ---
campos_obrigatorios = [c_nome, a_nome, t_nome_final, t_periodo_final]
todos_preenchidos   = all(str(campo).strip() != "" for campo in campos_obrigatorios)

st.divider()

buffer = None
if todos_preenchidos:
    try:
        base_path     = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_path, "template.docx")

        if os.path.exists(template_path):
            doc = DocxTemplate(template_path)
            contexto = {
                'c_nome': c_nome, 'c_cargo': c_cargo, 'c_area': c_area,
                'a_nome': a_nome, 'a_cargo': a_cargo, 'a_area': a_area,
                't_nome': t_nome_final,
                # FIX: usa string vazia se código não foi buscado
                't_codigo': t_cod_busca if t_cod_busca else "",
                't_periodo': t_periodo_final, 't_inst': t_inst, 't_local': t_local
            }
            doc.render(contexto)
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
        else:
            st.error("Arquivo 'template.docx' não encontrado.")
    except Exception as e:
        st.error(f"Erro ao processar documento: {e}")

col1, col2, col3 = st.columns([3, 1, 3])
with col2:
    if buffer is not None:
        st.download_button(
            label="📥 Exportar Word",
            data=buffer,
            file_name=f"Eficacia_{c_nome}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="btn_export_topo"
        )
    else:
        st.button("📥 Exportar Word", disabled=True, use_container_width=True)