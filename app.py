import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import zipfile
import pandas as pd

# --- CONFIG ---
st.set_page_config(page_title="Gerador de Eficácia", layout="wide")

def local_css(file_name):
    css_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)
    if os.path.exists(css_path):
        with open(css_path, encoding="utf-8") as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

local_css("style.css")

st.markdown("""
<div class="custom-header">
    <div class="header-content-wrapper">
        <span class="header-icon">📋</span>
        <span class="header-title-text">Avaliação de Treinamento</span>
    </div>
</div>
""", unsafe_allow_html=True)

# --- CARREGA DADOS ---
@st.cache_data
def carregar_dados():
    try:
        nome_arquivo = "base de treinamentos (1).xlsx"
        # 1. Carrega o Excel
        tabs = pd.read_excel(nome_arquivo, sheet_name=None)
        
        # 2. Pega a primeira aba disponível (já que só tem uma, não importa o nome)
        primeira_aba = list(tabs.keys())[0]
        df = tabs[primeira_aba]
        
        # 3. LIMPEZA CRÍTICA: Remove espaços extras dos nomes das colunas
        # Isso resolve o erro se no Excel estiver "Setor " com um espaço no fim
        df.columns = [str(col).strip() for col in df.columns]
        
        # 4. Limpeza das matrículas
        df['Matrícula'] = df['Matrícula'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        # Retornamos o mesmo dataframe para ambos, já que os dados estão juntos
        return df, df
        
    except Exception as e:
        st.error(f"Erro ao carregar Excel: {e}")
        return None, None
    
df_colaboradores, df_base = carregar_dados()
# ─────────────────────────────────────────────
# CARD 1 — COLABORADOR
# ─────────────────────────────────────────────
st.markdown('<div class="main-card-title">Dados do Colaborador</div>', unsafe_allow_html=True)

col_m1, col_m2 = st.columns([1, 2])
with col_m1:
    matricula_input = st.text_input("Digite a Matrícula", placeholder="Ex: 1001", key="mat_search")

if matricula_input and df_colaboradores is not None:
    res_colab = df_colaboradores[df_colaboradores['Matrícula'] == matricula_input.strip()]
    if not res_colab.empty:
        nome_encontrado  = res_colab.iloc[0]['Nome do Colaborador']
        cargo_encontrado = res_colab.iloc[0]['Cargo']
        area_encontrada  = res_colab.iloc[0]['Setor']

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
        if st.session_state.get("ultima_matricula") != matricula_input:
            st.session_state["ultima_matricula"] = matricula_input
            st.session_state["n1"] = ""
            st.session_state["c1"] = ""
            st.session_state["s1"] = ""
        with col_m2:
            st.write("")
            st.write("")
            st.warning("⚠️ Matrícula não encontrada.")

st.markdown('<div class="sub-header">👤 Dados do Colaborador</div>', unsafe_allow_html=True)
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    c_nome  = st.text_input("Nome",       key="n1")
with col_f2:
    c_cargo = st.text_input("Cargo",      key="c1")
with col_f3:
    c_area  = st.text_input("Area/Setor", key="s1")

# ─────────────────────────────────────────────
# CARD 2 — SELEÇÃO DE TREINAMENTOS
# ─────────────────────────────────────────────
st.markdown('<div class="training-card-title">🎓 Treinamentos do Colaborador</div>', unsafe_allow_html=True)

treinamentos_selecionados = []   # lista de dicts com dados de cada treinamento selecionado

if not matricula_input or not c_nome:
    st.info("💡 Preencha a matrícula e os dados do colaborador para ver os treinamentos disponíveis.")
elif df_base is None:
    st.error("Base de treinamentos não carregada.")
else:
    df_treinos_colab = df_base[df_base['Matrícula'] == matricula_input.strip()].copy()

    if df_treinos_colab.empty:
        st.warning("Nenhum treinamento encontrado para esta matrícula.")
    else:
        # Formata data para exibição
        df_treinos_colab['Data Formatada'] = pd.to_datetime(
            df_treinos_colab['Data do Treinamento'], errors='coerce'
        ).dt.strftime('%d/%m/%Y')

        # Cria label único para o multiselect
        df_treinos_colab['label'] = (
            df_treinos_colab['Código do Procedimento'].fillna('') + ' — ' +
            df_treinos_colab['Procedimento'].fillna('') + ' (' +
            df_treinos_colab['Data Formatada'].fillna('') + ')'
        )

        opcoes = df_treinos_colab['label'].tolist()

        selecao = st.multiselect(
            f"Selecione os treinamentos de **{c_nome}** para gerar os documentos:",
            options=opcoes,
            placeholder="Clique para selecionar um ou mais treinamentos...",
        )

        if selecao:
            st.markdown(f"**{len(selecao)} treinamento(s) selecionado(s)**")

            # ── Para cada treinamento selecionado, permite editar inst/local e avaliador ──
            st.markdown("---")
            st.markdown("#### ✏️ Confirme os dados de cada treinamento selecionado")
            st.caption("Os campos abaixo são pré-preenchidos automaticamente. Edite se necessário.")

            for i, label in enumerate(selecao):
                row = df_treinos_colab[df_treinos_colab['label'] == label].iloc[0]

                with st.expander(f"📄 {label}", expanded=(len(selecao) == 1)):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        inst  = st.text_input("Instituição", value="DESSMA",               key=f"inst_{i}")
                        local = st.text_input("Local",       value="SALA DE REUNIÃO DESSMA", key=f"local_{i}")
                    with col_b:
                        av_nome  = st.text_input("Nome do Avaliador",  value="", key=f"av_nome_{i}")
                        av_cargo = st.text_input("Cargo do Avaliador", value="", key=f"av_cargo_{i}")
                        av_area  = st.text_input("Área do Avaliador",  value="", key=f"av_area_{i}")

                treinamentos_selecionados.append({
                    "t_nome":    row['Procedimento'],
                    "t_codigo":  row['Código do Procedimento'],
                    "t_periodo": row['Data Formatada'],
                    "t_inst":    inst,
                    "t_local":   local,
                    "a_nome":    av_nome,
                    "a_cargo":   av_cargo,
                    "a_area":    av_area,
                })

# ─────────────────────────────────────────────
# GERAÇÃO DO ZIP
# ─────────────────────────────────────────────
st.divider()

campos_base_ok = bool(str(c_nome).strip())
pode_gerar     = campos_base_ok and len(treinamentos_selecionados) > 0

col1, col2, col3 = st.columns([3, 1, 3])
with col2:
    if not pode_gerar:
        st.button("📥 Exportar ZIP", disabled=True, use_container_width=True)
    else:
        base_path     = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_path, "template.docx")

        if not os.path.exists(template_path):
            st.error("Arquivo 'template.docx' não encontrado.")
        else:
            try:
                zip_buffer = io.BytesIO()
                nomes_usados = {}

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for treino in treinamentos_selecionados:
                        doc = DocxTemplate(template_path)
                        contexto = {
                            'c_nome':    c_nome,
                            'c_cargo':   c_cargo,
                            'c_area':    c_area,
                            'a_nome':    treino['a_nome'],
                            'a_cargo':   treino['a_cargo'],
                            'a_area':    treino['a_area'],
                            't_nome':    treino['t_nome'],
                            't_codigo':  treino['t_codigo'],
                            't_periodo': treino['t_periodo'],
                            't_inst':    treino['t_inst'],
                            't_local':   treino['t_local'],
                        }
                        doc.render(contexto)

                        # Nome do arquivo dentro do ZIP (garante unicidade)
                        codigo_limpo = str(treino['t_codigo']).replace("/", "-").replace(" ", "_")
                        nome_base    = f"Eficacia_{c_nome}_{codigo_limpo}.docx"
                        if nome_base in nomes_usados:
                            nomes_usados[nome_base] += 1
                            nome_base = nome_base.replace(".docx", f"_{nomes_usados[nome_base]}.docx")
                        else:
                            nomes_usados[nome_base] = 1

                        doc_buffer = io.BytesIO()
                        doc.save(doc_buffer)
                        doc_buffer.seek(0)
                        zf.writestr(nome_base, doc_buffer.read())

                zip_buffer.seek(0)

                label_btn = (
                    f"📥 Exportar {len(treinamentos_selecionados)} documento(s) (.zip)"
                    if len(treinamentos_selecionados) > 1
                    else "📥 Exportar Documento (.docx no .zip)"
                )

                st.download_button(
                    label=label_btn,
                    data=zip_buffer,
                    file_name=f"Eficacia_{c_nome}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="btn_export_zip"
                )

            except Exception as e:
                st.error(f"Erro ao gerar documentos: {e}")
