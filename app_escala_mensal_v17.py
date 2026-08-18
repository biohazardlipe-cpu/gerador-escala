import streamlit as st
import pandas as pd
import calendar
import json
import os
from datetime import datetime
from collections import defaultdict
from fpdf import FPDF

# CONFIGURAÇÕES DO SUPERMERCADO
NOME_SUPERMERCADO = "Supermercado Economico"
NOME_CRIADOR = "Felipe Silva"
CARGO_CRIADOR = "Responsável Setor T.I"
ARQUIVO_DADOS = "dados_escala.json"
COR_PRINCIPAL = "#FF6F00" # Laranja
COR_FUNDO = "#FFFFFF" # Branco

FUNCOES_PADRAO = ["Caixa", "Repositor", "Açougue", "Padaria", "Hortifruti", "Atendimento", "Estoque", "Fiscal", "Gerente"]
DIAS_SEMANA = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
DIAS_SEMANA_CURTO = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
HORAS_POR_TURNO = 8
MAX_DIAS_SEGUIDOS = 6
MOTIVOS_PADRAO = ["Troca entre funcionários", "Atestado Médico", "Folga", "Falta", "Férias", "Outro"]

# FUNÇÕES SALVAR E CARREGAR
def salvar_dados():
    dados = {
        "funcionarios": st.session_state.funcionarios,
        "modelo_dia": st.session_state.modelo_dia,
        "feriados": st.session_state.feriados,
        "cargos": st.session_state.cargos,
        "historico_trocas": st.session_state.historico_trocas
    }
    with open(ARQUIVO_DADOS, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)
    st.success("Dados salvos com sucesso!")

def carregar_dados():
    if os.path.exists(ARQUIVO_DADOS):
        with open(ARQUIVO_DADOS, "r", encoding="utf-8") as f:
            dados = json.load(f)
        st.session_state.funcionarios = dados.get("funcionarios", {})
        st.session_state.modelo_dia = dados.get("modelo_dia", {})
        st.session_state.feriados = dados.get("feriados", [])
        st.session_state.cargos = dados.get("cargos", FUNCOES_PADRAO.copy())
        st.session_state.historico_trocas = dados.get("historico_trocas", [])
        st.success("Dados carregados!")
    else:
        st.warning("Nenhum arquivo de dados encontrado.")

# INICIALIZAR SESSION STATE
if 'funcionarios' not in st.session_state: st.session_state.funcionarios = {}
if 'modelo_dia' not in st.session_state: st.session_state.modelo_dia = {}
if 'feriados' not in st.session_state: st.session_state.feriados = []
if 'cargos' not in st.session_state: st.session_state.cargos = FUNCOES_PADRAO.copy()
if 'historico_trocas' not in st.session_state: st.session_state.historico_trocas = []

# TEMA LARANJA E BRANCO
st.set_page_config(page_title=f"Escala - {NOME_SUPERMERCADO}", layout="wide")
st.markdown(f"""
    <style>
.stApp {{ background-color: {COR_FUNDO}; }}
    h1, h2, h3 {{ color: {COR_PRINCIPAL}; }}
.stButton>button {{ background-color: {COR_PRINCIPAL}; color: white; border-radius: 8px; border: none;}}
.stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {{ font-size:16px; font-weight: bold; color: {COR_PRINCIPAL};}}
.footer {{ text-align: center; color: gray; font-size: 12px; margin-top: 30px;}}
    </style>
    """, unsafe_allow_html=True)

st.title(f"🛒 Gerador de Escala MENSAL")
st.subheader(f"{NOME_SUPERMERCADO}")
st.caption(f"**Desenvolvido por:** {NOME_CRIADOR} - {CARGO_CRIADOR}")

# BARRA DE SALVAR/CARREGAR NO TOPO
col1, col2, col3 = st.columns([1,1,4])
with col1:
    if st.button("💾 Salvar Dados"):
        salvar_dados()
with col2:
    if st.button("📂 Carregar Dados"):
        carregar_dados()

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["1. Funcionários + Excel", "2. Modelo + Cargos + Feriados", "3. Calendário", "4. Gerar e Exportar", "5. Ver e Trocar Escala", "6. Sobre"])

# 1. ABA FUNCIONÁRIOS
with tab1:
    st.header("1. Importar Funcionários do Excel")
    st.info("Colunas: Nome | Cargo | Folgas_Semana")
    arquivo = st.file_uploader("Envie o arquivo.xlsx", type=["xlsx"])
    if arquivo:
        df_import = pd.read_excel(arquivo)
        for _, row in df_import.iterrows():
            nome = str(row['Nome']).strip().title()
            cargo = str(row['Cargo']).strip().title()
            folgas_str = str(row['Folgas_Semana']) if pd.notna(row['Folgas_Semana']) else ""
            folgas = [d.strip().title() for d in folgas_str.split(",") if d.strip()]
            st.session_state.funcionarios[nome] = {"cargo": cargo, "folgas_semana": folgas, "horas_mes": 0, "dias_trabalhados_seguidos": 0, "ultimo_dia_trabalhado": None, "ultimas_funcoes": []}
        st.success(f"{len(df_import)} funcionários importados!")

    st.header("2. Adicionar Manualmente")
    col1, col2, col3 = st.columns(3)
    with col1: nome = st.text_input("Nome")
    with col2: cargo = st.selectbox("Cargo", st.session_state.cargos)
    with col3: folgas = st.multiselect("Folga Fixa na Semana", DIAS_SEMANA)
    if st.button("Adicionar Funcionário Manual"):
        if nome: st.session_state.funcionarios[nome.title()] = {"cargo": cargo, "folgas_semana": folgas, "horas_mes": 0, "dias_trabalhados_seguidos": 0, "ultimo_dia_trabalhado": None, "ultimas_funcoes": []}
    if st.session_state.funcionarios: st.dataframe(pd.DataFrame.from_dict(st.session_state.funcionarios, orient='index'))

# 2. ABA MODELO + CARGOS + FERIADOS
with tab2:
    st.header("1. Adicionar Novo Cargo")
    col1, col2 = st.columns([3,1])
    with col1: novo_cargo = st.text_input("Nome do Novo Cargo")
    with col2:
        if st.button("Adicionar Cargo"):
            if novo_cargo and novo_cargo.title() not in st.session_state.cargos:
                st.session_state.cargos.append(novo_cargo.title())
                st.success(f"Cargo {novo_cargo.title()} adicionado!")
    st.write("Cargos Atuais:", ", ".join(st.session_state.cargos))

    st.divider()
    st.header("2. Marcar Feriados do Mês")
    col1, col2 = st.columns(2)
    with col1: mes_feriado = st.selectbox("Mês Feriado", range(1,13), format_func=lambda x: calendar.month_name[x], key="mes_f")
    with col2: ano_feriado = st.number_input("Ano Feriado", 2026, 2030, 2026, key="ano_f")
    num_dias = calendar.monthrange(ano_feriado, mes_feriado)[1]
    dias_mes = [datetime(ano_feriado, mes_feriado, d).strftime("%d/%m/%Y") for d in range(1, num_dias+1)]
    feriados_selecionados = st.multiselect("Selecione os dias de Feriado/Folga Geral", dias_mes, default=st.session_state.feriados)
    st.session_state.feriados = feriados_selecionados

    st.divider()
    st.header("3. Modelo de Escala por Dia da Semana")
    modelo_dia = st.session_state.modelo_dia.copy()
    for dia_semana in DIAS_SEMANA:
        with st.expander(f"**{dia_semana}**"):
            if dia_semana not in modelo_dia: modelo_dia[dia_semana] = {}
            cols = st.columns(len(st.session_state.cargos))
            for i, funcao in enumerate(st.session_state.cargos):
                with cols[i]:
                    qtd_default = modelo_dia[dia_semana].get(funcao, 0)
                    qtd = st.number_input(funcao, 0, 20, qtd_default, key=f"modelo_{dia_semana}_{funcao}")
                    if qtd > 0: modelo_dia[dia_semana][funcao] = qtd
                    elif funcao in modelo_dia[dia_semana]: del modelo_dia[dia_semana][funcao]
    st.session_state.modelo_dia = modelo_dia

# 3. LÓGICA DE GERAÇÃO
def gerar_escala_mensal(ano, mes, modelo_dia, funcionarios, feriados):
    num_dias = calendar.monthrange(ano, mes)[1]
    escala_dict = defaultdict(lambda: defaultdict(list))
    for dia in range(1, num_dias + 1):
        data_atual = datetime(ano, mes, dia)
        data_str = data_atual.strftime("%d/%m/%Y")
        dia_semana_str = DIAS_SEMANA[data_atual.weekday()]
        if data_str in feriados:
            escala_dict[dia]["FERIADO"] = ["LOJA FECHADA"]
            continue
        modelo_do_dia = modelo_dia.get(dia_semana_str, {})
        disponiveis = [n for n,d in funcionarios.items() if dia_semana_str not in d["folgas_semana"]]
        disponiveis = [n for n in disponiveis if not (funcionarios[n]["ultimo_dia_trabalhado"] == dia-1 and funcionarios[n]["dias_trabalhados_seguidos"] >= MAX_DIAS_SEGUIDOS)]
        indice_func = 0
        for funcao, qtd in modelo_do_dia.items():
            for i in range(qtd):
                candidatos_cargo = [f for f in disponiveis if funcionarios[f]["cargo"] == funcao]
                candidatos = candidatos_cargo if candidatos_cargo else disponiveis
                candidatos = sorted(candidatos, key=lambda x: funcionarios[x]["ultimas_funcoes"].count(funcao))
                if not candidatos: continue
                escolhido = candidatos[indice_func % len(candidatos)]
                escala_dict[dia][funcao].append(escolhido)
                funcionarios[escolhido]["horas_mes"] += HORAS_POR_TURNO
                funcionarios[escolhido]["ultimas_funcoes"].append(funcao)
                if funcionarios[escolhido]["ultimo_dia_trabalhado"] == dia - 1: funcionarios[escolhido]["dias_trabalhados_seguidos"] += 1
                else: funcionarios[escolhido]["dias_trabalhados_seguidos"] = 1
                funcionarios[escolhido]["ultimo_dia_trabalhado"] = dia
                indice_func += 1
    return escala_dict

# 4. VISUAL CALENDÁRIO
with tab3:
    st.header("Visualização em Calendário")
    mes = st.selectbox("Mês", range(1,13), format_func=lambda x: calendar.month_name[x], key="mes_cal")
    ano = st.number_input("Ano", 2026, 2030, 2026, key="ano_cal")
    if st.button("Gerar Calendário"):
        if not st.session_state.funcionarios or not st.session_state.modelo_dia: st.error("Cadastre funcionários e o modelo primeiro!")
        else:
            funcs_copia = {k: v.copy() for k,v in st.session_state.funcionarios.items()}
            for k in funcs_copia: funcs_copia[k].update({"horas_mes":0, "dias_trabalhados_seguidos":0, "ultimo_dia_trabalhado":None})
            st.session_state.escala_gerada = gerar_escala_mensal(ano, mes, st.session_state.modelo_dia, funcs_copia, st.session_state.feriados)
            st.session_state.funcs_resumo = funcs_copia
            st.session_state.mes_gerado = mes
            st.session_state.ano_gerado = ano
            cal = calendar.monthcalendar(ano, mes)
            st.subheader(f"{calendar.month_name[mes]} de {ano}")
            for semana in cal:
                cols = st.columns(7)
                for i, dia in enumerate(semana):
                    with cols[i]:
                        if dia!= 0:
                            data_str = datetime(ano, mes, dia).strftime("%d/%m/%Y")
                            if data_str in st.session_state.feriados: st.markdown(f"**{dia} - {DIAS_SEMANA_CURTO[i]}**"); st.error("FERIADO")
                            else:
                                st.markdown(f"**{dia} - {DIAS_SEMANA_CURTO[i]}**")
                                if dia in st.session_state.escala_gerada:
                                    for funcao, pessoas in st.session_state.escala_gerada[dia].items(): st.caption(f"{funcao}: {', '.join(pessoas)}")

# 5. PDF MENSAL + PDF DO DIA - CORRIGIDO V18.1
class PDF(FPDF):
    titulo_cabecalho = ""

    def header(self):
        self.set_fill_color(255, 111, 0); self.set_text_color(255, 255, 255); self.set_font('Arial', 'B', 16)
        self.cell(0, 12, self.titulo_cabecalho, 0, 1, 'C', 1); self.set_text_color(0, 0, 0); self.ln(3)

def criar_pdf_mensal(escala, ano, mes, feriados):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.titulo_cabecalho = f"{NOME_SUPERMERCADO} - ESCALA {calendar.month_name[mes].upper()} {ano}"
    pdf.add_page()

    largura_col = 40; altura_linha = 6; pdf.set_font("Arial", "B", 9); pdf.set_fill_color(255, 243, 224)
    for dia_sem in DIAS_SEMANA_CURTO: pdf.cell(largura_col, 8, dia_sem, 1, 0, "C", 1); pdf.ln()
    cal = calendar.monthcalendar(ano, mes); pdf.set_font("Arial", "", 7)
    for semana in cal:
        y_inicial = pdf.get_y(); alturas = [0]*7
        for i, dia in enumerate(semana):
            pdf.set_xy(10 + i*largura_col, y_inicial)
            if dia == 0: pdf.cell(largura_col, 50, "", 1); alturas[i] = 50
            else:
                data_str = datetime(ano, mes, dia).strftime("%d/%m/%Y"); txt = f"DIA {dia}\n"
                if data_str in feriados:
                    pdf.set_fill_color(255,0,0); pdf.set_text_color(255,255,255)
                    pdf.cell(largura_col, altura_linha, " FERIADO - FECHADO", 1, 1, "L", 1); pdf.set_text_color(0,0,0)
                elif dia in escala:
                    for funcao, pessoas in escala[dia].items(): txt += f"{funcao}: {', '.join(pessoas)}\n"
                pdf.multi_cell(largura_col, altura_linha, txt, 1); alturas[i] = pdf.get_y() - y_inicial
        max_altura = max(alturas); pdf.set_y(y_inicial + max_altura)
    pdf.set_y(-20); pdf.set_font('Arial', 'I', 8)
    pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y')} | Desenvolvido por: {NOME_CRIADOR} - {CARGO_CRIADOR} | Assinatura Gerencia: ____________________", 0, 0, 'C')
    return bytes(pdf.output()) # CORREÇÃO AQUI

def criar_pdf_dia(escala, ano, mes, dia):
    pdf = PDF(orientation='P', unit='mm', format='A4')
    data_str = datetime(ano, mes, dia).strftime("%d/%m/%Y"); dia_semana = DIAS_SEMANA[datetime(ano, mes, dia).weekday()]
    pdf.titulo_cabecalho = f"ESCALA DO DIA - {dia} {dia_semana.upper()}"
    pdf.add_page()

    pdf.set_font("Arial", "B", 12); pdf.cell(0, 10, f"{NOME_SUPERMERCADO}", 0, 1, "C"); pdf.ln(5)
    if dia in escala:
        for funcao, pessoas in escala[dia].items():
            pdf.set_font("Arial", "B", 11); pdf.cell(0, 8, f"{funcao}:", 0, 1); pdf.set_font("Arial", "", 10)
            for p in pessoas: pdf.cell(0, 6, f" - {p}", 0, 1); pdf.ln(2)
    else: pdf.set_font("Arial", "", 12); pdf.cell(0, 10, "Sem funcionarios escalados.", 0, 1)
    pdf.set_y(-20); pdf.set_font('Arial', 'I', 8)
    pdf.cell(0, 10, f"Desenvolvido por: {NOME_CRIADOR} - {CARGO_CRIADOR} | Assinatura Enc: ____________________", 0, 0, 'L')
    return bytes(pdf.output()) # CORREÇÃO AQUI

# 6. ABA GERAR E EXPORTAR
with tab4:
    st.header("Exportar para Imprimir")
    if 'escala_gerada' in st.session_state:
        mes = st.session_state.mes_gerado; ano = st.session_state.ano_gerado
        pdf_mensal = criar_pdf_mensal(st.session_state.escala_gerada, ano, mes, st.session_state.feriados)
        st.download_button("📄 Baixar PDF Mensal - Quadro de Avisos", pdf_mensal, f"escala_mensal_{mes}_{ano}.pdf", "application/pdf")
        st.divider(); st.header("Exportar Escala do Dia")
        dia_export = st.selectbox("Selecione o dia", range(1, calendar.monthrange(ano, mes)[1]+1))
        pdf_dia = criar_pdf_dia(st.session_state.escala_gerada, ano, mes, dia_export)
        st.download_button(f"📄 Baixar PDF do Dia {dia_export}", pdf_dia, f"escala_dia_{dia_export}_{mes}_{ano}.pdf", "application/pdf")
        st.divider(); st.header("Exportar Planilha Excel")
        dados_excel = []
        for dia, funcoes in st.session_state.escala_gerada.items():
            for funcao, pessoas in funcoes.items():
                for pessoa in pessoas: dados_excel.append({"Dia": dia, "Data": datetime(ano, mes, dia).strftime("%d/%m/%Y"), "Funcao": funcao, "Funcionario": pessoa})
        df_export = pd.DataFrame(dados_excel)
        st.download_button("📊 Baixar Excel da Escala", df_export.to_excel(index=False), f"escala_excel_{mes}_{ano}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.divider(); st.header("Resumo de Horas")
        df_horas = pd.DataFrame.from_dict(st.session_state.funcs_resumo, orient='index')[['cargo', 'horas_mes']]; st.dataframe(df_horas)
    else: st.warning("Gere o calendário primeiro na aba 3")

# 7. ABA: VER E TROCAR ESCALA + HISTÓRICO COM MOTIVO
with tab5:
    st.header("Consultar e Trocar Funcionários por Dia")
    if 'escala_gerada' in st.session_state:
        mes = st.session_state.mes_gerado; ano = st.session_state.ano_gerado
        dia_consulta = st.selectbox("Selecione o dia para ver/trocar", range(1, calendar.monthrange(ano, mes)[1]+1), key="dia_consulta")
        data_str = datetime(ano, mes, dia_consulta).strftime("%d/%m/%Y")
        st.subheader(f"Escala de {dia_consulta} de {calendar.month_name[mes]} - {data_str}")

        if data_str in st.session_state.feriados: st.error("FERIADO - LOJA FECHADA")
        elif dia_consulta in st.session_state.escala_gerada:
            for funcao, pessoas in st.session_state.escala_gerada[dia_consulta].items():
                st.write(f"### **{funcao}**")
                for idx, p in enumerate(pessoas):
                    col1, col2 = st.columns([4,2])
                    with col1: st.write(f"- {p}")
                    with col2:
                        dia_semana_str = DIAS_SEMANA[datetime(ano, mes, dia_consulta).weekday()]
                        candidatos_troca = [n for n,d in st.session_state.funcionarios.items()
                                            if d["cargo"] == funcao and dia_semana_str not in d["folgas_semana"] and n!= p]
                        novo = st.selectbox("Trocar por:", ["Manter"] + candidatos_troca, key=f"troca_{dia_consulta}_{funcao}_{idx}")
                        if novo!= "Manter":
                            with st.form(f"form_troca_{dia_consulta}_{funcao}_{idx}"):
                                st.write(f"**Trocar {p} por {novo}**")
                                motivo_select = st.selectbox("Motivo", MOTIVOS_PADRAO)
                                motivo_obs = ""
                                if motivo_select == "Outro":
                                    motivo_obs = st.text_input("Descreva o motivo")
                                motivo_final = motivo_obs if motivo_select == "Outro" else motivo_select
                                submitted = st.form_submit_button("Confirmar Troca")
                                if submitted:
                                    registro = {
                                        "Data Troca": datetime.now().strftime("%d/%m/%Y %H:%M"),
                                        "Data Escala": data_str,
                                        "Função": funcao,
                                        "Saiu": p,
                                        "Entrou": novo,
                                        "Motivo": motivo_final
                                    }
                                    st.session_state.historico_trocas.append(registro)
                                    st.session_state.escala_gerada[dia_consulta][funcao][idx] = novo
                                    st.success(f"Troca registrada!")
                                    st.rerun()

            st.divider()
            st.header("📝 Histórico de Trocas do Mês")
            if st.session_state.historico_trocas:
                df_hist = pd.DataFrame(st.session_state.historico_trocas)
                st.dataframe(df_hist, use_container_width=True)
                st.download_button("📊 Baixar Histórico em Excel", df_hist.to_excel(index=False), f"historico_trocas_{mes}_{ano}.xlsx")
                if st.button("Limpar Histórico"):
                    st.session_state.historico_trocas = []
                    st.rerun()
            else:
                st.info("Nenhuma troca realizada ainda neste mês.")
        else: st.info("Nenhum funcionário escalado para este dia.")
    else: st.warning("Gere o calendário primeiro na aba 3")

# 8. ABA SOBRE
with tab6:
    st.header("Sobre o Sistema")
    st.markdown(f"""
    ### {NOME_SUPERMERCADO}
    **Sistema de Gerenciamento de Escala Mensal**

    **Versão:** 18.1 - Corrigido para Web Cloud

    ---
    ### **Desenvolvimento**
    **Nome:** {NOME_CRIADOR}
    **Cargo:** {CARGO_CRIADOR}

    Sistema desenvolvido para otimizar a criação e gestão de escalas de trabalho do supermercado.

    **Funcionalidades:**
    - Geração automática de escala mensal
    - Salvar e Carregar dados
    - Marcação de feriados
    - Modelo por dia da semana
    - Troca rápida de funcionários
    - Histórico de trocas com motivo
    - Exportação em PDF e Excel

    ---
    <div class="footer">
    © 2026 {NOME_SUPERMERCADO} - Todos os direitos reservados
    </div>
    """, unsafe_allow_html=True)
