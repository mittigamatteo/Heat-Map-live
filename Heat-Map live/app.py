import streamlit as st
import pandas as pd
import time
import string
import os
import base64
import io
import numpy as np
from PIL import Image, ImageDraw
from streamlit_image_coordinates import streamlit_image_coordinates
import plotly.graph_objects as go

# Tentativo di importazione per Excel
try:
    import xlsxwriter
except ImportError:
    st.error("Modulo 'xlsxwriter' mancante. Installalo con: pip install xlsxwriter")

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Tactical Analysis Pro", layout="wide")

NOME_FILE_CAMPO = "istockphoto-962800488-612x612.jpg"
NOME_FILE_PORTA = "istockphoto-2020168572-612x612.jpg"

# --- CACHE RISORSE PESANTI ---
@st.cache_resource(show_spinner=False)
def load_base_image(path):
    if os.path.exists(path):
        return Image.open(path).convert("RGBA")
    return None

@st.cache_data(show_spinner=False)
def get_gridded_image(img_path, rows, cols, is_porta=False):
    img = load_base_image(img_path)
    if img is None: return None
    img_draw = img.copy()
    draw = ImageDraw.Draw(img_draw)
    w, h = img_draw.size
    cw, ch = w / cols, h / rows
    color = (255, 255, 0, 100) if is_porta else (255, 255, 255, 100)
    for r in range(rows):
        for c in range(cols):
            draw.rectangle([c*cw, r*ch, (c+1)*cw, (r+1)*ch], outline=color)
            lettera = string.ascii_uppercase[c] if c < 26 else f"Z{c}"
            prefix = "P-" if is_porta else ""
            draw.text((c*cw+5, r*ch+5), f"{prefix}{lettera}{r+1}", fill="white" if not is_porta else "yellow")
    return img_draw

# --- INIZIALIZZAZIONE SESSION STATE ---
if 'log_eventi' not in st.session_state: st.session_state.log_eventi = []
if 'log_possesso' not in st.session_state: st.session_state.log_possesso = [] 
if 'tipi_evento' not in st.session_state: 
    st.session_state.tipi_evento = ["⚽ Goal", "🎯 Tiro", "🚩 Corner", "❌ Fallo", "🟨 Cartellino"]
if 'periodo_attuale' not in st.session_state: st.session_state.periodo_attuale = "1° Tempo"

for k in ['timer_match_attivo', 'timer_noi_attivo', 'timer_avv_attivo', 
          'match_inizio', 'match_prec', 'noi_inizio', 'noi_prec', 'avv_inizio', 'avv_prec']:
    if k not in st.session_state: st.session_state[k] = 0 if 'inizio' in k or 'prec' in k else False

if 'zona_campo' not in st.session_state: st.session_state.zona_campo = "Non selezionata"
if 'zona_porta' not in st.session_state: st.session_state.zona_porta = "N/A"

# Inizializzazione per tracciamento cambio baricentro
if 'last_baricentro' not in st.session_state: st.session_state.last_baricentro = "📉 Basso (Difesa)"

# --- FUNZIONI UTILI ---
def get_time(inizio, prec, attivo):
    if attivo and inizio > 0:
        return prec + (time.time() - inizio)
    return prec

def fmt_time(s):
    return f"{int(s // 60):02d}:{int(s % 60):02d}"

def time_to_seconds(t_str):
    """Converte 'MM:SS' in secondi"""
    try:
        parts = str(t_str).split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except:
        return 0

def calcola_percentuale(t1, t2):
    totale = t1 + t2
    return (round((t1/totale)*100, 1), round((t2/totale)*100, 1)) if totale > 0 else (0, 0)

def get_cell_name(r, c, is_porta=False):
    lettera = string.ascii_uppercase[c] if c < 26 else f"Z{c}"
    return f"{'P-' if is_porta else ''}{lettera}{r+1}"

# --- HEATMAP ---
def generate_heatmap(data_df, img_path, max_r, max_c, title, coord_col):
    if data_df.empty: return None
    counts = np.zeros((max_r, max_c))
    for val in data_df[coord_col]:
        if not isinstance(val, str) or val in ["N/A", "Non selezionata"]: continue
        clean = val.replace("P-", "")
        try:
            l_part = "".join([c for c in clean if c.isalpha()])
            n_part = "".join([c for c in clean if c.isdigit()])
            c_idx = string.ascii_uppercase.index(l_part[0])
            r_idx = int(n_part) - 1
            if 0 <= r_idx < max_r and 0 <= c_idx < max_c: counts[r_idx, c_idx] += 1
        except: pass
    
    z_data = counts[::-1]
    z_display = [[v if v > 0 else None for v in row] for row in z_data]
    fig = go.Figure(go.Heatmap(z=z_display, colorscale=[[0,"rgba(0,255,0,0.4)"],[1,"rgba(255,0,0,0.7)"]], showscale=False))
    
    annotations = []
    for r in range(max_r):
        for c in range(max_c):
            val = counts[r, c]
            if val > 0:
                annotations.append(dict(x=c, y=max_r - 1 - r, text=str(int(val)), showarrow=False, font=dict(color="white", size=16, family="Arial Black")))

    img = load_base_image(img_path)
    if img:
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        img_b64 = base64.b64encode(buffered.getvalue()).decode()
        fig.update_layout(
            images=[dict(source=f"data:image/png;base64,{img_b64}", xref="x", yref="y", x=-0.5, y=max_r-0.5, sizex=max_c, sizey=max_r, sizing="stretch", layer="below")],
            xaxis=dict(visible=False, range=[-0.5, max_c-0.5]),
            yaxis=dict(visible=False, range=[-0.5, max_r-0.5], scaleanchor="x"),
            margin=dict(l=0, r=0, t=30, b=0), height=400, template="plotly_dark", title=title, annotations=annotations
        )
    return fig

# --- INTERFACCIA ---
st.title("🏟️ Analysis Pro: Precision View")

with st.sidebar:
    st.header("📐 Griglie")
    r_c, c_c = st.number_input("Righe Campo", 1, 26, 6), st.number_input("Col. Campo", 1, 26, 10)
    r_p, c_p = st.number_input("Righe Porta", 1, 15, 3), st.number_input("Col. Porta", 1, 15, 5)
    
    st.divider()
    if st.button("↩️ Rimuovi Ultimo Evento", use_container_width=True):
        if st.session_state.log_eventi: st.session_state.log_eventi.pop(); st.rerun()

# --- CAMPO E CONTROLLI ---
col_field, col_ctrl = st.columns([2, 1])

with col_field:
    img_c = get_gridded_image(NOME_FILE_CAMPO, r_c, c_c)
    if img_c:
        res_c = streamlit_image_coordinates(img_c, key="fc")
        if res_c:
            cw, ch = img_c.size[0]/c_c, img_c.size[1]/r_c
            st.session_state.zona_campo = get_cell_name(int(res_c['y']//ch), int(res_c['x']//cw))
    
    img_p = get_gridded_image(NOME_FILE_PORTA, r_p, c_p, True)
    if img_p:
        res_p = streamlit_image_coordinates(img_p, key="gc")
        if res_p:
            cwp, chp = img_p.size[0]/c_p, img_p.size[1]/r_p
            st.session_state.zona_porta = get_cell_name(int(res_p['y']//chp), int(res_p['x']//cwp), True)

with col_ctrl:
    # --- RIAVVIO AUTOMATICO TIMER CAMBIO TEMPO ---
    nuovo_periodo = st.select_slider("Tempo:", options=["1° Tempo", "2° Tempo", "Supplementari"], value=st.session_state.periodo_attuale)
    
    if nuovo_periodo != st.session_state.periodo_attuale:
        st.session_state.periodo_attuale = nuovo_periodo
        # Reset di tutti i timer al cambio tempo
        st.session_state.match_inizio = 0
        st.session_state.match_prec = 0
        st.session_state.timer_match_attivo = False
        
        st.session_state.noi_inizio = 0
        st.session_state.noi_prec = 0
        st.session_state.timer_noi_attivo = False
        
        st.session_state.avv_inizio = 0
        st.session_state.avv_prec = 0
        st.session_state.timer_avv_attivo = False
        st.rerun()

    @st.fragment(run_every="1s")
    def show_timers():
        tm = get_time(st.session_state.match_inizio, st.session_state.match_prec, st.session_state.timer_match_attivo)
        tn = get_time(st.session_state.noi_inizio, st.session_state.noi_prec, st.session_state.timer_noi_attivo)
        ta = get_time(st.session_state.avv_inizio, st.session_state.avv_prec, st.session_state.timer_avv_attivo)
        st.metric("⏱️ Tempo Match", fmt_time(tm))
        pn, pa = calcola_percentuale(tn, ta)
        st.write(f"🔵 NOI: {fmt_time(tn)} ({pn}%) | 🔴 AVV: {fmt_time(ta)} ({pa}%)")
    
    show_timers()
    
    # --- RESET CRONOMETRO GENERALE ---
    st.subheader("⏱️ Cronometro Match")
    bt1, bt2, bt3 = st.columns(3)
    if bt1.button("▶️ Avvio"):
        if not st.session_state.timer_match_attivo: st.session_state.match_inizio, st.session_state.timer_match_attivo = time.time(), True; st.rerun()
    if bt2.button("⏹️ Stop"):
        st.session_state.match_prec = get_time(st.session_state.match_inizio, st.session_state.match_prec, st.session_state.timer_match_attivo)
        st.session_state.timer_match_attivo = False
        
        # AGGIUNTO: Interrompe anche il possesso palla
        st.session_state.noi_prec = get_time(st.session_state.noi_inizio, st.session_state.noi_prec, st.session_state.timer_noi_attivo)
        st.session_state.avv_prec = get_time(st.session_state.avv_inizio, st.session_state.avv_prec, st.session_state.timer_avv_attivo)
        st.session_state.timer_noi_attivo = False
        st.session_state.timer_avv_attivo = False
        
        st.rerun()
    if bt3.button("🔄 Reset Match"):
        st.session_state.match_inizio = st.session_state.match_prec = 0; st.session_state.timer_match_attivo = False; st.rerun()

    st.divider()
    st.markdown(f"📍 **Campo:** `{st.session_state.zona_campo}` | 🥅 **Porta:** `{st.session_state.zona_porta}`")

    # --- CONTROLLI POSSESSO ---
    st.subheader("⚽ Possesso Palla")
    tn_static = get_time(st.session_state.noi_inizio, st.session_state.noi_prec, st.session_state.timer_noi_attivo)
    ta_static = get_time(st.session_state.avv_inizio, st.session_state.avv_prec, st.session_state.timer_avv_attivo)

    c1, c2, c3 = st.columns(3)
    if c1.button("🔵 NOI"):
        st.session_state.avv_prec = ta_static; st.session_state.timer_avv_attivo, st.session_state.timer_noi_attivo, st.session_state.noi_inizio = False, True, time.time(); st.rerun()
    if c2.button("⏸️ PAUSA"):
        st.session_state.noi_prec, st.session_state.avv_prec = tn_static, ta_static; st.session_state.timer_noi_attivo = st.session_state.timer_avv_attivo = False; st.rerun()
    if c3.button("🔴 AVV"):
        st.session_state.noi_prec = tn_static; st.session_state.timer_noi_attivo, st.session_state.timer_avv_attivo, st.session_state.avv_inizio = False, True, time.time(); st.rerun()

    st.divider()
    
    # --- LOGICA AUTOMATICA CAMBIO BARICENTRO ---
    def on_baricentro_change():
        """Salva automaticamente lo stato precedente quando si cambia baricentro"""
        tn_c = get_time(st.session_state.noi_inizio, st.session_state.noi_prec, st.session_state.timer_noi_attivo)
        ta_c = get_time(st.session_state.avv_inizio, st.session_state.avv_prec, st.session_state.timer_avv_attivo)
        tm_c = get_time(st.session_state.match_inizio, st.session_state.match_prec, st.session_state.timer_match_attivo)
        pn_c, pa_c = calcola_percentuale(tn_c, ta_c)
        
        st.session_state.log_possesso.append({
            "Periodo": st.session_state.periodo_attuale,
            "Minuto": fmt_time(tm_c),
            "Possesso NOI %": f"{pn_c}%",
            "Possesso AVV %": f"{pa_c}%",
            "Tempo NOI": fmt_time(tn_c),
            "Tempo AVV": fmt_time(ta_c),
            "Baricentro": st.session_state.last_baricentro
        })
        st.toast(f"🔄 Cambio Baricentro: Dati salvati per {st.session_state.last_baricentro}")
        st.session_state.last_baricentro = st.session_state.radio_baricentro

    tipo_possesso = st.radio(
        "Baricentro Possesso:", 
        ["📉 Basso (Difesa)", "📈 Alto (Attacco)"], 
        horizontal=True,
        key="radio_baricentro",
        on_change=on_baricentro_change
    )

    r1, r2 = st.columns(2)
    if r1.button("🔄 Reset Possesso", use_container_width=True):
        st.session_state.noi_inizio = st.session_state.noi_prec = st.session_state.avv_inizio = st.session_state.avv_prec = 0
        st.session_state.timer_noi_attivo = st.session_state.timer_avv_attivo = False; st.rerun()
    
    if r2.button("💾 Salva Possesso", use_container_width=True):
        tm_log = get_time(st.session_state.match_inizio, st.session_state.match_prec, st.session_state.timer_match_attivo)
        pn_log, pa_log = calcola_percentuale(tn_static, ta_static)
        st.session_state.log_possesso.append({
            "Periodo": st.session_state.periodo_attuale,
            "Minuto": fmt_time(tm_log),
            "Possesso NOI %": f"{pn_log}%",
            "Possesso AVV %": f"{pa_log}%",
            "Tempo NOI": fmt_time(tn_static),
            "Tempo AVV": fmt_time(ta_static),
            "Baricentro": tipo_possesso
        })
        st.toast(f"Possesso salvato: {tipo_possesso}")

    st.divider()
    # SEZIONE INPUT EVENTI
    soggetto = st.radio("Team", ["🔵 NOI", "🔴 AVVERSARIO"], horizontal=True)
    esito = st.selectbox("Esito (solo per tiri)", ["Parata", "Respinta", "Fuori", "Goal"])
    
    st.subheader("Registro Eventi")
    
    for ev in st.session_state.tipi_evento:
        if st.button(ev, use_container_width=True):
            tm_now = get_time(st.session_state.match_inizio, st.session_state.match_prec, st.session_state.timer_match_attivo)
            is_shot = "Tiro" in ev or "Goal" in ev
            
            if is_shot and esito != "Fuori":
                porta_val = st.session_state.zona_porta
            else:
                porta_val = "N/A" 

            evento_text = f"{ev} ({esito})" if is_shot else ev
            
            st.session_state.log_eventi.append({
                "Minuto": fmt_time(tm_now), 
                "Periodo": st.session_state.periodo_attuale, 
                "Team": soggetto,
                "Evento": evento_text, 
                "Posizione": st.session_state.zona_campo, 
                "Porta": porta_val 
            })
            st.toast(f"{ev} Registrato!")

# --- VISUALIZZAZIONE LIVE ---
st.divider()
tab_eventi, tab_possesso, tab_heat = st.tabs(["📊 Eventi Match", "⏱️ Log Possesso", "🔥 Heatmaps"])

with tab_eventi:
    if st.session_state.log_eventi: 
        st.dataframe(pd.DataFrame(st.session_state.log_eventi).iloc[::-1], use_container_width=True)

with tab_possesso:
    if st.session_state.log_possesso:
        st.dataframe(pd.DataFrame(st.session_state.log_possesso).iloc[::-1], use_container_width=True)

with tab_heat:
    if st.session_state.log_eventi:
        df_all = pd.DataFrame(st.session_state.log_eventi)
        
        # --- SEZIONE HEATMAPS TIRI ---
        st.subheader("🎯 Heatmaps Tiri & Goal")
        df_h = df_all[df_all['Evento'].str.contains("Tiro|Goal")]
        if not df_h.empty:
            for team in ["🔵 NOI", "🔴 AVVERSARIO"]:
                dft = df_h[df_h['Team'] == team]
                if not dft.empty:
                    c1, c2 = st.columns(2)
                    with c1: st.plotly_chart(generate_heatmap(dft, NOME_FILE_CAMPO, r_c, c_c, f"Campo {team}", "Posizione"), use_container_width=True, key=f"c_{team}")
                    with c2: st.plotly_chart(generate_heatmap(dft, NOME_FILE_PORTA, r_p, c_p, f"Porta {team}", "Porta"), use_container_width=True, key=f"p_{team}")
        else:
            st.info("Nessun tiro registrato.")
            
        st.divider()
        
        # --- SEZIONE HEATMAPS FALLI ---
        st.subheader("❌ Heatmaps Falli")
        df_f = df_all[df_all['Evento'].str.contains("Fallo")]
        if not df_f.empty:
            c1_f, c2_f = st.columns(2)
            teams = ["🔵 NOI", "🔴 AVVERSARIO"]
            for i, team in enumerate(teams):
                dft_f = df_f[df_f['Team'] == team]
                if not dft_f.empty:
                    with c1_f if i == 0 else c2_f:
                        st.plotly_chart(generate_heatmap(dft_f, NOME_FILE_CAMPO, r_c, c_c, f"Zone Falli {team}", "Posizione"), use_container_width=True, key=f"falli_c_{team}")
        else:
            st.info("Nessun fallo registrato.")

# --- ESPORTAZIONE REPORT ---
st.divider()
st.header("📥 Download Report")

def create_excel_heatmap_bytes(data_df, img_path, rows, cols, is_porta=False):
    """Genera heatmap con gradiente e NOMI ZONE per Excel"""
    base = load_base_image(img_path)
    if not base or data_df.empty: return None
    
    canvas = base.copy().convert("RGBA")
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    
    w, h = canvas.size
    cw, ch = w / cols, h / rows
    
    counts = np.zeros((rows, cols))
    coord_col = "Porta" if is_porta else "Posizione"
    
    for val in data_df[coord_col]:
        if not isinstance(val, str) or val in ["N/A", "Non selezionata"]: continue
        clean = val.replace("P-", "")
        try:
            l_part = "".join([c for c in clean if c.isalpha()])
            n_part = "".join([c for c in clean if c.isdigit()])
            c_idx = string.ascii_uppercase.index(l_part[0])
            r_idx = int(n_part) - 1
            if 0 <= r_idx < rows and 0 <= c_idx < cols: counts[r_idx, c_idx] += 1
        except: pass

    max_val = np.max(counts)

    for r in range(rows):
        for c in range(cols):
            lettera = string.ascii_uppercase[c] if c < 26 else f"Z{c}"
            prefix = "P-" if is_porta else ""
            zone_name = f"{prefix}{lettera}{r+1}"
            
            x1, y1 = c*cw, r*ch
            cx, cy = x1 + cw/2, y1 + ch/2
            
            val = counts[r, c]
            
            if val > 0:
                ratio = val / max_val if max_val > 0 else 0
                red = int(255 * ratio)
                green = int(255 * (1 - ratio))
                color = (red, green, 0, 180)
                
                draw.rectangle([x1, y1, x1+cw, y1+ch], fill=color, outline=(200,200,200,150))
                
                text_cnt = str(int(val))
                draw.text((cx-6, cy-6), text_cnt, fill="black")
                draw.text((cx-5, cy-5), text_cnt, fill="white")
                draw.text((x1+3, y1+3), zone_name, fill="white")
            else:
                draw.text((x1+4, y1+4), zone_name, fill="black") 
                draw.text((x1+3, y1+3), zone_name, fill="white")

    final_img = Image.alpha_composite(canvas, overlay)
    img_byte_arr = io.BytesIO()
    final_img.save(img_byte_arr, format='PNG')
    return img_byte_arr

def get_shot_stats(df, team, periodo=None):
    if df.empty: return 0, 0
    dft = df[df['Team'] == team]
    if periodo:
        dft = dft[dft['Periodo'] == periodo]
    dft = dft[dft['Evento'].str.contains("Tiro|Goal")]
    totale = len(dft)
    fuori = len(dft[dft['Evento'].str.contains("Fuori")])
    return totale, fuori

def get_foul_stats(df, team, periodo=None):
    """Calcola numero di Falli per team e periodo"""
    if df.empty: return 0
    dft = df[df['Team'] == team]
    if periodo:
        dft = dft[dft['Periodo'] == periodo]
    return len(dft[dft['Evento'].str.contains("Fallo")])

# Controllo e inizializzazione per evitare errori di array vuoti in Excel
df_eventi_export = pd.DataFrame(st.session_state.log_eventi)
if df_eventi_export.empty or 'Periodo' not in df_eventi_export.columns:
    df_eventi_export = pd.DataFrame(columns=["Minuto", "Periodo", "Team", "Evento", "Posizione", "Porta"])

df_possesso_export = pd.DataFrame(st.session_state.log_possesso)
if df_possesso_export.empty or 'Periodo' not in df_possesso_export.columns:
    df_possesso_export = pd.DataFrame(columns=["Periodo", "Minuto", "Possesso NOI %", "Possesso AVV %", "Tempo NOI", "Tempo AVV", "Baricentro"])

if 'Baricentro' not in df_possesso_export.columns:
    df_possesso_export['Baricentro'] = "N/D"
df_possesso_export['Baricentro'] = df_possesso_export['Baricentro'].fillna("N/D")

if len(st.session_state.log_eventi) > 0 or len(st.session_state.log_possesso) > 0:
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        sheet_name = 'Match Report'
        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet
        
        # STILI
        fmt_poss_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#FFD700', 'border': 1})
        fmt_avg_row = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#FFFACD', 'border': 1})
        fmt_noi_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#2E86C1', 'font_color': 'white', 'border': 1})
        fmt_avv_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#C0392B', 'font_color': 'white', 'border': 1})
        fmt_stats_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#85929E', 'font_color': 'white', 'border': 1})
        fmt_stats_cell = workbook.add_format({'align': 'center', 'border': 1})
        fmt_stats_total = workbook.add_format({'align': 'center', 'border': 1, 'bold': True, 'bg_color': '#FFFF00'}) # Evidenziatore Giallo
        fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})

        row_cursor = 0
        
        # --- 1. SEZIONE POSSESSO (RIEPILOGO PER PERIODO E GENERALE) ---
        if not df_possesso_export.empty and not df_possesso_export.dropna(how='all').empty:
            # Creiamo il riepilogo prendendo l'ultima riga salvata di ogni periodo
            df_summary = df_possesso_export.groupby('Periodo').tail(1).copy()

            # Pulizia colonne non necessarie per il riepilogo generale
            df_summary_clean = df_summary.drop(columns=['Baricentro', 'Minuto'], errors='ignore').reset_index(drop=True)

            # Scrittura Intestazioni
            worksheet.merge_range(row_cursor, 0, row_cursor, len(df_summary_clean.columns)-1, "STORICO POSSESSO PALLA (Riepilogo)", fmt_poss_header)
            row_cursor += 1
            for col_num, value in enumerate(df_summary_clean.columns.values):
                worksheet.write(row_cursor, col_num, value, fmt_poss_header)

            # Scrittura Dati (1° Tempo, 2° Tempo, ecc.)
            df_summary_clean.to_excel(writer, sheet_name=sheet_name, startrow=row_cursor+1, startcol=0, header=False, index=False)
            row_cursor += len(df_summary_clean) + 1 
            
            # --- CALCOLO POSSESSO GENERALE PARTITA ---
            tot_noi_sec = 0
            tot_avv_sec = 0
            
            # Somma totale di tutti i tempi
            for idx, row in df_summary_clean.iterrows():
                tot_noi_sec += time_to_seconds(row.get('Tempo NOI', '00:00'))
                tot_avv_sec += time_to_seconds(row.get('Tempo AVV', '00:00'))
            
            tot_match = tot_noi_sec + tot_avv_sec
            
            if tot_match > 0:
                p_noi_tot = (tot_noi_sec / tot_match) * 100
                p_avv_tot = (tot_avv_sec / tot_match) * 100
                
                worksheet.write(row_cursor, 0, "TOTALE MATCH", fmt_avg_row)
                worksheet.write(row_cursor, 1, f"{round(p_noi_tot, 1)}%", fmt_avg_row)
                worksheet.write(row_cursor, 2, f"{round(p_avv_tot, 1)}%", fmt_avg_row)
                worksheet.write(row_cursor, 3, fmt_time(tot_noi_sec), fmt_avg_row)
                worksheet.write(row_cursor, 4, fmt_time(tot_avv_sec), fmt_avg_row)
                row_cursor += 2
            else:
                row_cursor += 1

            # --- ANALISI TATTICA POSSESSO (Calcolo sui dati dettagliati) ---
            worksheet.write(row_cursor, 0, "ANALISI TATTICA POSSESSO (Alto vs Basso)", fmt_title)
            row_cursor += 2
            
            headers_tactical = ["TEAM", "SEC. BASSO", "SEC. ALTO", "% BASSO", "% ALTO"]
            for i, h in enumerate(headers_tactical):
                worksheet.write(row_cursor, i, h, fmt_stats_header)
            
            noi_low_sec = 0
            noi_high_sec = 0
            avv_low_sec = 0 
            avv_high_sec = 0 
            
            prev_noi = 0
            prev_avv = 0
            current_period_tracker = None
            
            # Calcolo dei delta temporali rispettando i cambi di periodo (i timer si azzerano)
            for index, row in df_possesso_export.iterrows():
                if current_period_tracker != row['Periodo']:
                    prev_noi = 0
                    prev_avv = 0
                    current_period_tracker = row['Periodo']
                
                curr_noi = time_to_seconds(row['Tempo NOI'])
                curr_avv = time_to_seconds(row['Tempo AVV'])
                
                delta_noi = curr_noi - prev_noi
                delta_avv = curr_avv - prev_avv
                bari = row['Baricentro']
                
                if delta_noi > 0:
                    if "Basso" in bari: noi_low_sec += delta_noi
                    elif "Alto" in bari: noi_high_sec += delta_noi
                if delta_avv > 0:
                    if "Basso" in bari: avv_low_sec += delta_avv
                    elif "Alto" in bari: avv_high_sec += delta_avv
                    
                prev_noi = curr_noi
                prev_avv = curr_avv
            
            tot_noi_calc = noi_low_sec + noi_high_sec
            p_noi_low = (noi_low_sec / tot_noi_calc * 100) if tot_noi_calc > 0 else 0
            p_noi_high = (noi_high_sec / tot_noi_calc * 100) if tot_noi_calc > 0 else 0

            tot_avv_calc = avv_low_sec + avv_high_sec
            p_avv_low = (avv_low_sec / tot_avv_calc * 100) if tot_avv_calc > 0 else 0
            p_avv_high = (avv_high_sec / tot_avv_calc * 100) if tot_avv_calc > 0 else 0

            row_cursor += 1
            worksheet.write(row_cursor, 0, "🔵 NOI", fmt_stats_cell)
            worksheet.write(row_cursor, 1, noi_low_sec, fmt_stats_cell)
            worksheet.write(row_cursor, 2, noi_high_sec, fmt_stats_cell)
            worksheet.write(row_cursor, 3, f"{round(p_noi_low, 1)}%", fmt_stats_cell)
            worksheet.write(row_cursor, 4, f"{round(p_noi_high, 1)}%", fmt_stats_cell)
            
            row_cursor += 1
            worksheet.write(row_cursor, 0, "🔴 AVV", fmt_stats_cell)
            worksheet.write(row_cursor, 1, avv_low_sec, fmt_stats_cell)
            worksheet.write(row_cursor, 2, avv_high_sec, fmt_stats_cell)
            worksheet.write(row_cursor, 3, f"{round(p_avv_low, 1)}%", fmt_stats_cell)
            worksheet.write(row_cursor, 4, f"{round(p_avv_high, 1)}%", fmt_stats_cell)
            
            row_cursor += 3
        else:
            row_cursor += 2

        # --- 2. SEZIONE EVENTI ---
        if not df_eventi_export.empty and len(df_eventi_export) > 0:
            df_noi = df_eventi_export[df_eventi_export['Team'] == "🔵 NOI"].drop(columns=['Team'])
            df_avv = df_eventi_export[df_eventi_export['Team'] == "🔴 AVVERSARIO"].drop(columns=['Team'])
            
            start_col_noi = 0
            start_col_avv = 7 
            
            if not df_noi.empty:
                worksheet.merge_range(row_cursor, start_col_noi, row_cursor, start_col_noi + len(df_noi.columns)-1, "EVENTI: NOI", fmt_noi_header)
                for col_num, value in enumerate(df_noi.columns.values):
                    worksheet.write(row_cursor+1, start_col_noi + col_num, value, fmt_noi_header)
                df_noi.to_excel(writer, sheet_name=sheet_name, startrow=row_cursor+2, startcol=start_col_noi, header=False, index=False)
            
            if not df_avv.empty:
                worksheet.merge_range(row_cursor, start_col_avv, row_cursor, start_col_avv + len(df_avv.columns)-1, "EVENTI: AVVERSARIO", fmt_avv_header)
                for col_num, value in enumerate(df_avv.columns.values):
                    worksheet.write(row_cursor+1, start_col_avv + col_num, value, fmt_avv_header)
                df_avv.to_excel(writer, sheet_name=sheet_name, startrow=row_cursor+2, startcol=start_col_avv, header=False, index=False)
            
            max_len = max(len(df_noi), len(df_avv))
            row_cursor += max_len + 4

        # --- 3. SEZIONE STATISTICHE MATCH ---
        if not df_eventi_export.empty:
            worksheet.write(row_cursor, 0, "STATISTICHE MATCH", fmt_title)
            row_cursor += 2
            
            headers_stats = ["METRICA", "NOI 1°T", "NOI 2°T", "NOI TOT", "AVV 1°T", "AVV 2°T", "AVV TOT"]
            for i, h in enumerate(headers_stats):
                worksheet.write(row_cursor, i, h, fmt_stats_header)
                
            noi_1_tot, noi_1_out = get_shot_stats(df_eventi_export, "🔵 NOI", "1° Tempo")
            noi_2_tot, noi_2_out = get_shot_stats(df_eventi_export, "🔵 NOI", "2° Tempo")
            noi_match_tot, noi_match_out = get_shot_stats(df_eventi_export, "🔵 NOI")
            avv_1_tot, avv_1_out = get_shot_stats(df_eventi_export, "🔴 AVVERSARIO", "1° Tempo")
            avv_2_tot, avv_2_out = get_shot_stats(df_eventi_export, "🔴 AVVERSARIO", "2° Tempo")
            avv_match_tot, avv_match_out = get_shot_stats(df_eventi_export, "🔴 AVVERSARIO")

            noi_1_fl = get_foul_stats(df_eventi_export, "🔵 NOI", "1° Tempo")
            noi_2_fl = get_foul_stats(df_eventi_export, "🔵 NOI", "2° Tempo")
            noi_tot_fl = get_foul_stats(df_eventi_export, "🔵 NOI")
            avv_1_fl = get_foul_stats(df_eventi_export, "🔴 AVVERSARIO", "1° Tempo")
            avv_2_fl = get_foul_stats(df_eventi_export, "🔴 AVVERSARIO", "2° Tempo")
            avv_tot_fl = get_foul_stats(df_eventi_export, "🔴 AVVERSARIO")
            
            row_cursor += 1
            worksheet.write(row_cursor, 0, "TIRI TOTALI", fmt_stats_cell)
            worksheet.write(row_cursor, 1, noi_1_tot, fmt_stats_cell)
            worksheet.write(row_cursor, 2, noi_2_tot, fmt_stats_cell)
            worksheet.write(row_cursor, 3, noi_match_tot, fmt_stats_total) # Highlight
            worksheet.write(row_cursor, 4, avv_1_tot, fmt_stats_cell)
            worksheet.write(row_cursor, 5, avv_2_tot, fmt_stats_cell)
            worksheet.write(row_cursor, 6, avv_match_tot, fmt_stats_total) # Highlight
            
            row_cursor += 1
            worksheet.write(row_cursor, 0, "FUORI SPECCHIO", fmt_stats_cell)
            worksheet.write(row_cursor, 1, noi_1_out, fmt_stats_cell)
            worksheet.write(row_cursor, 2, noi_2_out, fmt_stats_cell)
            worksheet.write(row_cursor, 3, noi_match_out, fmt_stats_total) # Highlight
            worksheet.write(row_cursor, 4, avv_1_out, fmt_stats_cell)
            worksheet.write(row_cursor, 5, avv_2_out, fmt_stats_cell)
            worksheet.write(row_cursor, 6, avv_match_out, fmt_stats_total) # Highlight

            row_cursor += 1
            worksheet.write(row_cursor, 0, "FALLI COMMESSI", fmt_stats_cell)
            worksheet.write(row_cursor, 1, noi_1_fl, fmt_stats_cell)
            worksheet.write(row_cursor, 2, noi_2_fl, fmt_stats_cell)
            worksheet.write(row_cursor, 3, noi_tot_fl, fmt_stats_total) # Highlight
            worksheet.write(row_cursor, 4, avv_1_fl, fmt_stats_cell)
            worksheet.write(row_cursor, 5, avv_2_fl, fmt_stats_cell)
            worksheet.write(row_cursor, 6, avv_tot_fl, fmt_stats_total) # Highlight
            
            row_cursor += 3

            # --- 4. SEZIONE HEATMAPS TIRI ---
            worksheet.write(row_cursor, 0, "HEATMAPS (Analisi Tiri - Goal)", fmt_title)
            row_cursor += 2
            
            periodi_presenti = df_eventi_export['Periodo'].unique()
            squadre = ["🔵 NOI", "🔴 AVVERSARIO"]
            
            for periodo in sorted(periodi_presenti):
                worksheet.write(row_cursor, 0, f"--- {periodo} ---", fmt_title)
                row_cursor += 1
                col_img_cursor = 0
                
                for team in squadre:
                    df_h = df_eventi_export[
                        (df_eventi_export['Periodo'] == periodo) & 
                        (df_eventi_export['Team'] == team) &
                        (df_eventi_export['Evento'].str.contains("Tiro|Goal"))
                    ]
                    
                    if not df_h.empty:
                        img_bytes = create_excel_heatmap_bytes(df_h, NOME_FILE_CAMPO, r_c, c_c, is_porta=False)
                        if img_bytes:
                            worksheet.write(row_cursor, col_img_cursor, f"Tiri Campo {team}")
                            worksheet.insert_image(row_cursor+1, col_img_cursor, "hm_c.png", {'image_data': img_bytes, 'x_scale': 1.0, 'y_scale': 1.0})
                            col_img_cursor += 8
                            
                        img_bytes_p = create_excel_heatmap_bytes(df_h, NOME_FILE_PORTA, r_p, c_p, is_porta=True)
                        if img_bytes_p:
                            worksheet.write(row_cursor, col_img_cursor, f"Tiri Porta {team}")
                            worksheet.insert_image(row_cursor+1, col_img_cursor, "hm_p.png", {'image_data': img_bytes_p, 'x_scale': 1.0, 'y_scale': 1.0})
                            col_img_cursor += 8
                
                row_cursor += 25

            # --- 5. SEZIONE HEATMAPS FALLI ---
            worksheet.write(row_cursor, 0, "HEATMAPS (Analisi Falli)", fmt_title)
            row_cursor += 2
            
            for periodo in sorted(periodi_presenti):
                worksheet.write(row_cursor, 0, f"--- {periodo} ---", fmt_title)
                row_cursor += 1
                col_img_cursor = 0
                
                for team in squadre:
                    df_f = df_eventi_export[
                        (df_eventi_export['Periodo'] == periodo) & 
                        (df_eventi_export['Team'] == team) &
                        (df_eventi_export['Evento'].str.contains("Fallo"))
                    ]
                    
                    if not df_f.empty:
                        img_bytes_f = create_excel_heatmap_bytes(df_f, NOME_FILE_CAMPO, r_c, c_c, is_porta=False)
                        if img_bytes_f:
                            worksheet.write(row_cursor, col_img_cursor, f"Falli {team}")
                            worksheet.insert_image(row_cursor+1, col_img_cursor, "hm_falli.png", {'image_data': img_bytes_f, 'x_scale': 1.0, 'y_scale': 1.0})
                            col_img_cursor += 8
                
                row_cursor += 25

    buffer.seek(0)
    
    st.download_button(
        label="📥 Scarica Report Completo",
        data=buffer,
        file_name="match_report_final.xlsx",
        mime="application/vnd.ms-excel"
    )
else:
    st.info("Nessun dato registrato da scaricare.")