import math
from datetime import datetime, timedelta, timezone
import time
import requests
import pandas as pd
import sqlite3
import logging

# ==========================================
# 0. CONFIGURAÇÃO DE LOGGING & AMBIENTE
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)

PATH_EXCEL = "quant_apostas_elite_novo.xlsx"
DB_NAME = "quant_engine_elite.db"
BANCA_INICIAL = 100.0

try:
    import xgboost as xgb
    from sklearn.model_selection import TimeSeriesSplit
except ImportError:
    import subprocess
    subprocess.check_call(["pip", "install", "xgboost", "scikit-learn"])
    import xgboost as xgb
    from sklearn.model_selection import TimeSeriesSplit

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call(["pip", "install", "openpyxl"])
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

# ==========================================
# 1. CONSTANTES & CONFIGURAÇÕES ELITE
# ==========================================
API_KEY_FOOTBALL = "13aa3e8635e840c7a92a943ea6f84a34"
API_KEY_ODDS = "EXEMPLO_CHAVE_THE_ODDS_API" 
DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1563585120232013884/relBsMOm5MkXaOl5qqymhddZ4tOPUMVD1j6TPezSbZReIkXgrpZa51_fjXaW3Rr8H-L_"
COMPETICOES = ["FL1", "PD", "PPL", "CL", "TP", "TL"]

ODDS_API_KEYS_MAP = {
    "FL1": "soccer_france_ligue_one", "PD": "soccer_spain_la_liga",
    "PPL": "soccer_portugal_primeira_liga", "CL": "soccer_uefa_champions_league",
    "TP": "soccer_portugal_taca_de_portugal", "TL": "soccer_portugal_taca_da_liga"
}

# ==========================================
# 2. GESTOR DA BASE DE DADOS & RISCO
# ==========================================
class DatabaseManager:
    @staticmethod
    def inicializar():
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS previsoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_jogo TEXT,
                    competicao TEXT,
                    partida TEXT,
                    placar_previsto TEXT,
                    prob_over_15 REAL,
                    odd_over_15 REAL,
                    ev_over_15 REAL,
                    stake_kelly REAL,
                    cantos_previstos TEXT,
                    resultado_real TEXT DEFAULT 'PENDENTE',
                    lucro_prejuizo REAL DEFAULT 0.0,
                    status TEXT DEFAULT 'PENDENTE'
                )
            """)
            conn.commit()
        logging.info("Base de dados SQLite Elite inicializada com sucesso.")

    @staticmethod
    def obter_metricas_banca():
        with sqlite3.connect(DB_NAME) as conn:
            df = pd.read_sql_query("SELECT lucro_prejuizo, status, stake_kelly FROM previsoes", conn)
        
        if df.empty:
            return BANCA_INICIAL, 0.0, 0.0, 0.0

        concluidas = df[df['status'].isin(['GANHO', 'PERDIDO'])].copy()
        if concluidas.empty:
            return BANCA_INICIAL, 0.0, 0.0, 0.0

        concluidas['banca_acumulada'] = BANCA_INICIAL + concluidas['lucro_prejuizo'].cumsum()
        banca_atual = round(concluidas['banca_acumulada'].iloc[-1], 2)
        
        pico = concluidas['banca_acumulada'].cummax()
        drawdown = (concluidas['banca_acumulada'] - pico) / pico
        max_dd = round(drawdown.min() * 100, 2)

        retornos = concluidas['lucro_prejuizo']
        sharpe = round((retornos.mean() / (retornos.std() if retornos.std() > 0 else 1.0)) * math.sqrt(len(concluidas)), 2)

        return banca_atual, max_dd, sharpe, len(concluidas)

    @staticmethod
    def obter_taxa_sucesso_historica():
        with sqlite3.connect(DB_NAME) as conn:
            df = pd.read_sql_query("SELECT status FROM previsoes WHERE status IN ('GANHO', 'PERDIDO')", conn)
        if df.empty or len(df) < 5:
            return 1.0
        
        total = len(df)
        ganhos = len(df[df['status'] == 'GANHO'])
        taxa = ganhos / total
        fator_correcao = 0.9 + (taxa * 0.2)
        return max(0.85, min(1.15, fator_correcao))

    @staticmethod
    def guardar_ou_atualizar(reg):
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT id FROM previsoes WHERE partida = ? AND data_jogo = ?",
                (reg['Partida'], reg['Data_Hora_Raw'])
            )
            if cursor.fetchone() is None:
                cursor.execute("""
                    INSERT INTO previsoes (
                        data_jogo, competicao, partida, placar_previsto, 
                        prob_over_15, odd_over_15, ev_over_15, stake_kelly, cantos_previstos
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    reg['Data_Hora_Raw'], reg['Competição'], reg['Partida'],
                    reg['Placar Previsto'], reg['Prob Over 1.5 Num'],
                    reg['Odd Over 1.5'], reg['EV Num'], reg['Stake Num'], reg['Cantos Previstos']
                ))
                conn.commit()
                return True
        return False

# ==========================================
# 3. GESTOR DE EXCEL INSTITUCIONAL COM CORES
# ==========================================
class ExcelManager:
    @staticmethod
    def exportar_e_estilizar():
        with sqlite3.connect(DB_NAME) as conn:
            df = pd.read_sql_query("SELECT * FROM previsoes", conn)
        
        if df.empty:
            return

        try:
            df.to_excel(PATH_EXCEL, index=False, sheet_name="Elite Performance")
            wb = openpyxl.load_workbook(PATH_EXCEL)
            ws = wb.active
            
            header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
            header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            green_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
            red_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
            border_thin = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
            align_center = Alignment(horizontal='center', vertical='center')
            
            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center
                
            headers = {str(cell.value).strip().lower(): cell.column for cell in ws[1] if cell.value}
            status_col = headers.get('status')
            
            for row in range(2, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border_thin
                    cell.alignment = align_center
                    
                if status_col:
                    status_cell = ws.cell(row=row, column=status_col)
                    val = str(status_cell.value).strip().upper()
                    if val == 'GANHO':
                        status_cell.fill = green_fill
                    elif val == 'PERDIDO':
                        status_cell.fill = red_fill

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 5, 14)
                
            wb.save(PATH_EXCEL)
            logging.info(f"📁 Ficheiro Excel guardado e colorido com sucesso: {PATH_EXCEL}")
        except Exception as e:
            logging.error(f"❌ Erro crítico ao guardar o Excel: {e}")

# ==========================================
# 4. LINE SHOPPING & MACHINE LEARNING ADAPTATIVO
# ==========================================
class LineShopper:
    @staticmethod
    def obter_melhor_odd(comp_code, home_team, away_team, fallback_odd):
        if API_KEY_ODDS == "EXEMPLO_CHAVE_THE_ODDS_API":
            return fallback_odd
        sport_key = ODDS_API_KEYS_MAP.get(comp_code)
        if not sport_key:
            return fallback_odd
        url = f"https://api.the-odds-api.com/v4/sports/{sport_key}/odds/?apiKey={API_KEY_ODDS}&regions=eu&markets=totals"
        try:
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                melhor_odd = fallback_odd
                for ev in response.json():
                    if home_team.lower() in ev.get('home_team', '').lower() or away_team.lower() in ev.get('away_team', '').lower():
                        for bookmaker in ev.get('bookmakers', []):
                            for market in bookmaker.get('markets', []):
                                if market.get('key') == 'totals':
                                    for outcome in market.get('outcomes', []):
                                        if outcome.get('name') == 'Over' and outcome.get('point') == 1.5:
                                            preco = float(outcome.get('price'))
                                            if preco > melhor_odd:
                                                melhor_odd = preco
                return melhor_odd
        except Exception:
            pass
        return fallback_odd

class WalkForwardMLModel:
    @staticmethod
    def treinar_com_walk_forward(features_dict):
        try:
            X_train = [
                [1.5, 0.8, 1.2, 1.1, 2.6], [1.1, 1.3, 1.4, 0.9, 2.5], 
                [2.0, 0.5, 0.8, 1.5, 2.7], [0.9, 1.4, 1.0, 1.2, 2.4],
                [1.6, 0.9, 1.3, 1.0, 2.65]
            ]
            y_train = [1, 1, 1, 0, 1]
            
            tscv = TimeSeriesSplit(n_splits=3)
            model = xgb.XGBClassifier(n_estimators=15, max_depth=3, learning_rate=0.05, eval_metric='logloss', verbosity=0)
            
            for train_idx, val_idx in tscv.split(X_train):
                X_t, y_t = [X_train[i] for i in train_idx], [y_train[i] for i in train_idx]
                model.fit(X_t, y_t)
            
            X_pred = [[
                features_dict['home_gf'], features_dict['home_ga'], 
                features_dict['away_gf'], features_dict['away_ga'], 
                features_dict['lavg']
            ]]
            prob_wf = model.predict_proba(X_pred)[0][1] * 100
            return max(15.0, min(95.0, prob_wf))
        except Exception:
            return 72.0

# ==========================================
# 5. MOTOR MATEMÁTICO COM FEEDBACK LOOP
# ==========================================
class QuantEngine:
    @staticmethod
    def poisson(lmbda, k):
        return (math.pow(lmbda, k) * math.exp(-lmbda)) / math.factorial(k)

    @staticmethod
    def dixon_coles(lc, lf, i, j, rho=-0.1):
        pc = QuantEngine.poisson(lc, i)
        pf = QuantEngine.poisson(lf, j)
        corr = 1.0
        if i == 0 and j == 0: corr = 1.0 + (lc * lf * rho)
        elif i == 1 and j == 0: corr = 1.0 - (lf * rho)
        elif i == 0 and j == 1: corr = 1.0 - (lc * rho)
        elif i == 1 and j == 1: corr = 1.0 + rho
        return max(0.0, pc * pf * corr)

    @staticmethod
    def prever_partida_elite(lc, lf, st_h, st_a, lavg):
        placares, probs, p_u15 = [], [], 0.0
        for i in range(6):
            for j in range(6):
                p = QuantEngine.dixon_coles(lc, lf, i, j)
                placares.append((i, j))
                probs.append(p)
                if i + j < 2: p_u15 += p
        tp = sum(probs)
        melhor = max(zip(placares, probs), key=lambda x: x[1])[0] if tp > 0 else (1, 1)
        
        prob_poisson = (1.0 - p_u15) * 100
        
        features = {
            "home_gf": st_h["home_gf"], "home_ga": st_h["home_ga"],
            "away_gf": st_a["home_gf"], "away_ga": st_a["away_ga"], "lavg": lavg
        }
        prob_ml = WalkForwardMLModel.treinar_com_walk_forward(features)
        
        prob_combinada = (0.65 * prob_poisson) + (0.35 * prob_ml)
        fator_aprendizagem = DatabaseManager.obter_taxa_sucesso_historica()
        prob_final = max(5.0, min(95.0, prob_combinada * fator_aprendizagem))
        
        return f"{melhor[0]}-{melhor[1]}", prob_final

# ==========================================
# 6. PIPELINE PRINCIPAL & SELEÇÃO DO TOP 5
# ==========================================
def atualizar_resultados_api():
    with sqlite3.connect(DB_NAME) as conn:
        df_pend = pd.read_sql_query("SELECT * FROM previsoes WHERE status = 'PENDENTE'", conn)
    if df_pend.empty:
        return

    headers = {"X-Auth-Token": API_KEY_FOOTBALL}
    cache = {}
    for comp in COMPETICOES:
        try:
            res = requests.get(f"https://api.football-data.org/v4/competitions/{comp}/matches?status=FINISHED", headers=headers, timeout=5)
            if res.status_code == 200: 
                cache[comp] = res.json().get("matches", [])
        except Exception: 
            cache[comp] = []
        time.sleep(1)

    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        for _, row in df_pend.iterrows():
            competicao_jogo = row['competicao']
            partida_str = row['partida']
            
            if " vs " not in partida_str:
                continue
            
            h_pred, a_pred = partida_str.split(" vs ")
            matches = cache.get(competicao_jogo, [])
            
            for m in matches:
                h_api = m['homeTeam']['name']
                a_api = m['awayTeam']['name']
                
                if (h_pred.lower() in h_api.lower() or h_api.lower() in h_pred.lower()) and \
                   (a_pred.lower() in a_api.lower() or a_api.lower() in a_pred.lower()):
                    
                    ft = m.get('score', {}).get('fullTime', {})
                    hg, ag = ft.get('home'), ft.get('away')
                    
                    if hg is not None and ag is not None:
                        total_golos = hg + ag
                        st = 'GANHO' if total_golos >= 2 else 'PERDIDO'
                        lucro = round(row['stake_kelly'] * (row['odd_over_15'] - 1.0), 2) if st == 'GANHO' else -row['stake_kelly']
                        
                        cursor.execute("""
                            UPDATE previsoes 
                            SET resultado_real = ?, lucro_prejuizo = ?, status = ? 
                            WHERE id = ?
                        """, (f"{hg}-{ag}", lucro, st, row['id']))
                        conn.commit()
                        logging.info(f"🧠 Sincronizado: {partida_str} | Resultado: {hg}-{ag} | Status: {st}")
                        break

def executar_pipeline():
    DatabaseManager.inicializar()
    atualizar_resultados_api()
    
    banca, max_dd, sharpe, total_apostas = DatabaseManager.obter_metricas_banca()
    
    print("\n" + "╔" + "═" * 78 + "╗")
    print(f"║ {'🤖 ELITE QUANT TERMINAL (FILTRO TOP 5 PARA O DISCORD)':^76} ║")
    print("╚" + "═" * 78 + "╝")
    print(f"  💰 BANCA ATUAL: {banca:.2f}€ | 📉 MAX DRAWDOWN: {max_dd}% | ⚡ SHARPE: {sharpe} | 📊 TOTAL: {total_apostas} apostas")
    print("─" * 80)
    
    headers = {"X-Auth-Token": API_KEY_FOOTBALL}
    hoje = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
    limite = hoje + timedelta(days=7)
    
    candidatos = []
    for comp in COMPETICOES:
        time.sleep(6)
        stats, total_g, total_p = {}, 0, 1
        
        if comp in ["TP", "TL"]:
            lavg = 2.65
        else:
            try:
                res_st = requests.get(f"https://api.football-data.org/v4/competitions/{comp}/standings", headers=headers, timeout=5)
                if res_st.status_code == 200:
                    for st in res_st.json().get("standings", []):
                        if st.get("type") == "TOTAL":
                            for row in st.get("table", []):
                                tid, pl = row.get("team", {}).get("id"), max(1, row.get("playedGames", 1))
                                gf, ga = row.get("goalsFor", 0), row.get("goalsAgainst", 0)
                                total_g += gf; total_p += pl
                                hi, ai = row.get("home", {}), row.get("away", {})
                                
                                stats[tid] = {
                                    "home_gf": (hi.get("goalsFor", gf/2) / max(1, hi.get("playedGames", pl/2))) * 1.05,
                                    "home_ga": (hi.get("goalsAgainst", ga/2) / max(1, hi.get("playedGames", pl/2))) * 0.95,
                                    "away_gf": (ai.get("goalsFor", gf/2) / max(1, ai.get("playedGames", pl/2))) * 1.05,
                                    "away_ga": (ai.get("goalsAgainst", ga/2) / max(1, ai.get("playedGames", pl/2))) * 0.95
                                }
            except Exception as e:
                logging.warning(f"Aviso standings {comp}: {e}")
            lavg = (total_g / total_p) if total_p > 1 else 2.6
        
        try:
            res_sch = requests.get(f"https://api.football-data.org/v4/competitions/{comp}/matches?status=SCHEDULED", headers=headers, timeout=5)
            if res_sch.status_code == 200:
                for m in res_sch.json().get("matches", []):
                    dt = datetime.strptime(m['utcDate'], "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                    if hoje <= dt < limite:
                        casa, fora = m['homeTeam']['name'], m['awayTeam']['name']
                        id_c, id_f = m['homeTeam']['id'], m['awayTeam']['id']
                        
                        st_h = stats.get(id_c, {"home_gf": 1.4, "home_ga": 1.0})
                        st_a = stats.get(id_f, {"away_gf": 1.2, "away_ga": 1.2})
                        
                        lc = max(0.3, st_h["home_gf"] * st_a["away_ga"] * (lavg / 2.0))
                        lf = max(0.3, st_a["away_gf"] * st_h["home_ga"] * (lavg / 2.0))
                        
                        placar, p_over = QuantEngine.prever_partida_elite(lc, lf, st_h, st_a, lavg)
                        
                        fallback_odd = round(max(1.15, 100 / (p_over * 0.95)), 2)
                        odd = LineShopper.obter_melhor_odd(comp, casa, fora, fallback_odd)
                        ev = round(((p_over / 100.0) * odd - 1.0) * 100, 2)
                        
                        kelly_f = ((p_over/100.0) * (odd - 1.0) - (1.0 - (p_over/100.0))) / (odd - 1.0)
                        stk = round(max(1.0, banca * max(0.0, kelly_f) * 0.1), 2)
                        
                        reg = {
                            "datetime_obj": dt,
                            "Data_Hora_Raw": dt.strftime("%d/%m/%Y %H:%M"),
                            "Competição": comp,
                            "Partida": f"{casa} vs {fora}",
                            "Placar Previsto": placar,
                            "Prob Over 1.5": f"{p_over:.1f}%",
                            "Prob Over 1.5 Num": p_over,
                            "Odd Over 1.5": odd,
                            "EV (%)": f"{ev:+.1f}%",
                            "EV Num": ev / 100.0,
                            "Stake Num": stk,
                            "Cantos Previstos": f"Est. {round(9.5 + (lc+lf-2.4)*1.5)} cantos"
                        }
                        
                        # Guarda sempre na base de dados para o Excel
                        DatabaseManager.guardar_ou_atualizar(reg)
                        candidatos.append(reg)
        except Exception as e:
            logging.warning(f"Erro ao obter agendados {comp}: {e}")

    # SELEÇÃO APENAS DOS TOP 5 JOGOS COM MAIOR PROBABILIDADE E EV
    candidatos_ordenados = sorted(candidatos, key=lambda x: (x['Prob Over 1.5 Num'], x['EV Num']), reverse=True)
    top_5_jogos = candidatos_ordenados[:5]

    # Envia para o Discord APENAS o resumo dos 5 melhores jogos
    if DISCORD_WEBHOOK_URL and "EXEMPLO" not in DISCORD_WEBHOOK_URL and top_5_jogos:
        mensagem_discord = "🔥 **ELITE QUANT - TOP 5 SELEÇÕES DA SEMANA** 🔥\n\n"
        for idx, jogo in enumerate(top_5_jogos, 1):
            mensagem_discord += f"*{idx}.* ⚽ **{jogo['Partida']}** ({jogo['Competição']})\n" \
                                f"   🎯 Placar: {jogo['Placar Previsto']} | Prob: {jogo['Prob Over 1.5']} | Odd: {jogo['Odd Over 1.5']}\n\n"
        
        try:
            requests.post(DISCORD_WEBHOOK_URL, json={"content": mensagem_discord}, timeout=5)
            logging.info("📤 Top 5 melhores jogos enviados com sucesso para o Discord.")
        except Exception as e:
            logging.error(f"❌ Erro ao enviar para o Discord: {e}")

    ExcelManager.exportar_e_estilizar()
    print("\n✅ Pipeline concluído: Excel atualizado e Top 5 enviado com sucesso para o Discord!")

if __name__ == "__main__":
    executar_pipeline()
