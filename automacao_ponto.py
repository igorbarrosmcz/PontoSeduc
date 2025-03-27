import os
import sys
import time
import tkinter as tk
from tkinter import simpledialog, messagebox
from datetime import datetime, time as dt_time, timedelta, date
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def verificar_chromedriver():
    """Verifica e configura o caminho do chromedriver"""
    # Caminho para o diretório de dados do aplicativo
    app_data_path = os.path.join(os.environ.get('APPDATA'), 'PontoSeduc')
    os.makedirs(app_data_path, exist_ok=True)
    
    chromedriver_path = os.path.join(app_data_path, 'chromedriver.exe')
    
    # Se não encontrou o chromedriver, tenta copiar do diretório de recursos
    if not os.path.exists(chromedriver_path):
        try:
            # Tenta encontrar o chromedriver no diretório de recursos (para PyInstaller)
            original_path = resource_path('chromedriver.exe')
            if os.path.exists(original_path):
                import shutil
                shutil.copy(original_path, chromedriver_path)
        except Exception as e:
            print(f"Erro ao copiar chromedriver: {e}")
            messagebox.showerror("Erro", "Chromedriver não encontrado. Por favor, instale manualmente.")
            return None
    
    return chromedriver_path

def formatar_horas_decimais(horas_decimais):
    """Converte horas decimais para formato HH:MM"""
    horas = int(horas_decimais)
    minutos = int(round((horas_decimais - horas) * 60))
    return f"{horas:02d}:{minutos:02d}"

def obter_dias_sem_expediente(ano, mes):
    """Exibe pop-up para usuário informar dias sem expediente"""
    root = tk.Tk()
    root.withdraw()
    
    # Configurações para manter a janela visível
    root.attributes('-topmost', True)
    root.lift()
    root.focus_force()
    
    mensagem = (
        "Informe os dias do mês (apenas números) sem expediente (feriados, recessos, etc.)\n"
        "Separados por vírgula. Exemplo: 3,4,5\n\n"
        f"Mês/Ano referência: {mes}/{ano}"
    )
    
    while True:
        input_dias = simpledialog.askstring("Dias sem expediente", mensagem, parent=root)
        if input_dias is None:
            root.destroy()
            return []
            
        try:
            dias = [int(dia.strip()) for dia in input_dias.split(",") if dia.strip()]
            ultimo_dia_mes = (datetime(ano, mes + 1, 1) - timedelta(days=1)).day
            for dia in dias:
                if not 1 <= dia <= ultimo_dia_mes:
                    raise ValueError(f"Dia {dia} inválido para o mês {mes}")
            root.destroy()
            return dias
        except ValueError as e:
            messagebox.showerror("Erro", f"Entrada inválida: {str(e)}\nPor favor, tente novamente.", parent=root)
            root.attributes('-topmost', True)
            root.lift()

def obter_feriados(ano, mes):
    """Retorna lista de feriados baseado no input do usuário"""
    dias_sem_expediente = obter_dias_sem_expediente(ano, mes)
    return [date(ano, mes, dia) for dia in dias_sem_expediente]

def calcular_dias_uteis_mes(ano, mes, feriados):
    """Calcula dias úteis no mês considerando feriados informados"""
    from calendar import monthrange
    dias_no_mes = monthrange(ano, mes)[1]
    return sum(
        1 for dia in range(1, dias_no_mes + 1)
        if datetime(ano, mes, dia).weekday() < 5 
        and date(ano, mes, dia) not in feriados
    )

def calcular_dias_uteis_ate_hoje(ano, mes, feriados):
    """Calcula dias úteis até a data atual, excluindo finais de semana e feriados"""
    hoje = date.today()
    dias_uteis = 0
    
    for dia in range(1, hoje.day + 1):
        data = date(ano, mes, dia)
        if data.weekday() < 5 and data not in feriados:
            dias_uteis += 1
    
    return dias_uteis

def calcular_dias_faltantes(horarios, dias_uteis_ate_hoje, feriados):
    """Identifica dias úteis sem registro de ponto"""
    dias_registrados = {h['data'] for h in horarios}
    dias_faltantes = []
    
    for dia in range(1, date.today().day + 1):
        data = date(horarios[0]['data'].year, horarios[0]['data'].month, dia)
        if (data.weekday() < 5 and 
            data not in feriados and 
            data not in dias_registrados and
            data < date.today()):
            dias_faltantes.append(data)
    
    return dias_faltantes

def calcular_horas_trabalhadas(horarios):
    """Calcula o total de horas trabalhadas"""
    total_horas = 0.0
    
    for dia in horarios:
        if dia['entrada_1'] and dia['saida_1']:
            periodo1 = (dia['saida_1'].hour * 60 + dia['saida_1'].minute) - \
                      (dia['entrada_1'].hour * 60 + dia['entrada_1'].minute)
            total_horas += periodo1 / 60
        
        if dia['entrada_2'] and dia['saida_2']:
            periodo2 = (dia['saida_2'].hour * 60 + dia['saida_2'].minute) - \
                      (dia['entrada_2'].hour * 60 + dia['entrada_2'].minute)
            total_horas += periodo2 / 60
    
    return total_horas

def calcular_saldo_diario(horario, jornada_diaria):
    """Calcula o saldo de horas para um dia específico"""
    horas_trabalhadas = 0
    
    if horario['entrada_1'] and horario['saida_1']:
        periodo1 = (horario['saida_1'].hour * 60 + horario['saida_1'].minute) - \
                  (horario['entrada_1'].hour * 60 + horario['entrada_1'].minute)
        horas_trabalhadas += periodo1 / 60
    
    if horario['entrada_2'] and horario['saida_2']:
        periodo2 = (horario['saida_2'].hour * 60 + horario['saida_2'].minute) - \
                  (horario['entrada_2'].hour * 60 + horario['entrada_2'].minute)
        horas_trabalhadas += periodo2 / 60
    
    return horas_trabalhadas - jornada_diaria

def determinar_jornada(horarios):
    """Determina se a jornada é de 6h ou 8h"""
    for dia in horarios:
        if dia.get('saida_2') is not None:
            return 8
    return 6

def verificar_dependencias():
    """Verifica se todas as dependências estão instaladas"""
    try:
        import selenium
        import tkinter
    except ImportError as e:
        messagebox.showerror("Erro", f"Dependência faltando: {e}\nInstale com: pip install selenium tkinter")
        return False
    return True

def main():
    if not verificar_dependencias():
        return
    
    caminho_chromedriver = verificar_chromedriver()
    if not caminho_chromedriver:
        return
    
    driver = None
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        
        service = Service(caminho_chromedriver)
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(30)

        try:
            driver.get("https://ponto.dev.educacao.al.gov.br/login")
            print("Digite seu CPF e senha no site e pressione ENTER no campo de senha para continuar...")
            
            campo_senha = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "senha"))
            )
            campo_senha.send_keys(Keys.RETURN)
            
            WebDriverWait(driver, 30).until(
                lambda d: d.current_url != "https://ponto.dev.educacao.al.gov.br/login"
            )
            print("Login realizado com sucesso!")
        except Exception as e:
            raise Exception(f"Erro durante o login: {str(e)}")

        try:
            time.sleep(5)
            print("Coletando horários...")
            
            linhas = WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((By.XPATH, "//table[@class='table table-bordered table-striped table-hover']/tbody/tr"))
            )
            
            horarios = []
            for linha in linhas:
                try:
                    colunas = linha.find_elements(By.TAG_NAME, "td")
                    if len(colunas) < 5:
                        continue
                        
                    data = datetime.strptime(colunas[0].text, "%d/%m/%Y").date()
                    entrada_1 = datetime.strptime(colunas[1].text, "%H:%M:%S").time() if colunas[1].text else None
                    saida_1 = datetime.strptime(colunas[2].text, "%H:%M:%S").time() if colunas[2].text else None
                    entrada_2 = datetime.strptime(colunas[3].text, "%H:%M:%S").time() if colunas[3].text else None
                    saida_2 = datetime.strptime(colunas[4].text, "%H:%M:%S").time() if colunas[4].text else None

                    if (entrada_1 and saida_1) or (entrada_2 and saida_2):
                        horarios.append({
                            "data": data,
                            "entrada_1": entrada_1,
                            "saida_1": saida_1,
                            "entrada_2": entrada_2,
                            "saida_2": saida_2
                        })
                except Exception as e:
                    print(f"Erro ao processar linha: {str(e)}")
                    continue

            if not horarios:
                print("Nenhum horário válido encontrado.")
                return

            carga_horaria_diaria = determinar_jornada(horarios)
            mes_atual = horarios[0]['data'].month
            ano_atual = horarios[0]['data'].year
            
            feriados = obter_feriados(ano_atual, mes_atual)
            dias_uteis_mes = calcular_dias_uteis_mes(ano_atual, mes_atual, feriados)
            dias_uteis_ate_hoje = calcular_dias_uteis_ate_hoje(ano_atual, mes_atual, feriados)
            
            dias_faltantes = calcular_dias_faltantes(horarios, dias_uteis_ate_hoje, feriados)
            total_horas = calcular_horas_trabalhadas(horarios)
            
            dias_com_ponto = len(horarios)
            dias_com_falta = len(dias_faltantes)
            total_dias_uteis_ate_hoje = dias_com_ponto + dias_com_falta

            carga_horaria_prevista_dias = total_dias_uteis_ate_hoje * carga_horaria_diaria
            saldo_horas = total_horas - carga_horaria_prevista_dias
            carga_horaria_prevista_mes = dias_uteis_mes * carga_horaria_diaria

            horas_previstas_format = formatar_horas_decimais(carga_horaria_prevista_dias)
            horas_trabalhadas_format = formatar_horas_decimais(total_horas)
            saldo_format = formatar_horas_decimais(abs(saldo_horas))
            tipo_saldo = "a favor" if saldo_horas >= 0 else "em débito"

            data_inicial = min(h['data'] for h in horarios)
            data_final = max(h['data'] for h in horarios)
            
            saldos_diarios = []
            for dia in horarios:
                saldo = calcular_saldo_diario(dia, carga_horaria_diaria)
                saldos_diarios.append({
                    'data': dia['data'],
                    'saldo': saldo
                })

            resumo = f"""
RESUMO DA JORNADA DE TRABALHO
──────────────────────────────
• Período analisado: {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}
• Tipo de Jornada: {carga_horaria_diaria}h/dia

DETALHAMENTO DOS DIAS
──────────────────────────────
• Dias úteis no período: {total_dias_uteis_ate_hoje}
  ✓ Dias trabalhados: {dias_com_ponto}
  ✗ Dias com falta: {dias_com_falta}

CÁLCULO DE HORAS
──────────────────────────────
• Carga horária prevista: {horas_previstas_format} ({carga_horaria_prevista_dias:.2f}h)
• Total de horas trabalhadas: {horas_trabalhadas_format} ({total_horas:.2f}h)
• Saldo geral: {saldo_format} {tipo_saldo} ({saldo_horas:+.2f}h)

PROJEÇÃO MENSAL
──────────────────────────────
• Dias úteis no mês: {dias_uteis_mes}
• Carga horária mensal prevista: {dias_uteis_mes * carga_horaria_diaria}h

REGISTROS DETALHADOS:
──────────────────────────────
"""

            for dia in horarios:
                e1 = dia['entrada_1'].strftime("%H:%M") if dia['entrada_1'] else "--:--"
                s1 = dia['saida_1'].strftime("%H:%M") if dia['saida_1'] else "--:--"
                e2 = dia['entrada_2'].strftime("%H:%M") if dia['entrada_2'] else "--:--"
                s2 = dia['saida_2'].strftime("%H:%M") if dia['saida_2'] else "--:--"
                
                saldo_dia = calcular_saldo_diario(dia, carga_horaria_diaria)
                saldo_format = formatar_horas_decimais(abs(saldo_dia))
                tipo_saldo = "excedente" if saldo_dia >= 0 else "faltante"
                
                resumo += f"{dia['data'].strftime('%d/%m/%Y')}: {e1} ▶ {s1} | {e2} ▶ {s2} | Saldo: {saldo_format} ({tipo_saldo})\n"

            resumo += "\nRESUMO DE SALDOS DIÁRIOS:\n"
            for saldo in saldos_diarios:
                saldo_format = formatar_horas_decimais(abs(saldo['saldo']))
                tipo_saldo = "excedente" if saldo['saldo'] >= 0 else "faltante"
                resumo += f"• {saldo['data'].strftime('%d/%m/%Y')}: {saldo_format} ({tipo_saldo})\n"

            if dias_faltantes:
                resumo += "\nDIAS FALTANTES:\n"
                resumo += "\n".join(f"→ {d.strftime('%d/%m/%Y')}" for d in dias_faltantes)

            desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
            arquivo = os.path.join(desktop, f"resumo_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.txt")
            
            with open(arquivo, 'w', encoding='utf-8') as f:
                f.write(resumo)
            
            os.system(f'notepad.exe "{arquivo}"')

        except Exception as e:
            print(f"Erro durante a coleta de dados: {str(e)}")
            raise

    except Exception as e:
        print(f"Ocorreu um erro durante a execução: {str(e)}")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except:
                pass

if __name__ == "__main__":
    main()