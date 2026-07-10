import os
import time
import sys
import pandas as pd
from tqdm import tqdm
import win32com.client as win32
from datetime import datetime, time as dt_time #inclusão do time as dt_time

def get_current_folder():
    """Retorna a pasta onde o script está sendo executado."""
    if getattr(sys, 'frozen', False):  # Para executáveis .exe
        return os.path.dirname(sys.executable)
    else:  # Para scripts .py
        return os.path.dirname(os.path.abspath(__file__))

# Configurações iniciais
pasta_origem = get_current_folder()  # Pasta atual do script
pasta_destino = os.path.join(pasta_origem, "Limpeza_Completa")
os.makedirs(pasta_destino, exist_ok=True)  # Cria a pasta de destino

# Caminhos dos logs
caminho_log_txt = os.path.join(pasta_destino, "_log_execucao.txt")
caminho_log_xlsx = os.path.join(pasta_destino, "_log_execucao_detalhado.xlsx")

# Registro de data/hora de início
inicio = datetime.now()
inicio_str = inicio.strftime("%d/%m/%Y %H:%M:%S")

# Log inicial
with open(caminho_log_txt, "w", encoding="utf-8") as log:
    log.write(f"=== PROCESSAMENTO INICIADO EM: {inicio_str} ===\n")
    log.write(f"PASTA DE ORIGEM: {pasta_origem}\n")
    log.write(f"PASTA DE DESTINO: {pasta_destino}\n\n")

log_detalhado = []

def detectar_cabecalho(ws, max_rows=20):
    """Detecta a linha de cabeçalho analisando as primeiras linhas."""
    for linha in range(1, max_rows + 1):
        valores = [str(ws.Cells(linha, col).Value or "").strip() 
                  for col in range(1, min(20, ws.UsedRange.Columns.Count + 1))]
        if sum(1 for v in valores if v and (v[0].isalpha() or "DATA" in v.upper())) >= 3:
            return linha
    return 1

def identificar_tipo_coluna(nome_coluna, valores):
    """Classifica o tipo de dados da coluna."""
    nome = str(nome_coluna).upper()

    # Nomes de colunas de hora fornecidos
    nomes_colunas_hora = [
        "ENTRADA1", "SAIDA1", "ENTRADA2", "SAIDA2", "ENTRADA3", "SAIDA3",
        "ENTRADA4", "SAIDA4", "ENTRADA5", "SAIDA5",
        "HE50_COMPENSAR", "HE70_COMPENSAR", "HE100_COMPENSAR", "ADN_COMPENSAR",
        "ATRASO_COMPENSAR", "HE50_PAGDESC", "HE70_PAGDESC", "HE100_PAGDESC",
        "ADN_PAGDESC", "ATRASO_PAGDESC", "JORNADA_HORAS", "JORNADA_TRABALHADA",
        "EXTRA_EXECUTADA"
    ] 

    # PRIORIZAÇÃO: Verifica se o nome da coluna está na lista de nomes de hora
    if nome in nomes_colunas_hora:
        return {"Tipo": "Hora", "Detalhe": "Nome de coluna pré-definido"}

    # Lógica existente para 'Identificação'
    if nome in ["CHAPA", "MATRÍCULA", "MATRICULA"]:
        return {"Tipo": "Identificação", "Detalhe": "Chapa do empregado"}
    
    # Lógica existente para 'Data'
    termos_data = ["DATA", "DT", "DATE", "VENCIMENTO", "VIGENCIA"]
    if any(termo in nome for termo in termos_data):
        tem_datas = any(isinstance(v, (datetime, float)) and (20000 < v < 50000 if isinstance(v, float) else True)
                       for v in valores if v is not None)
        return {"Tipo": "Data", "Detalhe": "Nominada" if tem_datas else "Nome sugere data"}
    
    # Análise de valores (simplificada)
    tipos = []
    for v in valores[:5]:
        if v is None: 
            continue
        if isinstance(v, datetime): 
            tipos.append("Data")
        elif isinstance(v, float) and 20000 < v < 50000: 
            tipos.append("Data Numérica")
        elif isinstance(v, (int, float)): 
            tipos.append("Numérico")
        elif isinstance(v, (dt_time, datetime)) and isinstance(v, dt_time): 
            tipos.append("Hora")
        elif isinstance(v, str) and (len(v) <= 8 and ':' in v):
            try:
                # Tenta converter para hora para verificar se é um valor válido
                datetime.strptime(v, '%H:%M:%S')
                tipos.append("Hora")
            except ValueError:
                tipos.append("Texto")
        else:
            tipos.append("Texto")
    
    if not tipos: 
        return {"Tipo": "Desconhecido", "Detalhe": "Sem valores"}

    # Prioriza a detecção de 'Hora' se ela estiver presente.
    if "Hora" in tipos:
        return {"Tipo": "Hora", "Detalhe": "Valores predominantes"}
    
    return {"Tipo": max(set(tipos), key=tipos.count), "Detalhe": "Valores predominantes"}

def registrar(msg):
    """Registra mensagens no log.txt e no console."""
    print(msg)
    with open(caminho_log_txt, "a", encoding="utf-8") as log:
        log.write(msg + "\n")


# Inicializa o Excel ORIGINAL
# excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel.Visible = False
# excel.DisplayAlerts = False

# Inicializa o Excel
try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
except Exception as e:
    registrar(f"ERRO ao iniciar o Excel: {str(e)}")
    sys.exit(1)

# Processamento dos arquivos
arquivos = [f for f in os.listdir(pasta_origem) if f.lower().endswith(".xlsx")]

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    caminho_arquivo_origem = os.path.join(pasta_origem, arquivo)
    caminho_arquivo_destino = os.path.join(pasta_destino, arquivo)
    registrar(f"\nProcessando: {arquivo}")

    try:
        wb = excel.Workbooks.Open(caminho_arquivo_origem)
        
        for ws in wb.Worksheets:
            linha_cabecalho = detectar_cabecalho(ws)
            registrar(f"  Aba: {ws.Name} | Cabeçalho na linha: {linha_cabecalho}")
                
            # Desfazendo todas as mesclagens da planilha
            try:
                ws.UsedRange.UnMerge()
            except Exception as e:
                registrar(f"ERRO ao desfazer mesclagens (toda planilha): {str(e)}")
                    

            
            for col in range(1, ws.UsedRange.Columns.Count + 1):
                nome_coluna = str(ws.Cells(linha_cabecalho, col).Value or "").strip()
                valores = [ws.Cells(row, col).Value for row in range(linha_cabecalho + 1, min(linha_cabecalho + 21, ws.UsedRange.Rows.Count + 1))]
                
                classificacao = identificar_tipo_coluna(nome_coluna, valores)
                
                # Limpa o formato da coluna antes de aplicar a nova formatação
                ws.Columns(col).ClearFormats()
                
                # Formatação de datas (exemplo simplificado)
                if classificacao["Tipo"] == "Data":
                    try:
                        formato = "dd/mm/aaaa hh:mm:ss" if any(
                            isinstance(v, datetime) and v.time() != datetime.min.time()
                            for v in valores[:5] if v is not None
                        ) else "dd/mm/aaaa"
                        ws.Columns(col).NumberFormat = formato
                        acao = f"Formatado ({formato})"
                    except Exception as e:
                        acao = f"Erro: {str(e)}"
                elif classificacao["Tipo"] == "Hora":
                    try:
                        ws.Columns(col).NumberFormat = "hh:mm:ss"
                        acao = "Formatado (hh:mm:ss)"
                    except Exception as e:
                        acao = f"Erro na formatação de hora: {str(e)}"

                else:
                    acao = "Ignorado"
                
                log_detalhado.append({
                    "Arquivo": arquivo,
                    "Aba": ws.Name,
                    "Coluna": nome_coluna,
                    "Tipo": classificacao["Tipo"],
                    "Ação": acao,
                    "Motivo": classificacao["Detalhe"]
                })

        wb.SaveAs(caminho_arquivo_destino)
        wb.Close(False)

    except Exception as e:
        registrar(f"ERRO no arquivo {arquivo}: {str(e)}")
        log_detalhado.append({
            "Arquivo": arquivo,
            "Erro": str(e)
        })

# Finalização
excel.Quit()
fim = datetime.now()
tempo_total = (fim - inicio).total_seconds()

# Adiciona metadados ao log detalhado
log_detalhado.insert(0, {
    "Arquivo": "METADADOS",
    "Aba": f"Início: {inicio_str}",
    "Coluna": f"Fim: {fim.strftime('%d/%m/%Y %H:%M:%S')}",
    "Tipo": f"Duração: {tempo_total:.2f} segundos",
    "Ação": f"Origem: {pasta_origem}",
    "Motivo": f"Destino: {pasta_destino}"
})

# Salva o log detalhado em Excel
df_log = pd.DataFrame(log_detalhado)
df_log.to_excel(caminho_log_xlsx, index=False)

# Mensagem final
registrar(f"\nProcessamento concluído em {tempo_total:.2f} segundos")
registrar(f"Log detalhado salvo em: {caminho_log_xlsx}")
