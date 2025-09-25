import os
import subprocess
import json
import re
from pathlib import Path
from datetime import datetime, timezone, timedelta
import xml.etree.ElementTree as ET
import pandas as pd
import traceback

# ===== Configurações =====
PST_TZ = timezone(timedelta(hours=-8))
SECTION_NAMES = {
    "specialize", "oobesystem", "setupcl", "windeploy",
    "setup", "oobeldr", "provisioning", "pasetup"
}

def parse_iso_utc(iso_str: str):
    if not iso_str:
        return None
    try:
        iso_norm = iso_str.replace("Z", "+00:00")
        return datetime.fromisoformat(iso_norm).astimezone(timezone.utc)
    except Exception:
        return None

def to_pst_str(dt_utc):
    if dt_utc is None:
        return None
    dt_pst = dt_utc.astimezone(PST_TZ)
    return dt_pst.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]

def get_child_case_insensitive(elem, name):
    target = name.lower()
    for child in list(elem):
        if (child.tag or "").split("}")[-1].lower() == target:
            return child
    return None

def get_text_or_attr(elem, key):
    child = get_child_case_insensitive(elem, key)
    if child is not None and child.text and child.text.strip():
        return child.text.strip()
    for k, v in elem.attrib.items():
        if k.lower() == key.lower() and v and str(v).strip():
            return str(v).strip()
    return None

def extract_section_record(machine_name, section_name, section_elem):
    start_raw = get_text_or_attr(section_elem, "StartTime")
    end_raw   = get_text_or_attr(section_elem, "EndTime")
    tick_raw  = get_text_or_attr(section_elem, "TickCount")

    start_utc = parse_iso_utc(start_raw)
    end_utc   = parse_iso_utc(end_raw)

    duration_s = duration_ms = None
    if start_utc and end_utc:
        delta = (end_utc - start_utc).total_seconds()
        duration_s = round(delta, 6)
        duration_ms = round(duration_s * 1000.0, 3)

    tick_ms = tick_s = None
    if tick_raw is not None:
        try:
            tick_ms = float(str(tick_raw).strip())
            tick_s = round(tick_ms / 1000.0, 6)
        except ValueError:
            pass

    return {
        "Machine": machine_name,
        "Section": section_name,
        "StartTime_PST": to_pst_str(start_utc),
        "EndTime_PST": to_pst_str(end_utc),
        "Duration_s": duration_s,
        "Duration_ms": duration_ms,
        "TickCount_ms": tick_ms,
        "TickCount_s": tick_s,
        "StartTime_UTC_raw": start_raw,
        "EndTime_UTC_raw": end_raw,
    }

def find_sections(root):
    found = {}
    for elem in root.iter():
        tag = (elem.tag or "").split("}")[-1]
        tag_lower = tag.lower()
        if tag_lower in SECTION_NAMES:
            found[tag_lower] = elem
    return found

def extract_telemetry_json_from_log(file_path: str):
    """Extrai o JSON do TelemetryData do arquivo de log WaSetup.xml"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Procura pelo padrão TelemetryData com JSON
        pattern = r'<TelemetryData[^>]*>({.*?})</TelemetryData>'
        match = re.search(pattern, content, re.DOTALL)
        
        if match:
            json_str = match.group(1)
            return json.loads(json_str)
        else:
            print(f"Aviso: Nenhum TelemetryData JSON encontrado em {file_path}")
            return None
            
    except Exception as e:
        print(f"Erro ao extrair JSON de {file_path}: {e}")
        return None

def parse_wasetup_xml(file_path: str, machine_name: str, pasetup_timing_ms: float = None):
    records = []
    try:
        # Tenta extrair dados do TelemetryData JSON primeiro
        telemetry_data = extract_telemetry_json_from_log(file_path)
        
        if telemetry_data:
            # Processa cada seção encontrada no JSON
            for section_key, section_data in telemetry_data.items():
                # Mapeia os nomes das seções do JSON para os nomes normalizados
                section_mapping = {
                    "specialize": "specialize",
                    "oobeSystem": "oobeSystem", 
                    "oobesystem": "oobeSystem",
                    "SetupCl": "SetupCl",
                    "setupcl": "SetupCl",
                    "WinDeploy": "WinDeploy",
                    "windeploy": "WinDeploy",
                    "Setup": "Setup",
                    "setup": "Setup",
                    "OobeLdr": "OobeLdr",
                    "oobeldr": "OobeLdr",
                    "provisioning": "provisioning",
                    "PaSetup": "PaSetup",
                    "pasetup": "PaSetup"
                }
                
                if section_key in section_mapping:
                    norm_name = section_mapping[section_key]
                    
                    # Extrai os dados da seção
                    start_raw = section_data.get("StartTime")
                    end_raw = section_data.get("EndTime")
                    tick_raw = section_data.get("TickCount")
                    
                    start_utc = parse_iso_utc(start_raw)
                    end_utc = parse_iso_utc(end_raw)

                    duration_s = duration_ms = None
                    if start_utc and end_utc:
                        delta = (end_utc - start_utc).total_seconds()
                        duration_s = round(delta, 6)
                        duration_ms = round(duration_s * 1000.0, 3)

                    tick_ms = tick_s = None
                    if tick_raw is not None:
                        try:
                            tick_ms = float(tick_raw)
                            tick_s = round(tick_ms / 1000.0, 6)
                        except (ValueError, TypeError):
                            pass

                    record = {
                        "Machine": machine_name,
                        "Section": norm_name,
                        "StartTime_PST": to_pst_str(start_utc),
                        "EndTime_PST": to_pst_str(end_utc),
                        "Duration_s": duration_s,
                        "Duration_ms": duration_ms,
                        "TickCount_ms": tick_ms,
                        "TickCount_s": tick_s,
                        "StartTime_UTC_raw": start_raw,
                        "EndTime_UTC_raw": end_raw,
                    }
                    records.append(record)
        
        # Se não conseguiu extrair dados do JSON, tenta o método XML original
        if not records:
            tree = ET.parse(file_path)
            root = tree.getroot()
            sections = find_sections(root)
            for section_lower, elem in sections.items():
                if section_lower == "oobesystem":
                    norm_name = "oobeSystem"
                elif section_lower == "setupcl":
                    norm_name = "SetupCl"
                elif section_lower == "windeploy":
                    norm_name = "WinDeploy"
                elif section_lower == "oobeldr":
                    norm_name = "OobeLdr"
                else:
                    norm_name = section_lower.capitalize()
                rec = extract_section_record(machine_name, norm_name, elem)
                records.append(rec)
        
        # Adiciona a seção PaSetup se o timing foi encontrado
        if pasetup_timing_ms is not None:
            pasetup_record = {
                "Machine": machine_name,
                "Section": "PaSetup",
                "StartTime_PST": None,
                "EndTime_PST": None,
                "Duration_s": round(pasetup_timing_ms / 1000.0, 6),
                "Duration_ms": round(pasetup_timing_ms, 3),
                "TickCount_ms": None,
                "TickCount_s": None,
                "StartTime_UTC_raw": None,
                "EndTime_UTC_raw": None,
            }
            records.append(pasetup_record)
        
        if not records:
            # Se ainda não tem dados, cria um registro de erro
            records.append({
                "Machine": machine_name,
                "Section": "NO_DATA",
                "StartTime_PST": None,
                "EndTime_PST": None,
                "Duration_s": None,
                "Duration_ms": None,
                "TickCount_ms": None,
                "TickCount_s": None,
                "StartTime_UTC_raw": None,
                "EndTime_UTC_raw": "No telemetry data found",
            })
            
    except Exception as e:
        records.append({
            "Machine": machine_name,
            "Section": "ERROR",
            "StartTime_PST": None,
            "EndTime_PST": None,
            "Duration_s": None,
            "Duration_ms": None,
            "TickCount_ms": None,
            "TickCount_s": None,
            "StartTime_UTC_raw": None,
            "EndTime_UTC_raw": f"Failed to parse: {e}",
        })
    return records

def extract_pasetup_timing(file_path: str):
    """Extrai o tempo do NonCVMInstall do arquivo PASetup.log"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Procura nas últimas linhas pelo padrão NonCVMInstall_total
        for line in reversed(lines[-10:]):  # Verifica as últimas 10 linhas
            if 'NonCVMInstall_total took' in line and 'ms' in line:
                # Extrai o número usando regex
                match = re.search(r'took\s+([\d.]+)\s+ms', line)
                if match:
                    return float(match.group(1))
        
        return None
    except Exception as e:
        print(f"Erro ao ler PASetup.log: {e}")
        return None

def discover_wasetup_files_os(base_dir: str):
    """Versão usando os.listdir em vez de pathlib para contornar problemas de rede"""
    try:
        machine_dirs = [d for d in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, d))]
        print(f"Encontradas {len(machine_dirs)} pastas de máquinas")
        
        for machine_name in machine_dirs:
            machine_path = os.path.join(base_dir, machine_name)
            try:
                # Procura pelas variações de "panther"
                panther_variants = ["panther", "panter", "Panther", "PANTHER", "Panter", "PANTER"]
                for variant in panther_variants:
                    panther_path = os.path.join(machine_path, variant)
                    if os.path.exists(panther_path) and os.path.isdir(panther_path):
                        wasetup_file = os.path.join(panther_path, "WaSetup.xml")
                        pasetup_file = os.path.join(panther_path, "PASetup.log")  # Corrigido para PASetup.log
                        if os.path.exists(wasetup_file):
                            # Verifica se PASetup.log existe
                            pasetup_timing = None
                            if os.path.exists(pasetup_file):
                                pasetup_timing = extract_pasetup_timing(pasetup_file)
                            yield wasetup_file, machine_name, pasetup_timing
                            break
            except Exception as e:
                print(f"Erro ao acessar máquina {machine_name}: {e}")
                continue
    except Exception as e:
        print(f"Erro ao listar diretório base {base_dir}: {e}")
        return

def build_machine_summary(all_records):
    """Constrói um resumo com uma máquina por linha e seções como colunas"""
    # Agrupa por máquina
    machines = {}
    
    for record in all_records:
        machine = record["Machine"]
        section = record["Section"]
        
        if machine not in machines:
            machines[machine] = {}
        
        # Pega o valor em milissegundos - prioriza Duration_ms, se não houver usa TickCount_ms
        duration_ms = record.get("Duration_ms")
        if duration_ms is None:
            # Se não tem Duration_ms, tenta usar TickCount_ms
            tick_count_ms = record.get("TickCount_ms")
            if tick_count_ms is not None:
                duration_ms = tick_count_ms
        
        if duration_ms is not None:
            machines[machine][section] = round(duration_ms, 3)
    
    # Cria lista de resultados
    summary_data = []
    
    # Define as colunas das seções na ordem desejada
    section_columns = ["specialize", "oobeSystem", "SetupCl", "Setup", "OobeLdr", "provisioning", "PaSetup"]
    
    for machine_name, sections in machines.items():
        row = {"Machine": machine_name}
        
        total_ms = 0
        for section in section_columns:
            value = sections.get(section)
            row[section] = value  # Remove o sufixo "_ms"
            if value is not None:
                total_ms += value
        
        row["Total"] = round(total_ms, 3) if total_ms > 0 else None  # Remove o sufixo "_ms"
        summary_data.append(row)
    
    return pd.DataFrame(summary_data)

def build_totals(df):
    """Mantido para compatibilidade - agora usa a nova função"""
    return build_machine_summary(df.to_dict('records'))

def main():
    # Permite caminho personalizado via argumento
    if len(os.sys.argv) > 1:
        base_path = os.sys.argv[1]
        print(f"Usando caminho personalizado: {base_path}")
    else:
        base_path = r"C:\PanterLogs"
    
    print(f"Procurando arquivos WaSetup.xml em: {base_path}")
    
    # Testa se o diretório existe usando os.path
    if not os.path.exists(base_path):
        print(f"❌ Diretório não acessível: {base_path}")
        print("\n💡 SOLUÇÕES:")
        print("1. Execute como administrador")
        print("2. Use o script map_network_drive.bat para mapear a unidade")
        print("3. Execute: python wasetup_report_os.py \"C:\\caminho\\local\\copiado\"")
        return

    all_records = []
    machine_data = {}  # Para organizar por máquina
    count_files = 0
    count_errors = 0
    pasetup_found = 0
    
    for file_path, machine, pasetup_timing in discover_wasetup_files_os(base_path):
        count_files += 1
        try:
            if pasetup_timing is not None:
                pasetup_found += 1
                print(f"Processando: {machine} -> WaSetup.xml + PASetup.log ({pasetup_timing} ms)")
            else:
                print(f"Processando: {machine} -> WaSetup.xml (sem PASetup.log)")
                
            recs = parse_wasetup_xml(file_path, machine, pasetup_timing)
            all_records.extend(recs)
            
            # Organiza por máquina para abas separadas
            machine_data[machine] = recs
            
        except Exception as e:
            count_errors += 1
            print(f"Erro ao processar {machine}: {e}")
            error_rec = {
                "Machine": machine,
                "Section": "PARSE_ERROR",
                "StartTime_PST": None,
                "EndTime_PST": None,
                "Duration_s": None,
                "Duration_ms": None,
                "TickCount_ms": None,
                "TickCount_s": None,
                "StartTime_UTC_raw": None,
                "EndTime_UTC_raw": f"Parse error: {e}",
            }
            all_records.append(error_rec)
            machine_data[machine] = [error_rec]

    print(f"Total de arquivos WaSetup.xml encontrados: {count_files}")
    print(f"Arquivos PASetup.log encontrados: {pasetup_found}")
    print(f"Erros de processamento: {count_errors}")
    
    if not all_records:
        print("Nenhum WaSetup.xml encontrado.")
        return

    # Cria o DataFrame principal (formato original detalhado)
    df_detailed = pd.DataFrame(all_records)
    df_detailed = df_detailed.sort_values(["Machine", "Section"], kind="stable").reset_index(drop=True)
    
    # Cria o DataFrame resumido (1 máquina por linha)
    df_summary = build_machine_summary(all_records)

    # Gera o relatório Excel com uma única aba
    out_xlsx = Path.cwd() / "wasetup_report.xlsx"
    
    # Verifica se o arquivo está aberto e tenta deletar se necessário
    try:
        if out_xlsx.exists():
            out_xlsx.unlink()  # Remove o arquivo existente
    except PermissionError:
        print(f"⚠️  Arquivo {out_xlsx} está aberto. Por favor, feche o Excel e execute novamente.")
        return
    
    try:
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
            # Aba única com resumo (1 máquina por linha)
            df_summary.to_excel(writer, sheet_name="WA_Setup_Report", index=False)
    except PermissionError:
        print(f"⚠️  Não foi possível salvar {out_xlsx}. Verifique se o arquivo está fechado.")
        return

    print(f"\n✅ Relatório gerado com sucesso: {out_xlsx}")
    print(f"Arquivos WaSetup.xml processados: {count_files - count_errors}")
    print(f"Arquivos PASetup.log encontrados: {pasetup_found}")
    print(f"Arquivos com erro: {count_errors}")
    print(f"Total de máquinas encontradas: {len(df_summary)}")
    print(f"Aba criada: WA_Setup_Report com resumo de todas as máquinas")

if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("Erro inesperado:")
        print(traceback.format_exc())