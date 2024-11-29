import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import StringVar
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import json
from datetime import datetime, timedelta, time
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Color, Alignment, PatternFill, Border, Side
from PIL import Image, ImageTk
import math
import re

dados_selecionados = {}

def salvar_dataframe(df, nome):
    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel file", "*.xlsx")],
        initialfile=f'{nome}.xlsx'
    )
    if output_path:
        df.to_excel(output_path, index=False)
        exibir_log(f"Arquivo {nome} salvo com Sucesso")

def aplicar_estilo(output_path):
    # Carregar a planilha salva
    wb = load_workbook(output_path)
    ws = wb.active

    # Ajuste da largura das colunas
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Pega a letra da coluna
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Preenchimento de fundo
    fill_grey = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    # Estilo de alinhamento e preenchimento das células
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):  # Começa na segunda linha
        for cell in row:
            if cell.row % 2 == 0:
                cell.fill = fill_grey  # Linhas pares recebem fundo cinza
            else:
                cell.fill = fill_white  # Linhas ímpares recebem fundo branco
            cell.alignment = Alignment(horizontal="left", vertical="center")  # Alinhamento das células

    # Aplicar bordas a todas as células
    apply_borders(ws)

    # Salvar a planilha com as formatações
    wb.save(output_path)

def abrir_popup_selecao_coleta():
    popup = ttk.Toplevel(root)
    popup.title("Selecione as funções para coletar")
    popup.geometry("400x300")
    
    frame_popup = ttk.Frame(popup)
    frame_popup.pack(expand=True)

    # Variáveis para armazenar as seleções
    selecoes = {
        "Coletar Plan. Cadastro": [tk.BooleanVar(), coleta_empresa],
        "Coletar Pessoas Para Alt.": [tk.BooleanVar(), alteracao_pessoas],
        "Planilha de Marcação": [tk.BooleanVar(), coleta_planilha_marcacoes],
        "Planilha de Inconsistência": [tk.BooleanVar(), coleta_planilha_marcacoes_inconsistencia],
        "Planilha de Incomum": [tk.BooleanVar(), coleta_planilha_marcacoes_incomum],
        "Planilha de Horas Faltas": [tk.BooleanVar(), coleta_planilha_marcacoes_faltantes],
        "Planilha de Ferias": [tk.BooleanVar(), coleta_planilha_ferias]
    }

    for i, (nome, (var, _)) in enumerate(selecoes.items()):
        ttk.Checkbutton(frame_popup, text=nome, variable=var).grid(row=i, column=0, sticky="w", padx=10, pady=5)

    # Centraliza o botão abaixo dos Checkbuttons
    ttk.Button(
        frame_popup, text="Iniciar Coleta", 
        command=lambda: iniciar_coleta(selecoes, popup), bootstyle="success").grid(row=len(selecoes), column=0, pady=10)
    
    popup.protocol("WM_DELETE_WINDOW", popup.destroy)
    
    frame_popup.grid_rowconfigure(len(selecoes), weight=1)
    frame_popup.grid_columnconfigure(0, weight=1)

def envio_justificativa():
    """Função para enviar marcações para a API usando o Excel selecionado."""
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de marcações",
        filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
    )
    if not caminho_arquivo:
        return

    try:
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return
    
    
    url = 'https://www.dimepkairos.com.br/RestServiceApi/PreJustificationRequest/PreJustificationRequest'
    
    resultados = []

    for index, row in df.iterrows():
        idjust = row['Justificativa']
        idfunc = row['Id Funcionário']
        # Converte Timestamp para string no formato YYYY-MM-DD
        data = row['Data'].strftime('%Y-%m-%d')
        # Converte time para string no formato HH:MM
        horas = row['Hora'].strftime('%H:%M')

        headers = {
            "identifier": dados_selecionados["CNPJ"],
            "key": dados_selecionados["Chave API"],
            'User-Agent': 'PostmanRuntime/7.30.0'
        }

        payload = {
            "IdJustification": idjust,
            "IdUser": "1",
            "IdEmployee": idfunc,
            "QtdHours": horas,
            "Date": data,
            "Notes": "Enviado via API",
            "RequestType": "1",
            "ResponseType": "AS400V1"
        }

        response = requests.post(url,json=payload, headers=headers
        )
        
        try:
            response_json = response.json()
            status = "Sucesso" if response_json.get("Sucesso") else "Falha"
            mensagem = response_json.get("Mensagem", "Erro desconhecido")
        except ValueError:
            status = "Falha"
            mensagem = "Resposta inválida da API"

        resultados.append([idfunc, data, status, mensagem])
        exibir_log(f"Matricula: {idfunc} | Status: {status} | Mensagem: {mensagem}")

    # Exibe mensagem final
    messagebox.showinfo("Processo Concluído", "Envio de marcações concluído!")

def abrir_popup_selecao_pessoas():
    popup = ttk.Toplevel(root)
    popup.title("Selecione as funções para coletar")
    popup.geometry("300x200")
    
    frame_popup = ttk.Frame(popup)
    frame_popup.pack(expand=True)

    # Variáveis para armazenar as seleções
    selecoes = {
        "Cadastrar Pessoas": [tk.BooleanVar(), cadastrar_pessoas],
        "Alteração de Pessoas": [tk.BooleanVar(), alteracao_pessoas_envio],
        "Envio de Ferias": [tk.BooleanVar(), envio_planilha_ferias]
    }

    # Criação dinâmica dos Checkbuttons
    for i, (nome, (var, _)) in enumerate(selecoes.items()):
        ttk.Checkbutton(frame_popup, text=nome, variable=var).grid(row=i, column=0, sticky="w", padx=10, pady=5)

    # Centraliza o botão abaixo dos Checkbuttons
    ttk.Button(
        frame_popup, text="Iniciar Coleta", 
        command=lambda: iniciar_coleta(selecoes, popup), bootstyle="success").grid(row=len(selecoes), column=0, pady=10)
    
    frame_popup.grid_rowconfigure(len(selecoes), weight=1)
    frame_popup.grid_columnconfigure(0, weight=1)

def iniciar_coleta(selecoes, popup):
    """Inicia a coleta com base nas seleções."""
    popup.destroy()  # Fecha o popup

    for nome, (var, func) in selecoes.items():
        if var.get():  # Se a opção foi marcada
            exibir_log(f"{nome} sendo coletado para {dados_selecionados['Razão Social']}")
            func()  # Chama a função correspondente

    messagebox.showinfo("Concluido", "Função Executada com Sucesso")

def apply_borders(ws):
    # Define o estilo da borda
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Aplica a borda a todas as células do worksheet
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

def generate_date_range(start_date, end_date):
    return pd.date_range(start=start_date, end=end_date)

def clean_json_date(date_str):
        # Remove o dia da semana (últimos 3 caracteres)
    return date_str[:10]

def process_faltas(df):
    faltas = pd.DataFrame(columns=df.columns)

    # Itera sobre cada linha do DataFrame
    for index, row in df.iterrows():
        apontamentos = row['Apontamentos']
        descontos = row['Descontos']
        horario = row['Horario']
        
        # Verifica se as palavras "Falta" ou "Atraso" estão nos campos
        if 'Falta' in str(apontamentos) or 'Atraso' in str(apontamentos) or \
           'Falta' in str(descontos) or 'Atraso' in str(descontos):
            faltas = pd.concat([faltas, row.to_frame().T])
        
        # Verifica se os Apontamentos estão ausentes e se há Horário, indicando uma falta
        elif pd.isna(apontamentos) and pd.notna(horario):
            faltas = pd.concat([faltas, row.to_frame().T])

    # Formata a coluna de data para o formato adequado
    faltas['Data'] = pd.to_datetime(faltas['Data'], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')

    return faltas

def coleta_planilha_marcacoes_faltantes():
    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",  # extensão padrão
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],  # tipos de arquivo
        title="Salvar arquivo como"
    )
    
    # URL do endpoint e payload
    url = 'https://www.dimepkairos.com.br/RestServiceApi/ReportEmployeePunch/GetReportEmployeePunch'  # Modifique para o endpoint correto
    
    # Headers da requisição
    headers = {
        "identifier": dados_selecionados["CNPJ"],
        "key": dados_selecionados["Chave API"],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }

    payload = {
        "MatriculaPessoa": [],
        "DataInicio": dados_selecionados["Data Início"],
        "DataFim": dados_selecionados["Data fim"],
        "ResponseType":"AS400V1"
    }
    
    response = requests.post(url, json=payload, headers=headers)

    if response.status_code == 200:
        try:
            data = response.json()
            if "Obj" in data and isinstance(data['Obj'], list):
                all_entries = []
                all_fixed_data = []

                # Defina o intervalo de datas do relatório
                start_date = datetime.strptime(payload["DataInicio"], "%d/%m/%Y")
                end_date = datetime.strptime(payload["DataFim"], "%d/%m/%Y")
                full_date_range = generate_date_range(start_date, end_date)

                # Carregar justificativas selecionadas
                justificativas = carregar_justificativas()

                # Itera sobre cada objeto no campo "Obj"
                for item in data['Obj']:
                    # Adiciona os dados fixos para cada funcionário
                    fixed_data = {
                        'Funcionario': item['InfoFuncionario']['Nome'],
                        'PIS': item['InfoFuncionario']['PIS'],
                        'Matricula': item['InfoFuncionario']['Matricula'],
                        'Estrutura': item['InfoFuncionario']['Estrutura']
                    }

                    # Cria um dicionário para armazenar as marcações por data
                    entradas_por_data = {}
                    for entrada in item['Entradas']:
                        # Limpa a data removendo o dia da semana
                        data_limpa = clean_json_date(entrada['Data'])
                        # Converte a data para o formato datetime
                        entrada_data_formatada = datetime.strptime(data_limpa, "%d/%m/%Y")
                        entradas_por_data[entrada_data_formatada] = entrada

                    # Para cada data no intervalo, insere os dados ou deixa em branco
                    for date in full_date_range:
                        if date in entradas_por_data:
                            entrada = entradas_por_data[date]
                            entry_data = {
                                'Data': date.strftime("%d/%m/%Y"),
                                'Horario': entrada['Horario'],
                                'Apontamentos': entrada['Apontamentos'],
                                'HTrab': entrada['HTrab'],
                                'Descontos': entrada['Descontos'],
                                'Debito': entrada['Debito']
                            }
                        else:
                            # Preenche com valores em branco para datas sem entrada
                            entry_data = {
                                'Data': date.strftime("%d/%m/%Y"),
                                'Horario': '',
                                'Apontamentos': ''
                            }

                        # Adiciona os dados fixos e de entrada combinados
                        combined_data = {**fixed_data, **entry_data}
                        all_entries.append(combined_data)

                # Converte a lista de dados em um DataFrame
                final_df = pd.DataFrame(all_entries)

                # Processa as faltas, considerando o padrão de tempo
                faltas = process_faltas(final_df)

                # Exibir as faltas (pode salvar ou processar de acordo com o seu caso)
                if not faltas.empty:
                    print("Faltas identificadas:")
                    print(faltas)
                else:
                    print("Nenhuma falta identificada.")

                # Agora, substitua o DataFrame 'final_df' pelas faltas filtradas
                final_df = faltas  # Atualizando para o DataFrame de faltas filtradas

                # Adiciona a nova coluna "Justificativa" com valores em branco inicialmente
                final_df['Justificativa'] = ''
                final_df['Entrada'] = ''
                final_df['Almoço Ida'] = ''
                final_df['Almoço Volta'] = ''
                final_df['Saida'] = ''
                final_df['Empresa'] = item['InfoEmpresa']['Nome']

                # Salva o arquivo Excel
                final_df.to_excel(output_path, index=False)

                # Cria um Workbook do openpyxl
                wb = openpyxl.load_workbook(output_path)
                ws = wb.active

                # Criar a validação de dados para a lista de justificativas selecionadas
                if justificativas:
                    justificativa_str = 'Folga,' + ','.join(justificativas)  # A lista de justificativas selecionadas
                    # Criar a validação de dados para a lista suspensa
                    justificativa_validation = DataValidation(
                        type='list', 
                        formula1=f'"{justificativa_str}"', 
                        allow_blank=True
                    )
                    justificativa_validation.error = 'Escolher valores da lista'
                    justificativa_validation.errorTitle = 'Entrada Invalida'
                    justificativa_validation.prompt = 'Selecione uma justificativa'
                    justificativa_validation.promptTitle = 'Justificativas'
                    
                    # Adiciona a validação de dados na coluna "Justificativa"
                    justificativa_col_index = final_df.columns.get_loc("Justificativa") + 1
                    justificativa_col_letter = get_column_letter(justificativa_col_index)
                    ws.add_data_validation(justificativa_validation)
                    justificativa_validation.add(f'{justificativa_col_letter}2:{justificativa_col_letter}{len(final_df)+1}')

                    # Ajusta a largura das colunas automaticamente
                    for col in ws.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        ws.column_dimensions[column].width = adjusted_width

                    # Formatação de cor alternada nas linhas
                    fill_grey = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
                    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                        for cell in row:
                            if cell.row % 2 == 0:
                                cell.fill = fill_grey
                            else:
                                cell.fill = fill_white
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                        
                    apply_borders(ws)

                    # Salva o arquivo Excel com formatação
                    wb.save(output_path)
                    print(f"Arquivo Excel salvo com sucesso em {output_path}")
                else:
                    print("Nenhuma justificativa selecionada.")
            else:
                print("Nenhum dado no campo 'Obj'.")
        except ValueError as e:
            print(f"Erro ao decodificar JSON: {e}")
            print("Conteúdo da resposta:")
            print(response.text)
    else:
        print(f"Falha: {response.status_code}")
        print(response.text)
    
    exibir_log(f'Planilha de faltas e atrasos salva em {output_path}')
    print(payload)

def filtra_marcacoes_impares_e_htrab_vazio(entradas):
    marcacoes_filtradas = []

    for entrada in entradas:
        apontamentos = entrada.get("Apontamentos", "").strip()
        htrab = entrada.get("HTrab", "").strip()
        
        # Verifica se o campo HTrab está vazio e se há apontamentos
        if not htrab and apontamentos:
            # Divide os apontamentos usando o espaço como delimitador
            intervals = apontamentos.split()
            
            # Verifica cada intervalo para identificar se algum está incompleto
            has_incomplete_turn = False
            for interval in intervals:
                # Verifica se o intervalo tem apenas uma marcação ou uma entrada sem saída
                if '-' not in interval or interval.count('-') == 1 and interval.endswith('-'):
                    has_incomplete_turn = True
                    break
            
            # Se houver um turno incompleto, adiciona a entrada aos filtrados
            if has_incomplete_turn:
                marcacoes_filtradas.append(entrada)

    return marcacoes_filtradas

def coleta_planilha_marcacoes_inconsistencia():
    output_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    title="Salvar planilha de inconsistências"
                    )
    
    url = 'https://www.dimepkairos.com.br/RestServiceApi/ReportEmployeePunch/GetReportEmployeePunch'
    
    headers = {
            "identifier": dados_selecionados["CNPJ"],
            "key": dados_selecionados["Chave API"],
            'User-Agent': 'PostmanRuntime/7.30.0'
    }

    payload = {
            "MatriculaPessoa": [],
            "DataInicio": dados_selecionados["Data Início"],
            "DataFim": dados_selecionados["Data fim"],
            "ResponseType": "AS400V1"
    }
    
    response = requests.post(url, json=payload, headers=headers)

    if response.status_code == 200:
        try:
            data = response.json()
            if "Obj" in data and isinstance(data['Obj'], list) and len(data['Obj']) > 0:
                all_entries = []
                
                for item in data['Obj']:
                    fixed_data = {
                        'Funcionario': item['InfoFuncionario']['Nome'],
                        'Matricula': item['InfoFuncionario']['Matricula'],
                        'Estrutura': item['InfoFuncionario']['Estrutura']
                    }

                    for entrada in item['Entradas']:
                        data_limpa = clean_json_date(entrada['Data'])
                        entrada_data_formatada = datetime.strptime(data_limpa, "%d/%m/%Y")

                        # Substituímos a chamada para filtra_marcacoes_impares pela nova função
                        marcacoes_impares = filtra_marcacoes_impares_e_htrab_vazio([entrada])
                        if marcacoes_impares:
                            for marcacao in marcacoes_impares:
                                entry_data = {
                                    'Data': entrada_data_formatada.strftime("%d/%m/%Y"),
                                    'Horario': marcacao['Horario'],
                                    'Apontamentos': marcacao['Apontamentos'],
                                    'HTrab': entrada['HTrab'],
                                    'Desconto': entrada['Descontos'],
                                    'Debito': entrada['Debito']
                                }
                                # Combina dados fixos e dados filtrados e os adiciona ao DataFrame
                                combined_data = {**fixed_data, **entry_data}
                                all_entries.append(combined_data)

                # Verifica se há entradas antes de criar o DataFrame
                if all_entries:
                    final_df = pd.DataFrame(all_entries)

                    final_df['Entrada'] = ''
                    final_df['Almoço Ida'] = ''
                    final_df['Almoço Volta'] = ''
                    final_df['Saida'] = ''
                    final_df['Empresa'] = item['InfoEmpresa']['Nome']
                    
                    inconsistencia_df = final_df
                    
                    # Salvar o DataFrame com as marcações inconsistentes em Excel
                    inconsistencia_df.to_excel(output_path, index=False)

                    # Aplicar o estilo no arquivo salvo
                    aplicar_estilo(output_path)
                    
                    exibir_log(f'Planilha de Inconsistência da empresa {dados_selecionados["Razão Social"]} em {output_path}')
                    
                else:
                    exibir_log("Nenhuma marcação válida encontrada para o período especificado.")
            else:
                exibir_log("Nenhum dado disponível no campo 'Obj'.")
        except ValueError as e:
            exibir_log(f"Erro ao decodificar JSON: {e}")
            exibir_log("Conteúdo da resposta:")
            exibir_log(response.text)
    else:
        exibir_log(f"Falha: {response.status_code}")
        exibir_log(response.text)

    # Caso não haja dados válidos, exibe mensagem sem tentar salvar
    if response.status_code == 200 and ('Obj' not in data or not data['Obj']):
        exibir_log("Não há dados válidos para salvar a planilha.")

def process_incomum(df):
    incomuns = pd.DataFrame(columns=df.columns)

    for index, row in df.iterrows():
        apontamentos = row['Apontamentos']
        
        # Verifica se o campo Apontamentos não está vazio
        if not apontamentos:
            continue

        apontamentos_parts = re.findall(r'\d{2}:\d{2}', str(apontamentos))
        horario = row['Horario']
        horario_parts = re.findall(r'\d{2}:\d{2}', str(horario))
        
        # Verifica se o número de intervalos é par e diferente do número de horários, ou se há textos especiais
        if len(apontamentos_parts) > 0 and len(horario_parts) > 0:
            apontamentos_length = len(apontamentos_parts)
            horario_length = len(horario_parts)
            if apontamentos_length % 2 == 0 and apontamentos_length > 0 and apontamentos_length != horario_length:
                incomuns = pd.concat([incomuns, row.to_frame().T])

    # Formata a coluna de data
    incomuns['Data'] = pd.to_datetime(incomuns['Data'], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')
    return incomuns

def coleta_planilha_marcacoes_incomum():
    output_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    title="Salvar planilha de marcação incomum"
                    )
    
    url = 'https://www.dimepkairos.com.br/RestServiceApi/ReportEmployeePunch/GetReportEmployeePunch'
    
    headers = {
            "identifier": dados_selecionados["CNPJ"],
            "key": dados_selecionados["Chave API"],
            'User-Agent': 'PostmanRuntime/7.30.0'
    }

    payload = {
            "MatriculaPessoa": [],
            "DataInicio": dados_selecionados["Data Início"],
            "DataFim": dados_selecionados["Data fim"],
            "ResponseType": "AS400V1"
    }
    
    response = requests.post(url, json=payload, headers=headers)

    if response.status_code == 200:
        try:
            data = response.json()
            if "Obj" in data and isinstance(data['Obj'], list) and len(data['Obj']) > 0:
                all_entries = []
                
                justificativas = carregar_justificativas()
                
                for item in data['Obj']:
                    fixed_data = {
                        'Funcionario': item['InfoFuncionario']['Nome'],
                        'Matricula': item['InfoFuncionario']['Matricula'],
                        'Estrutura': item['InfoFuncionario']['Estrutura']
                    }
                    

                    for entrada in item['Entradas']:
                        data_limpa = clean_json_date(entrada['Data'])
                        entrada_data_formatada = datetime.strptime(data_limpa, "%d/%m/%Y")

                        entry_data = {
                            'Data': entrada_data_formatada.strftime("%d/%m/%Y"),
                            'Horario': entrada['Horario'],
                            'Apontamentos': entrada['Apontamentos'],
                            'HTrab': entrada['HTrab'],
                            'Desconto': entrada['Descontos'],
                            'Debito': entrada['Debito']
                        }
                        # Adiciona os dados fixos e dados da marcação
                        combined_data = {**fixed_data, **entry_data}
                        all_entries.append(combined_data)

                # Verifica se há entradas antes de criar o DataFrame
                if all_entries:
                    final_df = pd.DataFrame(all_entries)
                    
                    final_df['Justificativa'] = ''
                    final_df['Entrada'] = ''
                    final_df['Almoço Ida'] = ''
                    final_df['Almoço Volta'] = ''
                    final_df['Saida'] = ''
                    final_df['Empresa']= item['InfoEmpresa']['Nome']

                    # Processa as marcações incomuns
                    incomuns_df = process_incomum(final_df)
                    

                    # Se houver marcações incomuns, salva elas no arquivo
                    if not incomuns_df.empty:
                        # Salvar o DataFrame com as marcações incomuns em Excel
                        incomuns_df.to_excel(output_path, index=False)

                        wb = openpyxl.load_workbook(output_path)
                        ws = wb.active

                        # Criar a validação de dados para a lista de justificativas selecionadas
                        if justificativas:
                            justificativa_str = 'Folga,' + ','.join(justificativas)  # A lista de justificativas selecionadas
                            # Criar a validação de dados para a lista suspensa
                            justificativa_validation = DataValidation(
                                type='list', 
                                formula1=f'"{justificativa_str}"', 
                                allow_blank=True
                            )
                            justificativa_validation.error = 'Escolher valores da lista'
                            justificativa_validation.errorTitle = 'Entrada Invalida'
                            justificativa_validation.prompt = 'Selecione uma justificativa'
                            justificativa_validation.promptTitle = 'Justificativas'
                            
                            # Adiciona a validação de dados na coluna "Justificativa"
                            justificativa_col_index = final_df.columns.get_loc("Justificativa") + 1
                            justificativa_col_letter = get_column_letter(justificativa_col_index)
                            ws.add_data_validation(justificativa_validation)
                            justificativa_validation.add(f'{justificativa_col_letter}2:{justificativa_col_letter}{len(final_df)+1}')
                            
                            # Ajusta a largura das colunas automaticamente
                            for col in ws.columns:
                                max_length = 0
                                column = col[0].column_letter
                                for cell in col:
                                    try:
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(cell.value)
                                    except:
                                        pass
                                adjusted_width = (max_length + 2)
                                ws.column_dimensions[column].width = adjusted_width

                            # Formatação de cor alternada nas linhas
                            fill_grey = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
                            fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                            
                            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                                for cell in row:
                                    if cell.row % 2 == 0:
                                        cell.fill = fill_grey
                                    else:
                                        cell.fill = fill_white
                                    cell.alignment = Alignment(horizontal="left", vertical="center")
                                    
                        apply_borders(ws)
                        wb.save(output_path)
                        exibir_log(f'Planilha de Ponto Incomum da empresa {dados_selecionados["Razão Social"]} em {output_path}')
                    else:
                        exibir_log("Não há marcações incomuns para salvar.")
                    
                else:
                    exibir_log("Nenhuma marcação válida encontrada para o período especificado.")
            else:
                exibir_log("Nenhum dado disponível no campo 'Obj'.")
        except ValueError as e:
            exibir_log(f"Erro ao decodificar JSON: {e}")
            exibir_log("Conteúdo da resposta:")
            exibir_log(response.text)
    else:
        exibir_log(f"Falha: {response.status_code}")
        exibir_log(response.text)

    # Remove o log de salvamento se não houver dados válidos
    if response.status_code == 200 and ('Obj' not in data or not data['Obj']):
        exibir_log("Não há dados válidos para salvar a planilha.")

def coleta_planilha_marcacoes():
    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",  # extensão padrão
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],  # tipos de arquivo
        title="Salvar arquivo como"
    )

    # URL do endpoint e payload
    url = 'https://www.dimepkairos.com.br/RestServiceApi/ReportEmployeePunch/GetReportEmployeePunch'  # Modifique para o endpoint correto
    
    # Headers da requisição
    headers = {
        "identifier": dados_selecionados["CNPJ"],
        "key": dados_selecionados["Chave API"],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }

    payload = {
        "MatriculaPessoa": [],
        "DataInicio": dados_selecionados["Data Início"],
        "DataFim": dados_selecionados["Data fim"],
        "ResponseType":"AS400V1"
    }
    
    response = requests.post(url, json=payload, headers=headers)

    if response.status_code == 200:
        try:
            data = response.json()
            if "Obj" in data and isinstance(data['Obj'], list):
                all_entries = []
                all_fixed_data = []

                # Defina o intervalo de datas do relatório
                start_date = datetime.strptime(payload["DataInicio"], "%d/%m/%Y")
                end_date = datetime.strptime(payload["DataFim"], "%d/%m/%Y")
                full_date_range = generate_date_range(start_date, end_date)

                # Carregar justificativas selecionadas
                justificativas = carregar_justificativas()

                # Itera sobre cada objeto no campo "Obj"
                for item in data['Obj']:
                    # Adiciona os dados fixos para cada funcionário
                    fixed_data = {
                        'Funcionario': item['InfoFuncionario']['Nome'],
                        'PIS': item['InfoFuncionario']['PIS'],
                        'Matricula': item['InfoFuncionario']['Matricula'],
                        'Estrutura': item['InfoFuncionario']['Estrutura']
                    }
                    # Cria um dicionário para armazenar as marcações por data
                    entradas_por_data = {}
                    for entrada in item['Entradas']:
                        # Limpa a data removendo o dia da semana
                        data_limpa = clean_json_date(entrada['Data'])
                        # Converte a data para o formato datetime
                        entrada_data_formatada = datetime.strptime(data_limpa, "%d/%m/%Y")
                        entradas_por_data[entrada_data_formatada] = entrada

                    # Para cada data no intervalo, insere os dados ou deixa em branco
                    for date in full_date_range:
                        if date in entradas_por_data:
                            entrada = entradas_por_data[date]
                            entry_data = {
                                'Data': date.strftime("%d/%m/%Y"),
                                'Horario': entrada['Horario'],
                                'Apontamentos': entrada['Apontamentos'],
                                'HTrab': entrada['HTrab'],
                                'Desconto': entrada['Descontos'],
                                'Debito': entrada['Debito']
                            }
                        else:
                            # Preenche com valores em branco para datas sem entrada
                            entry_data = {
                                'Data': date.strftime("%d/%m/%Y"),
                                'Horario': '',
                                'Apontamentos': ''
                            }

                        # Adiciona os dados fixos e de entrada combinados
                        combined_data = {**fixed_data, **entry_data}
                        all_entries.append(combined_data)

                # Converte a lista de dados em um DataFrame
                final_df = pd.DataFrame(all_entries)

                # Adiciona a nova coluna "Justificativa" com valores em branco inicialmente
                final_df['Justificativa'] = ''
                final_df['Qtd Horas'] = ''
                final_df['Entrada'] = ''
                final_df['Almoço Ida'] = ''
                final_df['Almoço Volta'] = ''
                final_df['Saida'] = ''
                final_df['Empresa'] = item['InfoEmpresa']['Nome']

                # Salva o arquivo Excel
                final_df.to_excel(output_path, index=False)

                # Cria um Workbook do openpyxl
                wb = openpyxl.load_workbook(output_path)
                ws = wb.active

                # Criar a validação de dados para a lista de justificativas selecionadas
                if justificativas:
                    justificativa_str = 'Folga,' + ','.join(justificativas)  # A lista de justificativas selecionadas
                    # Criar a validação de dados para a lista suspensa
                    justificativa_validation = DataValidation(
                        type='list', 
                        formula1=f'"{justificativa_str}"', 
                        allow_blank=True
                    )
                    justificativa_validation.error = 'Escolher valores da lista'
                    justificativa_validation.errorTitle = 'Entrada Invalida'
                    justificativa_validation.prompt = 'Selecione uma justificativa'
                    justificativa_validation.promptTitle = 'Justificativas'
                    
                    # Adiciona a validação de dados na coluna "Justificativa"
                    justificativa_col_index = final_df.columns.get_loc("Justificativa") + 1
                    justificativa_col_letter = get_column_letter(justificativa_col_index)
                    ws.add_data_validation(justificativa_validation)
                    justificativa_validation.add(f'{justificativa_col_letter}2:{justificativa_col_letter}{len(final_df)+1}')

                    # Ajusta a largura das colunas automaticamente
                    for col in ws.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        ws.column_dimensions[column].width = adjusted_width

                    # Formatação de cor alternada nas linhas
                    fill_grey = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
                    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                        for cell in row:
                            if cell.row % 2 == 0:
                                cell.fill = fill_grey
                            else:
                                cell.fill = fill_white
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                        
                    apply_borders(ws)

                    # Salva o arquivo Excel com formatação
                    wb.save(output_path)
                    print(f"Arquivo Excel salvo com sucesso em {output_path}")
                else:
                    print("Nenhuma justificativa selecionada.")
            else:
                print("Nenhum dado no campo 'Obj'.")
        except ValueError as e:
            print(f"Erro ao decodificar JSON: {e}")
            print("Conteúdo da resposta:")
            print(response.text)
    else:
        print(f"Falha: {response.status_code}")
        print(response.text)
    
    exibir_log(f'Planilha de Marcação da empresa {dados_selecionados["Razão Social"]} salva em {output_path}')
    print(payload)

def carregar_justificativas():
    # Recupera as justificativas selecionadas do arquivo
    try:
        with open("justificativas_selecionadas.txt", "r") as f:
            justificativas = f.readlines()
            justificativas = [line.strip() for line in justificativas]  # Limpa o texto
    except FileNotFoundError:
        justificativas = []
    
    return justificativas

def get_data_from_api(url, payload, headers):
    """Faz uma requisição para a API e transforma a resposta em DataFrame."""
    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()  # Verifica se a requisição foi bem-sucedida

        data = response.json()  # Decodifica o JSON

        # Verifica se 'Obj' é uma lista e normaliza diretamente
        if 'Obj' in data and isinstance(data['Obj'], list):
            obj_data = pd.json_normalize(data['Obj'])
            return obj_data

        # Caso 'Obj' seja uma string JSON, faz o parsing
        elif 'Obj' in data and isinstance(data['Obj'], str):
            obj_data = json.loads(data['Obj'])  # Converte a string JSON em lista
            return pd.json_normalize(obj_data)

        # Normaliza o JSON principal se 'Obj' não estiver presente
        else:
            return pd.json_normalize(data)

    except ValueError as e:
        print(f"Erro ao decodificar JSON: {e}")
        print("Conteúdo da resposta:", response.text)
    except requests.exceptions.RequestException as e:
        print(f"Erro na requisição: {e}")

def preencher_detalhes():
    """Preenche os detalhes da empresa selecionada na interface."""
    try:
        empresa_selecionada = razao_social_var.get()

        if not empresa_selecionada:
            messagebox.showwarning("Aviso", "Nenhuma empresa foi selecionada.")
            return

        # Filtra o DataFrame para encontrar a empresa correspondente
        empresa_detalhes = df_empresas[df_empresas["Razão Social"] == empresa_selecionada]

        if empresa_detalhes.empty:
            messagebox.showwarning("Aviso", "Empresa selecionada não encontrada.")
            return

        empresa_detalhes = empresa_detalhes.iloc[0]

        # Preenche os campos com os dados da empresa selecionada
        cnpj_var.set(empresa_detalhes["CNPJ"])
        chave_var.set(empresa_detalhes["Chave API"])

        # Garante que o CPF mantenha zeros à esquerda, se existir
        cpf_responsavel = str(empresa_detalhes.get("CPF Responsável", "")).zfill(11)
        cpf_var.set(cpf_responsavel)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao preencher os detalhes: {e}")

def combinar_data_hora(data_completa, hora, tipo_marcacao=""):
    # Verifica se ambos os valores estão disponíveis
    if pd.notna(data_completa) and pd.notna(hora):
        # Garante que data_completa esteja no formato string
        if isinstance(data_completa, pd.Timestamp):
            data = data_completa.strftime("%d/%m/%Y")
        else:
            data = str(data_completa)
        
        # Converte 'hora' para string sem segundos, se necessário
        hora_str = hora if isinstance(hora, str) else hora.strftime("%H:%M")
        if len(hora_str) > 5:  # Caso tenha segundos, remove-os
            hora_str = hora_str[:5]
        
        # Combina a data e hora no formato especificado
        data_hora_str = f"{data} {hora_str}"
        try:
            # Tenta converter no formato %H:%M
            data_hora_dt = datetime.strptime(data_hora_str, "%d/%m/%Y %H:%M")
            
            # Ajusta a data para o próximo dia se for uma marcação noturna
            if tipo_marcacao != "Entrada":
                if data_hora_dt.hour == 23 and data_hora_dt.minute == 59 or data_hora_dt.hour < 6:
                    data_hora_dt += timedelta(days=1)
                    
            return data_hora_dt.strftime("%d/%m/%Y %H:%M")
        
        except ValueError:
            try:
                # Caso tenha segundos, tenta remover e converter novamente
                data_hora_dt = datetime.strptime(data_hora_str, "%d/%m/%Y %H:%M:%S")
                
                # Ajusta a data para o próximo dia se for uma marcação noturna
                if tipo_marcacao != "Entrada":
                    if data_hora_dt.hour == 23 and data_hora_dt.minute == 59 or data_hora_dt.hour < 6:
                        data_hora_dt += timedelta(days=1)
                        
                return data_hora_dt.strftime("%d/%m/%Y %H:%M")
            
            except ValueError as e:
                print(f"Erro ao converter data e hora: {e}")
                return None
    return None

def processar_marcacoes(df):
    if df.empty:
        exibir_log("Arquivo sem dados de marcações. Ignorando operação.")
        return

    url = "https://www.dimepkairos.com.br/RestServiceApi/Mark/SetMarks"
    headers = {
        "identifier": dados_selecionados['CNPJ'],
        "key": dados_selecionados['Chave API'],
        "cpf": dados_selecionados['CPF Responsável'],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }

    for index, row in df.iterrows():
        matricula = row.get('Matricula')
        data_completa = row.get('Data')

        marcacoes = {
            "Entrada": row.get('Entrada'),
            "Almoço Ida": row.get('Almoço Ida'),
            "Almoço Volta": row.get('Almoço Volta'),
            "Saida": row.get('Saida')
        }

        for tipo, hora in marcacoes.items():
            if pd.notna(hora):
                data_hora_marcacao = combinar_data_hora(data_completa, hora, tipo)
                if data_hora_marcacao:
                    payload = {
                        "Matricula": matricula,
                        "DataHoraApontamento": data_hora_marcacao,
                        "CpfResponsavel": dados_selecionados['CPF Responsável'],
                        "ResponseType": "AS400V1"
                    }
                    response = requests.post(url, json=payload, headers=headers)
                    status = "Sucesso" if response.status_code == 200 else "Falha"
                    mensagem = response.json().get("Mensagem", "Erro desconhecido") if response.ok else response.text
                    exibir_log(f"{tipo} enviada para {matricula}: {status} - {mensagem}")

def cadastrar_pessoas():
    """Cadastra pessoas na API utilizando dados do Excel."""
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de cadastro de pessoas",
        filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
    )
    if not caminho_arquivo:
        return

    try:
        df = pd.read_excel(caminho_arquivo, dtype={"CPF": str})
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return
    resultados = []
    url = "https://www.dimepkairos.com.br/RestServiceApi/People/SavePerson"
    headers = {
        "identifier": dados_selecionados["CNPJ"],
        "key": dados_selecionados["Chave API"],
        "cpf": dados_selecionados['CPF Responsável'],
        "User-Agent": "PostmanRuntime/7.30.0"
    }

    # Iteração sobre cada pessoa no DataFrame e envio da requisição POST
    for _, row in df.iterrows():
        try:
            Matricula = row['Matricula']
            Cracha = row['Cracha']
            Nome = row['Nome Completo']
            Cpf = str(row['CPF']).zfill(11)
            Email = row.get('Email', None)
            DataAdm = pd.to_datetime(row["Admissão"]).strftime('%d-%m-%Y')
            DataNasc = pd.to_datetime(row["Nascimento"]).strftime('%d-%m-%Y')
            BaseHoras = row['Base de Horas']
            EstruturaOrg = row['Estrutura']
            Horario = row['Horário']
            RegraCalculo = row['Cálculo']
            Pis = row.get('PIS', None)
            PisAuto = row.get('Possui PIS?', 1)
            Sexo = row['Sexo']
            Cargo = str(row.get('Cargo', None))
            CampoAlternativo = row.get('Campo Alternativo', None)
            
            payload = {
                "Matricula": Matricula,
                "Cracha": Cracha,
                "Nome": Nome,
                "Cpf": Cpf,
                "CpfResponsavel": dados_selecionados["CPF Responsável"],
                "DataAdmissao": DataAdm,
                "DataNascimento": DataNasc,
                "BaseHoras": BaseHoras,
                "Estrutura": {"Id": EstruturaOrg},
                "TipoFuncionario": {"IdTipoFuncionario": 1},
                "TipoSalario": {"Id": 101},
                "Horarios": [
                    {
                        "Id": 0,
                        "Horario": {"Id": Horario},
                        "Inicio": DataAdm,
                        "Fim": "31/12/9999 00:00:00"
                    }
                ],
                "RegrasCalculo": [
                    {
                        "Id": 0,
                        "Regra": {"Id": RegraCalculo},
                        "Inicio": DataAdm,
                        "Fim": "31/12/9999 00:00:00"
                    }
                ],
                "FlagGerarNumeroPISAutomatico": bool(PisAuto == 0),
                "Sexo": Sexo
            }
            
            if pd.notna(Email) and str(Email).strip():
                payload["Email"] = str(Email).strip()
                
            if pd.notna(CampoAlternativo) and str(CampoAlternativo).strip():
                payload["CampoAlternativo1"] = str(CampoAlternativo).strip()   
            
            if pd.notna(Cargo):
                payload['Cargo'] = {"Id":Cargo}

            # Adicionar PIS se necessário
            if PisAuto == 1 and Pis:
                payload["CodigoPis"] = Pis

            # Enviar a requisição
            response = requests.post(url, json=payload, headers=headers)
            response_json = response.json()

            # Verificar sucesso ou falha
            if response_json.get("Sucesso"):
                status = "Sucesso"
                mensagem = response_json.get("Mensagem", "")
            else:
                status = "Falha"
                mensagem = response_json.get("Mensagem", "Erro desconhecido")

        except Exception as e:
            status = "Erro"
            mensagem = str(e)

        # Armazenar o resultado
        resultados.append([Matricula, status, mensagem])
        exibir_log(f"Matricula: {Matricula} | Status: {status} | Mensagem: {mensagem}")

def coleta_empresa():
    url1= 'https://www.dimepkairos.com.br/RestServiceApi/CalculationRules/GetCalculationRulesSummary' # Cálculo
    url2= 'https://www.dimepkairos.com.br/RestServiceApi/Schedules/GetSchedulesSummary' # Horário
    url3= 'https://www.dimepkairos.com.br/RestServiceApi/OrganizationalStructure/GetOrganizationalStructure' # Estrutura
    url4= 'https://www.dimepkairos.com.br/RestServiceApi/JobPosition/SearchJobPosition' # Cargo
    url5= 'https://www.dimepkairos.com.br/RestServiceApi/Company/GetCompany' # Empresa
    
    payload1={}
    payload2 = {"ResponseType": "AS400V1"}
    
    headers = {
        'identifier': dados_selecionados["CNPJ"],
        'key': dados_selecionados['Chave API'],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }
    
    df1 = get_data_from_api(url1, payload1, headers)
    df2 = get_data_from_api(url2, payload1, headers)
    df3 = get_data_from_api(url3, payload1, headers)  # Já tratado dentro da função
    df4 = get_data_from_api(url4, payload1, headers)
    df5 = get_data_from_api(url5, payload2, headers)
    
    # Adiciona prefixos para diferenciar as colunas dos DataFrames
    df1 = df1.add_prefix('Calculo_')
    df2 = df2.add_prefix('Horario_')
    df4 = df4.add_prefix('Cargo_')
    df5 = df5.add_prefix('Empresa_')

    # Combina os DataFrames horizontalmente (por colunas)
    combined_df = pd.concat([df1, df2, df3, df4, df5], axis=1)

    # Remove espaços extras dos nomes das colunas
    combined_df.columns = combined_df.columns.str.strip()

    # Seleciona as colunas desejadas, verificando se estão presentes
    selected_columns = [
        'Empresa_name', 'Empresa_CnpjCpf', 'Calculo_Id', 'Calculo_Descricao', 'Horario_Id', 'Horario_Descricao', 'Id', 'Description',
        'Cargo_Id', 'Cargo_Descricao'
    ]
    selected_columns = [col for col in selected_columns if col in combined_df.columns]
    
    for col in selected_columns:
        if col not in combined_df.columns:
            combined_df[col] = None
    
    
    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",  # extensão padrão
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],  # tipos de arquivo
        title="Salvar arquivo como"
    )

    # Cria um DataFrame final apenas com as colunas selecionadas
    final_df = combined_df[selected_columns]
    
    first_sheet_df = pd.DataFrame(columns=["Matricula", "Cracha", "Nome Completo", "CPF", "Email", "Admissão", "Nascimento", "Base de Horas",
                                        "Estrutura", "Horário", "Cálculo", "Cargo", "Possui PIS?", "PIS", "Sexo", "Campo Alternativo"])

    # Escrevendo múltiplas planilhas no mesmo arquivo Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        first_sheet_df.to_excel(writer, sheet_name='Cadastro', index=False)
        final_df.to_excel(writer, sheet_name='Dados Empresa', index=False)
        
    # Carrega o Workbook do openpyxl para adicionar validações e ajustar colunas
    wb = openpyxl.load_workbook(output_path)
    ws = wb['Cadastro']

    # Ajusta a largura das colunas automaticamente
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Pega a letra da coluna
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column].width = max_length + 2
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="196B24", end_color="196B24", fill_type="solid")


    ws = wb['Dados Empresa']

    # Ajusta a largura das colunas automaticamente
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Pega a letra da coluna
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column].width = max_length + 2
        
    ws.protection.sheet = True  # Ativa a proteção da planilha
    ws.protection.password = "123456789a"  # Define uma senha (opcional)

    # Impede alterações de conteúdo, mas permite dimensionar células e selecionar
    ws.protection.enable()

    # Ajustes opcionais de permissões
    ws.protection.allow_format_cells = True  # Permite formatar células
    ws.protection.allow_select_locked_cells = True  # Permite selecionar células bloqueadas

    # Salva o arquivo Excel com as alterações
    wb.save(output_path)
    
    exibir_log(f'Planilha de Cadastro salva em {output_path}')

def coleta_planilha_ferias():
    
    output_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    title="Salvar planilha de inconsistências"
                    )
    
    url = 'https://www.dimepkairos.com.br/RestServiceApi/People/SearchPeople'
    payload = {
        "Matricula": 0 
    }
    headers = {
        "identifier": dados_selecionados['CNPJ'],
        "key": dados_selecionados['Chave API'],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }
    
    df = get_data_from_api(url, payload, headers)
    df = df.add_prefix('Colaborador_')
    
    selected_colomns =[ 'Colaborador_Matricula','Colaborador_Nome']
    
    
    df_filtrado = df[selected_colomns]
    final_df = pd.DataFrame(df_filtrado)
    final_df['Data Inicio'] = ''
    final_df['Data Fim'] = ''
    
    ferias_df = final_df
    
    ferias_df.to_excel(output_path, index=False)
    aplicar_estilo(output_path)
    
    exibir_log(f'Planilha de feriass salva em {output_path}')

#Em breve
def envio_planilha_ferias():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de cadastro de pessoas",
        filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
    )
    if not caminho_arquivo:
        return
    
    try:
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return
    
    
    resultados = []
    df = df.fillna('')  # Preenche NaNs com string vazia
    url = 'https://www.dimepkairos.com.br/RestServiceApi/Holiday/MarkHoliday'
    headers = {
        'identifier': dados_selecionados['CNPJ'],
        'key': dados_selecionados['Chave API'],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }
    
    for index, row in df.iterrows():
        try:
            Matricula = row['Colaborador_Matricula']
            Data_inicio = row['Data Inicio']
            Data_fim = row['Data Fim']
            
            
            try:
                if isinstance(Data_inicio, datetime):
                    data_formatada_inicio = Data_inicio.strftime("%Y-%m-%d")
                else:
                    data_formatada_inicio = datetime.strptime(str(Data_inicio), "%d/%m/%Y").strftime("%Y-%m-%d")
            except ValueError:
                exibir_log(f"Data Inicio invalida para matrícula {Matricula}. Ignorando.")
                continue

            try:
                if isinstance(Data_inicio, datetime):
                    data_formatada_fim = Data_fim.strftime("%Y-%m-%d")
                else:
                    data_formatada_fim = datetime.strptime(str(Data_fim), "%d/%m/%Y").strftime("%Y-%m-%d")
            except ValueError:
                exibir_log(f"Data Final invalida para matrícula {Matricula}. Ignorando.")
                continue
            
            payload = {
                "MATRICULA": Matricula,
                "DATAINICIO": data_formatada_inicio,
                "DATAFIM": data_formatada_fim,
                "ConfirmacaoFeriasPrimeiroPeriodoAquisitivo": 'true',
            }
            
            response = requests.post(url, json=payload, headers=headers)
            response_json = response.json()

            # Verificar sucesso ou falha
            if response_json.get("Sucesso"):
                status = "Sucesso"
                mensagem = response_json.get("Mensagem", "")
                exibir_log(f'Ferias enviadas com SUCESSO para a {Matricula}')
            else:
                status = "Falha"
                mensagem = response_json.get("Mensagem", "Erro desconhecido")

        except Exception as e:
            status = "Erro"
            mensagem = str(e)
            
            # Armazenar o resultado
        resultados.append([Matricula, status, mensagem])

def alteracao_pessoas():
    url = 'https://www.dimepkairos.com.br/RestServiceApi/People/SearchPeople'
    payload = {
        "Matricula": 0 
    }
    headers = {
        "identifier": dados_selecionados['CNPJ'],
        "key": dados_selecionados['Chave API'],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }
    
    df = get_data_from_api(url, payload, headers)
    df = df.add_prefix('People_')
    
    selected_colomns =[
        'People_Id', 'People_Matricula', 'People_Cracha', 'People_Nome', 'People_DataNascimento', 'People_DataAdmissao',
    'People_Rg', 'People_Cpf', 'People_Email', 'People_BaseHoras', 'People_CodigoPis', 'People_Sexo', 'People_PessoaStatus', 'People_Estrutura.Id', 'People_Estrutura.Codigo',  'People_TipoSalario.Id',
    'People_Cargo.Id'
    ]
    
    # Verifica e preenche colunas ausentes com valores None
    for col in selected_colomns:
        if col not in df.columns:
            df[col] = None
    
    
    df_filtrado = df[selected_colomns]
    df_filtrado = aplicar_estilo()
    
    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",  # extensão padrão
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],  # tipos de arquivo
        title="Salvar arquivo como"
    )
    
    df_filtrado.to_excel(output_path, index=False)
    
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    for col in ws.columns:
        max_length = 0
    column = col[0].column_letter
    for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
            except:
                pass
            
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

    wb.save(output_path)
    
    exibir_log(f'Planilha de funcionários salva em {output_path}')

def alteracao_pessoas_envio():
    
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de cadastro de pessoas",
        filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
    )
    if not caminho_arquivo:
        return

    try:
        df = pd.read_excel(caminho_arquivo, dtype={"People_Cpf": str})
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return

    resultados = []
    df = df.fillna('')  # Preenche NaNs com string vazia
    url = 'https://www.dimepkairos.com.br/RestServiceApi/People/ChangePerson'
    headers = {
        'identifier': dados_selecionados['CNPJ'],
        'key': dados_selecionados['Chave API'],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }

    for index, row in df.iterrows():
        try:
            Id = row['People_Id']
            Matricula = row['People_Matricula']
            Cracha = row['People_Cracha']
            Nome = row['People_Nome']
            DataAdm = row['People_DataAdmissao']
            Rg = row.get('People_Rg', ' ')
            Cpf = str(row['People_Cpf'].replace(".","").replace("-",""))
            
            # Verifica se a coluna Email existe no DataFrame
            if 'People_Email' in df.columns:
                Email = row['People_Email']
                Email = str(Email).strip() if pd.notna(Email) else None  # Converte para string e remove espaços
                if Email == '':  # Verifica se a string está vazia
                    Email = None  # Define como None se a string for vazia
            else:
                Email = None  # Coluna não existe, então Email é None

            BaseHoras = row['People_BaseHoras']
            Pis = row['People_CodigoPis']
            Sexo = row['People_Sexo']
            EstruturaId = row['People_Estrutura.Id']
            TipoSalario = row['People_TipoSalario.Id']
            Cargo = row['People_Cargo.Id']

            # Verifica se o PIS está vazio ou ausente e trata como automático
            Pis = None if pd.isna(Pis) or Pis == '' else str(Pis)

            # Se PIS estiver vazio ou for uma string vazia, define PIS automático
            if Pis is None:
                payload = {
                    "Id": Id,
                    "Matricula": Matricula,
                    "Cracha": Cracha,
                    "Nome": Nome,
                    "DataAdmissao": DataAdm,
                    "Rg": Rg,
                    "Cpf": Cpf,
                    "BaseHoras": BaseHoras,
                    "Estrutura": {"Id": EstruturaId},
                    "TipoFuncionario": {"Id": 1},
                    "TipoSalario": {"Id": 101},
                    "Sexo": Sexo,
                    "FlagGerarNumeroPISAutomatico": True,  # Gera automaticamente o PIS
                    "CpfResponsavel": dados_selecionados["CPF Responsável"]
                }
            else:
                # Caso PIS esteja preenchido, envia normalmente
                payload = {
                    "Id": Id,
                    "Matricula": Matricula,
                    "Cracha": Cracha,
                    "Nome": Nome,
                    "DataAdmissao": DataAdm,
                    "Rg": Rg,
                    "Cpf": Cpf,
                    "BaseHoras": BaseHoras,
                    "Estrutura": {"Id": EstruturaId},
                    "TipoFuncionario": {"Id": 1},
                    "TipoSalario": {"Id": 101},
                    "Sexo": Sexo,
                    "FlagGerarNumeroPISAutomatico": False,  # Não gera o PIS automático
                    "Pis": Pis,  # Envia o PIS preenchido
                    "CpfResponsavel": dados_selecionados["CPF Responsável"]
                }

            # Exclui o Email do payload se for None
            if Email is not None:
                payload["Email"] = Email

            print(payload)  # Para verificar o payload antes de enviar
            response = requests.post(url, json=payload, headers=headers)
            response_json = response.json()

            # Verificar sucesso ou falha
            if response_json.get("Sucesso"):
                status = "Sucesso"
                mensagem = response_json.get("Mensagem", "")
            else:
                status = "Falha"
                mensagem = response_json.get("Mensagem", "Erro desconhecido")

        except Exception as e:
            status = "Erro"
            mensagem = str(e)

        # Armazenar o resultado
        resultados.append([Matricula, status, mensagem])
        
        exibir_log(f"Matricula: {Matricula} | Status: {status} | Mensagem: {mensagem}")

def exibir_log(mensagem):
    """Exibe o log de mensagens na interface."""
    log_widget.configure(state="normal")
    log_widget.insert(tk.END, mensagem + "\n")
    log_widget.see(tk.END)
    log_widget.configure(state="disabled")

def confirmar_selecao():
    """Confirma a seleção da empresa e salva as credenciais globalmente."""
    global dados_selecionados
    if not razao_social_var.get():
        messagebox.showwarning("Aviso", "Selecione uma empresa.")
        return
    
    cpf_responsavel = cpf_var.get()[:11]
    cpf_var.set(cpf_responsavel)  # Atualiza a variável com o valor truncado
    
    dados_selecionados = {
        "Razão Social": razao_social_var.get(),
        "CNPJ": cnpj_var.get(),
        "Chave API": chave_var.get(),
        "CPF Responsável": cpf_responsavel,
        "Data Início" : data_inicio_var.get(),
        "Data fim": data_fim_var.get()
    }
    messagebox.showinfo("Sucesso", "Informações selecionadas com sucesso!")

def formatar_data(entry, new_value):
    # Remove tudo que não for número
    new_value = ''.join([c for c in new_value if c.isdigit()])
    
    # Limita a quantidade de caracteres para 8 (ddmmyyyy)
    if len(new_value) > 8:
        new_value = new_value[:8]

    # Aplica a formatação de data dd/mm/aaaa
    if len(new_value) > 2:
        new_value = new_value[:2] + '/' + new_value[2:]
    if len(new_value) > 5:
        new_value = new_value[:5] + '/' + new_value[5:]

    # Atualiza a entrada com a data formatada
    entry.delete(0, tk.END)
    entry.insert(0, new_value)
    
def funcao_justificativa_get():
    if not dados_selecionados.get('CNPJ') or not dados_selecionados.get('Chave API'):
        messagebox.showerror("Erro", "Nenhuma empresa selecionada. Selecione uma empresa antes de carregar justificativas.")
        return

    def carregar_justificativas():
        try:
            payload = {"Code": 0, "IdType": 1202, "ResponseType": "AS400V1"}
            headers = {
                "identifier": dados_selecionados['CNPJ'],
                "key": dados_selecionados['Chave API'],
                'User-Agent': 'PostmanRuntime/7.30.0'
            }
            url = 'https://www.dimepkairos.com.br/RestServiceApi/Justification/GetJustification'
            response = requests.post(url, json=payload, headers=headers)
            response.raise_for_status()
            data = response.json()

            for widget in frame_checkboxes.winfo_children():
                widget.destroy()

            for item in data.get("Obj", []):
                var = tk.BooleanVar()
                checkbox = ttk.Checkbutton(
                    frame_checkboxes,
                    text=f"{item['Description']} (ID: {item['Id']})",
                    variable=var
                )
                checkbox.grid(sticky="w", padx=10, pady=5)
                justificativas_selecionadas[item['Id']] = (var, item['Description'])

        except requests.RequestException as e:
            exibir_log(f"Erro ao carregar justificativas: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar justificativas: {e}")

    # Função para gravar as justificativas selecionadas
    def gravar_justificativas():
        selecionadas = [(id_justificativa, descricao) for id_justificativa, (var, descricao) in justificativas_selecionadas.items() if var.get()]
        
        if not selecionadas:
            messagebox.showwarning("Atenção", "Nenhuma justificativa selecionada.")
            return
        
        with open("justificativas_selecionadas.txt", "w") as f:
            for _, descricao in selecionadas:
                f.write(f"{descricao}\n")
        
        with open("justificativas_ids_descricoes.txt", "w") as f:
            for id_justificativa, descricao in selecionadas:
                f.write(f"{id_justificativa},{descricao}\n")
        
        messagebox.showinfo("Sucesso", "Justificativas selecionadas foram gravadas.")
        popup_futuro.destroy()

    popup_futuro = tk.Toplevel(root)
    popup_futuro.title("Seleção de Justificativas")
    popup_futuro.geometry("400x380")

    canvas = tk.Canvas(popup_futuro)
    scrollbar = ttk.Scrollbar(popup_futuro, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    frame_checkboxes = ttk.Frame(canvas)

    canvas.create_window((0, 0), window=frame_checkboxes, anchor="nw")

    scrollbar.pack(side="right", fill="y")
    canvas.pack(fill="both", expand=True)

    justificativas_selecionadas = {}

    carregar_justificativas()

    frame_checkboxes.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    ttk.Button(
        popup_futuro, text="Gravar Seleção", command=gravar_justificativas, bootstyle="success"
    ).pack(pady=10)

    ttk.Button(
        popup_futuro, text="Fechar", command=popup_futuro.destroy, bootstyle="danger"
    ).pack(pady=10)

def selecionar_arquivo_empresas():
    """Abre uma janela para o usuário selecionar o arquivo de empresas."""
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Empresas",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if not caminho_arquivo:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
        return
    
    try:
        # Recarrega o DataFrame com o novo arquivo, garantindo CPF como string
        global df_empresas
        df_empresas = pd.read_excel(caminho_arquivo, dtype={"CPF Responsável": str})

        # Atualiza a combobox com os novos valores
        combo_razao_social['values'] = df_empresas["Razão Social"].tolist()
        combo_razao_social.set('')  # Limpa a seleção atual
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")
        return

def carregar_ids_justificativas():
    ids_justificativas = {}
    with open("justificativas_ids_descricoes.txt", "r") as file:
        for line in file:
            id_just, descricao = line.strip().split(",")
            ids_justificativas[descricao] = int(id_just)
    return ids_justificativas

def enviar_justificativa(id_funcionario, id_justificativa, descricao, data, qtd_horas):
    url = "https://www.dimepkairos.com.br/RestServiceApi/PreJustificationRequest/PreJustificationRequest"

    headers = {
        'identifier': dados_selecionados["CNPJ"],
        'key': dados_selecionados["Chave API"],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }
    payload = {
        "IdJustification": id_justificativa,
        "IdUser": '1',
        "IdEmployee": id_funcionario,
        "QtdHours": qtd_horas,
        "Date": data,
        "Notes" : "Enviado via API",
        "RequestType":"1",
        "ResponseType":"AS400V1"
    }
    print(f"Enviando justificativa com payload: {payload}")
    
    try:
        data_formatada = datetime.strptime(data, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        data_formatada = data  # Se a data não estiver no formato esperado, mantém o original
    
    response = requests.post(url, json=payload, headers=headers)
    if response.status_code == 200:
        exibir_log(f"Justificativa registrada para funcionário {id_funcionario} - Dia: {data_formatada} - Qtd. Horas {qtd_horas}")
    else:
        exibir_log(f"Erro ao enviar justificativa para ID {id_funcionario}: {response.status_code}")

def buscar_dados_funcionario(matricula):
    """Busca CPF e ID do funcionário pela matrícula."""
    url = "https://www.dimepkairos.com.br/RestServiceApi/People/SearchPerson"
    headers = {
        'identifier': dados_selecionados["CNPJ"],
        'key': dados_selecionados["Chave API"],
        'User-Agent': 'PostmanRuntime/7.30.0'
    }
    payload = {"Matricula": matricula}

    response = requests.post(url, json=payload, headers=headers)

    if response.status_code == 200:
        try:
            data_response = response.json()
            funcionarios = json.loads(data_response['Obj']) if isinstance(data_response['Obj'], str) else data_response['Obj']
            funcionario = funcionarios[0]  # Primeiro funcionário encontrado
            return funcionario.get('Cpf'), funcionario.get('Id')
        except (json.JSONDecodeError, IndexError, KeyError):
            exibir_log(f"Erro ao processar resposta da API para matrícula {matricula}.")
    else:
        exibir_log(f"Erro na API ao buscar dados do funcionário {matricula}: {response.status_code}")
    return None, None

def validar_horas(qtd_horas):
    """Valida e formata as horas."""
    if not qtd_horas or str(qtd_horas).strip() == "":
        return "12:00"
    try:
        # Tenta validar como HH:mm:ss e converter para HH:mm
        return datetime.strptime(str(qtd_horas), "%H:%M:%S").strftime("%H:%M")
    except ValueError:
        try:
            # Tenta validar diretamente como HH:mm
            return datetime.strptime(str(qtd_horas), "%H:%M").strftime("%H:%M")
        except ValueError:
            return "12:00"

def processar_arquivo_excel(df):
    """
    Processa o DataFrame para envio de justificativas e registros de folgas.
    """
    if df.empty:
        exibir_log("Arquivo vazio. Operação abortada.")
        return

    ids_justificativas = carregar_ids_justificativas()

    tem_folga = any(row.get('Justificativa') == "Folga" for index, row in df.iterrows())

    caminho_folgas = None  

    if tem_folga:
        caminho_folgas = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt")],
            title="Selecione o local para salvar o arquivo de folgas"
        )
        if not caminho_folgas:
            exibir_log("Nenhum local de salvamento selecionado para o arquivo de folgas. Ignorando folgas.")
            tem_folga = False
    else:
        exibir_log("Nenhuma justificativa de folga encontrada. Ignorando geração do arquivo de folgas.")

    if tem_folga and caminho_folgas:
        with open(caminho_folgas, "w") as arquivo_folgas:
            for index, row in df.iterrows():
                matricula = row.get('Matricula')
                descricao_justificativa = row.get('Justificativa')
                data = row.get('Data')
                pis = row.get('PIS')

                if descricao_justificativa == "Folga":
                    try:
                        if isinstance(data, datetime):
                            data_formatada = data.strftime("%Y-%m-%d")
                        else:
                            data_formatada = datetime.strptime(str(data), "%d/%m/%Y").strftime("%Y-%m-%d")
                    except ValueError:
                        exibir_log(f"Data inválida para matrícula {matricula}. Ignorando.")
                        continue

                    cpf, _ = buscar_dados_funcionario(matricula)
                    if not cpf:
                        exibir_log(f"CPF ausente para matrícula {matricula}. Ignorando linha.")
                        continue

                    linha_folga = (
                        f"{str(matricula).ljust(16)}"
                        f"{str(pis).ljust(23)}"
                        f"{str(data).ljust(12)}"
                        f"{str(cpf).ljust(11)}\n"
                    )
                    arquivo_folgas.write(linha_folga)
                    exibir_log(f"Folga registrada para matrícula {matricula} no dia {data}.")

    for index, row in df.iterrows():
        matricula = row.get('Matricula')
        descricao_justificativa = row.get('Justificativa')
        data = row.get('Data')
        qtd_horas = row.get('Qtd Horas')

        if not descricao_justificativa or pd.isna(descricao_justificativa) or descricao_justificativa == "Folga":
            continue

        exibir_log(f"QtdHoras original: {qtd_horas}")

        try:
            if isinstance(data, datetime):
                data_formatada = data.strftime("%Y-%m-%d")
            else:
                data_formatada = datetime.strptime(str(data), "%d/%m/%Y").strftime("%Y-%m-%d")
        except ValueError:
            exibir_log(f"Data inválida para matrícula {matricula}. Ignorando.")
            continue

        # Validar e formatar QtdHoras
        qtd_horas = validar_horas(qtd_horas)
        exibir_log(f"QtdHoras após validação: {qtd_horas}")

        _, id_funcionario = buscar_dados_funcionario(matricula)
        if not id_funcionario:
            exibir_log(f"ID do funcionário ausente para matrícula {matricula}. Ignorando linha.")
            continue

        id_justificativa = ids_justificativas.get(descricao_justificativa)
        if id_justificativa is None:
            exibir_log(f"Justificativa '{descricao_justificativa}' não encontrada. Ignorando.")
            continue

        enviar_justificativa(id_funcionario, id_justificativa, descricao_justificativa, data_formatada, qtd_horas)


    exibir_log("Processamento de justificativas concluído.")
    messagebox.showinfo("Sucesso", "Processamento do arquivo concluído.")

def enviar_dados_combinados():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de dados",
        filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
    )
    if not caminho_arquivo:
        exibir_log(f"Nenhum arquivo selecionado. Operação de tratamento a {dados_selecionados['Razão Social']} abortada.")
        return

    try:
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        exibir_log(f"Erro ao carregar o arquivo: {e}")
        return

    try:
        exibir_log(f"Iniciando envio de marcações para {dados_selecionados['Razão Social']} ")
        processar_marcacoes(df)
    except Exception as e:
        exibir_log(f"Erro ao enviar marcações: {e}")

    try:
        exibir_log(f"Iniciando envio de justificativas para {dados_selecionados['Razão Social']}")
        processar_arquivo_excel(df)
    except Exception as e:
        exibir_log(f"Erro ao processar justificativas: {e}")

    exibir_log(f"Tratamento da empresa {dados_selecionados['Razão Social']} concluido")

root = ttk.Window(themename="darkly")
root.title("Mantis")
root.iconbitmap("M-Kairos.ico")

razao_social_var = tk.StringVar()
cnpj_var = tk.StringVar()
chave_var = tk.StringVar()
cpf_var = tk.StringVar()
data_inicio_var = tk.StringVar()
data_fim_var = tk.StringVar()

frame_selecao = ttk.Frame(root, padding=10)
frame_selecao.grid(row=0, column=0, pady=10, padx=10, sticky="nsew")
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

frame_campos = ttk.Frame(frame_selecao)
frame_campos.grid(row=0, column=0, pady=10, padx=10)

# Adicionando a Label e a Combobox para Razão Social
ttk.Label(frame_campos, text="Selecione a Razão Social:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
combo_razao_social = ttk.Combobox(frame_campos, textvariable=razao_social_var, width=40)
combo_razao_social.grid(row=0, column=1, padx=5, pady=5)

# Tornar o dropdown visível ao clicar em qualquer parte da combobox
combo_razao_social.bind("<Button-1>", lambda e: combo_razao_social.event_generate("<Down>"))

# Vincular o evento de seleção
combo_razao_social.bind("<<ComboboxSelected>>", lambda e: preencher_detalhes())

# Adicionando a Label e a Entry para CNPJ
ttk.Label(frame_campos, text="CNPJ:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
ttk.Entry(frame_campos, textvariable=cnpj_var, state='readonly', width=42).grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")

# Botão para selecionar arquivo à direita do CNPJ
ttk.Button(frame_campos, text="Selecionar Arquivo", command=selecionar_arquivo_empresas, width=20, bootstyle="light").grid(row=0, column=3, padx=5, pady=5)

# Adicionando a Label e a Entry para Chave API
ttk.Label(frame_campos, text="Chave API:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
ttk.Entry(frame_campos, textvariable=chave_var, state='readonly', width=42).grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky="w")

# Adicionando a Label e a Entry para CPF Responsável
ttk.Label(frame_campos, text="CPF Responsável:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
ttk.Entry(frame_campos, textvariable=cpf_var, state='readonly', width=42).grid(row=3, column=1, columnspan=3, padx=5, pady=5, sticky="w")

ttk.Label(frame_campos, text="Data Início:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
entry_data_inicio = ttk.Entry(frame_campos, textvariable=data_inicio_var, width=10)
entry_data_inicio.grid(row=4, column=1, padx=5, pady=5, sticky="w")
entry_data_inicio.bind("<KeyRelease>", lambda event: formatar_data(entry_data_inicio, entry_data_inicio.get()))

ttk.Label(frame_campos, text="Data Fim:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
entry_data_fim = ttk.Entry(frame_campos, textvariable=data_fim_var, width=10)
entry_data_fim.grid(row=5, column=1, padx=5, pady=5, sticky="w")
entry_data_fim.bind("<KeyRelease>", lambda event: formatar_data(entry_data_fim, entry_data_fim.get()))

ttk.Button(frame_campos, text="Gravar Info", command=confirmar_selecao, width=20, bootstyle="light").grid(row=5, column=3, padx=10)

botao_frame = ttk.Frame(frame_selecao)
botao_frame.grid(row=6, column=0, pady=15, sticky="ew")

# Ajustando o peso das colunas para distribuir igualmente o espaço
for i in range(4):  # Ajustando peso para cada botão (número de colunas deve coincidir com o número de botões)
    botao_frame.grid_columnconfigure(i, weight=1)

# Ajustando largura dos botões
botao_width = 25  # Aumente conforme necessário para acomodar os textos maiores

ttk.Button(botao_frame, text="Tratamento", command=enviar_dados_combinados, width=botao_width, bootstyle="success").grid(row=0, column=0, padx=10, pady=5)
ttk.Button(botao_frame, text="Pessoas", command=abrir_popup_selecao_pessoas, width=botao_width, bootstyle="success").grid(row=0, column=1, padx=10, pady=5)
ttk.Button(botao_frame, text="Coleta de Planilhas", command=abrir_popup_selecao_coleta, width=botao_width, bootstyle="success").grid(row=0, column=2, padx=10, pady=5)
ttk.Button(botao_frame, text="Selecionar Justificativas", command=funcao_justificativa_get, width=botao_width, bootstyle="success").grid(row=0, column=3, padx=10, pady=5)

# Log de mensagens com barra de rolagem
frame_log = ttk.Frame(frame_selecao, padding=5)
frame_log.grid(row=7, column=0, pady=10, padx=5, sticky="nsew")

log_widget = tk.Text(frame_log, height=10, wrap="word", state="disabled", relief="solid", borderwidth=1)
log_widget.grid(row=0, column=0, sticky="nsew")

scrollbar = ttk.Scrollbar(frame_log, orient="vertical", command=log_widget.yview)
scrollbar.grid(row=0, column=1, sticky="ns")
log_widget.configure(yscrollcommand=scrollbar.set)

frame_log.grid_columnconfigure(0, weight=1)
frame_log.grid_rowconfigure(0, weight=1)

root.mainloop()
