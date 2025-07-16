import pandas as pd
import numpy as np
import customtkinter as ctk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import re
import os
import tempfile
import seaborn as sns
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class LeadAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Análise de Leads")
        self.root.geometry("1400x900")
        
        self.data = None
        self.current_city = "Total"
        self.all_periods = []
        self.selected_origins = []
        
        # Criar layout principal
        self.main_frame = ctk.CTkFrame(root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Painel de controle
        self.control_panel = ctk.CTkFrame(self.main_frame, width=250)
        self.control_panel.pack(side="left", fill="y", padx=(0, 10))
        
        # Botão para carregar arquivo
        self.load_button = ctk.CTkButton(self.control_panel, text="Carregar Planilha", command=self.load_file)
        self.load_button.pack(pady=10, padx=10, fill="x")
        
        # Seletor de cidade
        self.city_label = ctk.CTkLabel(self.control_panel, text="Selecione a Cidade:")
        self.city_label.pack(pady=(1, 1))
        
        self.city_var = ctk.StringVar(value="Total")
        self.city_selector = ctk.CTkComboBox(
            self.control_panel, 
            values=["Total"],
            variable=self.city_var
        )
        self.city_selector.pack(pady=5, padx=10, fill="x")
        
        # Adicionar seletores de período inicial e final
        self.period_label_start = ctk.CTkLabel(self.control_panel, text="Intervalo de Período:")
        self.period_label_start.pack(pady=(1, 1))
        
        self.period_var_start = ctk.StringVar(value="")
        self.period_selector_start = ctk.CTkComboBox(
            self.control_panel, 
            values=[],
            variable=self.period_var_start
        )
        self.period_selector_start.pack(pady=5, padx=10, fill="x")
        
        self.period_var_end = ctk.StringVar(value="")
        self.period_selector_end = ctk.CTkComboBox(
            self.control_panel, 
            values=[],
            variable=self.period_var_end
        )
        self.period_selector_end.pack(pady=5, padx=10, fill="x")
        
        # Filtro de origem - Layout melhorado
        self.origin_label = ctk.CTkLabel(self.control_panel, text="Filtrar por Origem:")
        self.origin_label.pack(pady=(1, 1))
        
        # Frame para controles de origem
        self.origin_control_frame = ctk.CTkFrame(self.control_panel)
        self.origin_control_frame.pack(fill="x", padx=10, pady=(0, 5))
        
        # Botão Selecionar Todas
        self.select_all_button = ctk.CTkButton(
            self.origin_control_frame, 
            text="Todas",
            width=60,
            command=self.select_all_origins
        )
        self.select_all_button.pack(side="left", padx=(0, 5))
        
        # Botão Limpar Seleção
        self.clear_button = ctk.CTkButton(
            self.origin_control_frame, 
            text="Nenhuma",
            width=60,
            command=self.clear_all_origins
        )
        self.clear_button.pack(side="left", padx=(0, 5))
        
        # Campo de pesquisa
        self.search_origin = ctk.CTkEntry(
            self.origin_control_frame,
            placeholder_text="Pesquisar..."
        )
        self.search_origin.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.search_origin.bind("<KeyRelease>", self.filter_origin_list)
        
        # Frame rolável para checkboxes
        self.origin_frame = ctk.CTkScrollableFrame(self.control_panel, height=150)
        self.origin_frame.pack(fill="x", padx=10, pady=5)
        self.origin_checkboxes = {}

        # Botão para aplicar filtros
        self.apply_button = ctk.CTkButton(
            self.control_panel,
            text="Aplicar Filtros",
            command=self.apply_filters
        )
        self.apply_button.pack(pady=10, padx=10, fill="x")
        
        # Seletor de visualização
        self.view_label = ctk.CTkLabel(self.control_panel, text="Tipo de Visualização:")
        self.view_label.pack(pady=(1, 1))
        
        self.view_var = ctk.StringVar(value="Visão Geral")
        self.view_selector = ctk.CTkComboBox(
            self.control_panel, 
            values=[
                "Visão Geral",
                "Desempenho por Origem",
                "Conversão por Canal",
                "Evolução Mensal",
                "Top Canais",
                "Eficiência de Vendas",
                "Correlação Leads-Vendas",
                "Dispersão Leads x Vendas"
            ],
            variable=self.view_var,
            command=self.update_dashboard
        )
        self.view_selector.pack(pady=5, padx=10, fill="x")
        
        # Botão de exportar
        self.export_button = ctk.CTkButton(self.control_panel, text="Exportar Relatório PDF", command=self.export_report)
        self.export_button.pack(pady=20, padx=10, fill="x")
        
        # Área de dashboard
        self.dashboard_frame = ctk.CTkFrame(self.main_frame)
        self.dashboard_frame.pack(side="right", fill="both", expand=True)
        
        # Notebook para múltiplas visualizações
        self.notebook = ttk.Notebook(self.dashboard_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # Abas do notebook
        self.tab1 = ctk.CTkFrame(self.notebook)
        self.tab2 = ctk.CTkFrame(self.notebook)
        
        self.notebook.add(self.tab1, text="Dashboard")
        self.notebook.add(self.tab2, text="Dados Detalhados")
        
        # Inicializar widgets vazios
        self.init_dashboard()
        
        self.period_mapping = {}  # Mapeamento período formatado -> original
        self.origin_vars = {}  # {origem: BooleanVar}
        self.origin_checkboxes = {}  # {origem: widget}
        self.all_origins = []  # Lista completa de origens disponíveis
        self.current_origins = []  # Lista de origens atuais

    def init_dashboard(self):
        # Widgets para a primeira aba
        self.summary_frame = ctk.CTkScrollableFrame(self.tab1)
        self.summary_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Widgets para a segunda aba
        self.data_frame = ctk.CTkFrame(self.tab2)
        self.data_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Placeholders
        self.summary_label = ctk.CTkLabel(self.summary_frame, text="Carregue uma planilha para começar a análise")
        self.summary_label.pack(pady=50)
        
        self.data_label = ctk.CTkLabel(self.data_frame, text="Os dados detalhados serão exibidos aqui")
        self.data_label.pack(pady=50)
    
    def load_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
            
        try:
            # Carregar todas as abas do Excel
            self.data = {}
            abas_excluidas = ["Salvador", "Planilha1", "Fortaleza", "Deliverysalvador1", 
                             "Deliverysalvador4", "SSA", "Brasília", "J4ASSUNCAO"]
            
            with pd.ExcelFile(file_path) as xls:
                for sheet_name in xls.sheet_names:
                    if sheet_name in abas_excluidas:
                        continue
                        
                    df = pd.read_excel(xls, sheet_name)
                    
                    # Padronizar nomes de colunas
                    df.columns = [self.clean_column_name(col) for col in df.columns]
                    
                    # Converter colunas numéricas
                    df = self.convert_numeric_columns(df)
                    
                    # Converter porcentagens
                    for col in df.columns:
                        if 'conversao' in col.lower() or '%' in col.lower():
                            df[col] = self.convert_percentage(df[col])
                    
                    # Adicionar coluna de cidade
                    df['cidade'] = sheet_name
                    
                    self.data[sheet_name] = df
            
            # Coletar e processar todos os períodos únicos
            unique_period_dates = set()
            unique_period_strings = {}
            
            for sheet_name, df in self.data.items():
                # Identificar coluna de período
                periodo_col = None
                for col in df.columns:
                    if 'periodo' in col.lower() or 'período' in col.lower():
                        periodo_col = col
                        break
                
                if periodo_col:
                    # Processar cada valor de período
                    for p in df[periodo_col].dropna().unique():
                        # Converter para objeto datetime
                        dt = self.parse_period(p)
                        display_value = self.format_period_display(dt)
                        unique_period_dates.add(dt)
                        unique_period_strings[dt] = p
            
            # Ordenar os períodos e converter para formato de exibição
            sorted_periods = sorted(unique_period_dates)
            self.all_periods = [str(unique_period_strings[dt]) for dt in sorted_periods]
            self.display_periods = [self.format_period_display(dt) for dt in sorted_periods]  # Lista formatada

            self.period_mapping = {}
            for i, dt in enumerate(sorted_periods):
                display_val = self.display_periods[i]
                original_val = self.all_periods[i]
                self.period_mapping[display_val] = original_val

            # Atualize os comboboxes com os valores formatados
            self.period_selector_start.configure(values=self.display_periods)
            self.period_selector_end.configure(values=self.display_periods)
            if self.display_periods:
                self.period_var_start.set(self.display_periods[0])
                self.period_var_end.set(self.display_periods[-1])
            
            # Atualizar seletor de cidades
            cities = ["Total"] + list(self.data.keys())
            self.city_selector.configure(values=cities)
            
            # Processar dados
            self.process_data()
            self.update_origin_checklist()
            self.update_dashboard()
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("Erro", f"Falha ao carregar arquivo:\n{str(e)}\n\nDetalhes:\n{error_details}")
    
    def clean_column_name(self, name):
        """Padroniza nomes de colunas"""
        name = str(name).strip().lower()
        
        # Preservar o nome da coluna de período
        if 'período' in name or 'periodo' in name:
            return 'periodo'
            
        # Substituir caracteres especiais, mantendo números
        name = re.sub(r'[^\w\s]', '_', name)
        name = re.sub(r'\s+', '_', name)
        return name
    
    def apply_filters(self):
        """Aplica os filtros sem recriar a lista de origens"""
        self.update_dashboard()

    def on_filter_change(self, event=None):
        """Atualiza a lista de origens quando os filtros mudam"""
        self.update_origin_checklist()
        self.update_dashboard()

    def convert_numeric_columns(self, df):
        """Converte colunas numéricas conhecidas para tipo float"""
        numeric_columns = ['contatos', 'aproveitados', 'vendas', 'leads', 'conversao']
        for col in numeric_columns:
            if col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                except:
                    pass
        return df
    
    def convert_percentage(self, series):
        """Converte porcentagens para valores numéricos"""
        if series.dtype == 'object':
            # Tentar converter strings com % para float
            series = series.astype(str).str.replace('%', '').str.replace(',', '.').str.strip()
            series = pd.to_numeric(series, errors='coerce') / 100
        return series
    
    def parse_period(self, period_str):
        """Converte período em formato datetime para ordenação"""
        try:
            # Se já for um objeto datetime/Timestamp
            if isinstance(period_str, (datetime, pd.Timestamp)):
                return period_str
            
            # Converter para string
            period_str = str(period_str).strip()
            
            # Tentar diferentes formatos de data
            formats = [
                "%Y-%m-%d",    # 2023-07-01
                "%Y-%m-%d %H:%M:%S",  # 2023-07-01 00:00:00
                "%Y-%m",        # 2023-07
                "%b-%y",        # Jul-23
                "%b %y",        # Jul 23
                "%B %Y",        # July 2023
                "%m/%d/%Y",     # 07/01/2023
                "%d/%m/%Y",     # 01/07/2023
                "%d-%m-%Y"      # 01-07-2023
            ]
            
            for fmt in formats:
                try:
                    return datetime.strptime(period_str, fmt)
                except ValueError:
                    continue
            
            # Mapeamento de meses em português
            month_map_full = {
                'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4, 
                'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 
                'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
            }
            
            month_map_abbr = {
                'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
                'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
            }
            
            # Verificar se é apenas o nome do mês
            if period_str.lower() in month_map_full:
                year = datetime.now().year
                return datetime(year, month_map_full[period_str.lower()], 1)
            
            # Tentar padrões como "Jan-23" ou "Janeiro 2023"
            match = re.match(r'(\D{3,}).*?(\d{2,4})', period_str, re.IGNORECASE)
            if match:
                month_str = match.group(1).lower()
                year_str = match.group(2)
                
                # Verificar mês completo
                if month_str in month_map_full:
                    month = month_map_full[month_str]
                # Verificar abreviação
                elif month_str[:3] in month_map_abbr:
                    month = month_map_abbr[month_str[:3]]
                else:
                    return datetime(1900, 1, 1)
                    
                year = int(year_str) if len(year_str) == 4 else 2000 + int(year_str)
                return datetime(year, month, 1)
            
            # Tentar padrão de apenas números (MM/AAAA)
            parts = re.findall(r'\d+', period_str)
            if len(parts) >= 2:
                month = int(parts[0])
                year = int(parts[1]) if len(parts[1]) == 4 else 2000 + int(parts[1])
                if 1 <= month <= 12:
                    return datetime(year, month, 1)
            
            return datetime(1900, 1, 1)  # Data padrão para ordenação
            
        except Exception as e:
            print(f"Erro ao analisar período: {period_str}, erro: {str(e)}")
            return datetime(1900, 1, 1)
    
    def format_period_display(self, dt):
        """Formata período para exibição no formato 'jan/25'"""
        month_abbr = {
            1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 
            5: 'mai', 6: 'jun', 7: 'jul', 8: 'ago', 
            9: 'set', 10: 'out', 11: 'nov', 12: 'dez'
        }
        return f"{month_abbr[dt.month]}/{dt.strftime('%y')}"
    
    def process_data(self):
        """Processa dados brutos e prepara para análise"""
        # Criar DataFrame consolidado
        self.consolidated_data = pd.concat(self.data.values(), ignore_index=True)
        
        # Identificar coluna de período
        periodo_col = None
        for col in self.consolidated_data.columns:
            if 'periodo' in col.lower():
                periodo_col = col
                break
                
        if periodo_col:
            self.consolidated_data['periodo_dt'] = self.consolidated_data[periodo_col].apply(self.parse_period)
            # Renomear para o nome padrão
            self.consolidated_data.rename(columns={periodo_col: 'periodo'}, inplace=True)
    
    def select_all_origins(self):
        """Marca todas as origens"""
        for origin in self.all_origins:
            self.origin_vars[origin].set(True)

    def clear_all_origins(self):
        """Desmarca todas as origens"""
        for origin in self.all_origins:
            self.origin_vars[origin].set(False)
    
    def filter_origin_list(self, event=None):
        """Filtra a lista de origens visíveis"""
        self.refresh_origin_list()

    def update_origin_checklist(self):
        """Atualiza a lista de origens mantendo o estado das seleções"""
        df = self.get_current_data()
        
        if df is None or 'origem' not in df.columns:
            self.all_origins = []
            self.refresh_origin_list()
            return
            
        # Obter novas origens únicas
        new_origins = sorted(df['origem'].astype(str).unique().tolist())
        new_origins = [o for o in new_origins if o.lower() not in ['total', 'geral', 'consolidado']]
        
        # Verificar se houve mudança na lista de origens
        if set(new_origins) == set(self.all_origins):
            return
            
        self.all_origins = new_origins
        
        # Criar novas variáveis para novas origens
        for origin in self.all_origins:
            if origin not in self.origin_vars:
                self.origin_vars[origin] = ctk.BooleanVar(value=True)
        
        # Remover variáveis de origens que não existem mais
        for origin in list(self.origin_vars.keys()):
            if origin not in self.all_origins:
                del self.origin_vars[origin]
        
        self.refresh_origin_list()
    
    def get_selected_origins(self):
        """Retorna as origens selecionadas"""
        return [origin for origin in self.all_origins if self.origin_vars[origin].get()]

    def refresh_origin_list(self):
        """Atualiza a exibição das checkboxes"""
        # Limpar frame atual
        for widget in self.origin_frame.winfo_children():
            widget.destroy()
        
        self.origin_checkboxes = {}
        
        search_term = self.search_origin.get().strip().lower()
        
        for origin in self.all_origins:
            # Filtrar por termo de pesquisa
            if search_term and search_term not in origin.lower():
                continue
                
            var = self.origin_vars[origin]
            cb = ctk.CTkCheckBox(
                self.origin_frame,
                text=origin,
                variable=var
            )
            cb.pack(anchor="w", padx=5, pady=2)
            self.origin_checkboxes[origin] = cb

    def get_current_data(self):
        """Obtém dados filtrados para seleção atual"""
        if self.data is None:
            return None

        city = self.city_var.get()
        period_start = self.period_mapping.get(self.period_var_start.get(), "")
        period_end = self.period_mapping.get(self.period_var_end.get(), "")

        if city in self.data:
            df = self.data[city].copy()
        else:
            return pd.DataFrame()

        # Filtrar por intervalo de períodos
        if period_start and period_end and 'periodo' in df.columns:
            # Converter períodos selecionados para datetime
            start_dt = self.parse_period(period_start)
            end_dt = self.parse_period(period_end)

            # Criar coluna de data se necessário
            if 'periodo_dt' not in df.columns:
                df['periodo_dt'] = df['periodo'].apply(self.parse_period)

            # Filtrar por intervalo
            df = df[(df['periodo_dt'] >= start_dt) & (df['periodo_dt'] <= end_dt)]

        # Filtrar por origem
        selected_origins = self.get_selected_origins()
        if selected_origins and 'origem' in df.columns:
            df = df[df['origem'].isin(selected_origins)]

        # Filtrar por origem (sempre excluindo totais)
        if 'origem' in df.columns:
            # Filtrar origens selecionadas
            selected_origins = self.get_selected_origins()
            if selected_origins:
                df = df[df['origem'].isin(selected_origins)]
            
            # Sempre excluir origens de consolidação
            origens = df['origem'].astype(str).str.lower().str.strip()
            mask = ~(
                origens.str.contains('total') |
                origens.str.contains('geral') |
                origens.str.contains('consolidado')
            )
            df = df[mask]
        
        return df
    
    def export_origin_performance_excel(self, df):
        """Exporta a tabela de desempenho por origem para Excel"""
        import pandas as pd
        from tkinter import filedialog, messagebox

        table_data = self.get_origin_performance_data(df)
        if not table_data or len(table_data) < 2:
            messagebox.showwarning("Exportar", "Nenhum dado disponível para exportação")
            return

        # Converter para DataFrame
        headers = table_data[0]
        rows = table_data[1:]
        df_export = pd.DataFrame(rows, columns=headers)

        # Solicitar local para salvar
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            return

        try:
            df_export.to_excel(file_path, index=False)
            messagebox.showinfo("Exportar", f"Tabela exportada com sucesso!\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar tabela:\n{str(e)}")

    def update_dashboard(self, event=None):
        if self.data is None:
            return
            
        # Atualizar lista de origens
        # self.update_origin_checklist()
        
        current_view = self.view_var.get()
        
        # Limpar frames anteriores
        for widget in self.summary_frame.winfo_children():
            widget.destroy()
        for widget in self.data_frame.winfo_children():
            widget.destroy()
        
        # Obter dados atuais
        df = self.get_current_data()
        if df is None or df.empty:
            ctk.CTkLabel(self.summary_frame, text="Nenhum dado disponível para a seleção atual").pack(pady=50)
            return
        
        # Gerar visualização selecionada
        if current_view == "Visão Geral":
            self.show_summary(df)
        elif current_view == "Desempenho por Origem":
            self.show_origin_performance(df)
        elif current_view == "Conversão por Canal":
            self.show_conversion_by_channel(df)
        elif current_view == "Evolução Mensal":
            self.show_monthly_trend(df)
        elif current_view == "Top Canais":
            self.show_top_channels(df)
        elif current_view == "Eficiência de Vendas":
            self.show_sales_efficiency(df)
        elif current_view == "Correlação Leads-Vendas":
            self.show_correlation(df)
        elif current_view == "Dispersão Leads x Vendas":
            self.show_scatter_plots(df)
        
        # Mostrar dados detalhados na segunda aba
        self.show_detailed_data(df)
    
    def show_summary(self, df):
        """Exibe métricas de resumo principais"""
        frame = self.summary_frame
        
        # Verificar se as colunas necessárias existem
        required_columns = ['contatos', 'aproveitados', 'vendas']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            ctk.CTkLabel(frame, text=f"Colunas obrigatórias ausentes: {', '.join(missing_columns)}").pack()
            return
        
        # Criar cards com métricas principais
        metrics_frame = ctk.CTkFrame(frame)
        metrics_frame.pack(fill="x", pady=10, padx=10)
        
        # Calcular métricas
        total_contacts = df['contatos'].sum()
        total_leads = df['aproveitados'].sum()
        total_sales = df['vendas'].sum()
        
        conversion_rate = total_sales / total_contacts if total_contacts > 0 else 0
        lead_conversion_rate = total_sales / total_leads if total_leads > 0 else 0
        lead_per_sale = total_contacts / total_sales if total_sales > 0 else 0
        
        # Métricas para mostrar
        metrics = [
            ("Total de Contatos", total_contacts, "{:,.0f}"),
            ("Leads Aproveitados", total_leads, "{:,.0f}"),
            ("Vendas Fechadas", total_sales, "{:,.0f}"),
            ("Taxa de Conversão", conversion_rate, "{:.1%}"),
            ("Conversão de Leads", lead_conversion_rate, "{:.1%}"),
            ("Lead por Venda", lead_per_sale, "{:,.1f}")
        ]
        
        # Exibir cards de métricas
        for i, (title, value, fmt) in enumerate(metrics):
            card = ctk.CTkFrame(metrics_frame, width=180, height=100)
            card.grid(row=0, column=i, padx=10, pady=10)
            
            ctk.CTkLabel(card, text=title, font=("Helvetica", 12, "bold")).pack(pady=(10, 5))
            ctk.CTkLabel(card, text=fmt.format(value), font=("Arial", 14)).pack(pady=(0, 10))
        
        # Gráficos
        charts_frame = ctk.CTkFrame(frame)
        charts_frame.pack(fill="both", expand=True, pady=10)
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        # Gráfico 1: Distribuição de origens (Top 5 + Outros)
        if 'origem' in df.columns and 'contatos' in df.columns:
            origin_counts = df.groupby('origem')['contatos'].sum()
            
            # Agrupar origens menores em "Outros"
            if len(origin_counts) > 5:
                top_origins = origin_counts.nlargest(5)
                other = origin_counts.sum() - top_origins.sum()
                top_origins = pd.concat([top_origins, pd.Series({'Outros': other})])
            else:
                top_origins = origin_counts
            
            # MODIFICAÇÃO 1: Remover o rótulo do eixo Y
            top_origins.plot.pie(autopct='%1.1f%%', ax=ax1, startangle=90, ylabel='')
            ax1.set_title("Distribuição por Origem (Top 5)")
            # Remover completamente o label do eixo Y
            ax1.set_ylabel('')
        else:
            ax1.set_title("Dados de origem não disponíveis")
        
        # Gráfico 2: Top 5 Canais por Conversão Total
        if 'origem' in df.columns and 'contatos' in df.columns and 'vendas' in df.columns:
            # Calcular eficiência dos canais
            channel_efficiency = df.groupby('origem').agg({
                'contatos': 'sum',
                'vendas': 'sum'
            })
            
            # Calcular taxa de conversão
            channel_efficiency['taxa_conversao'] = channel_efficiency['vendas'] / channel_efficiency['contatos']
            
            # Filtrar canais com volume significativo
            channel_efficiency = channel_efficiency[channel_efficiency['contatos'] >= 50]
            
            if not channel_efficiency.empty:
                # Pegar top 5
                top_conversion = channel_efficiency.nlargest(5, 'taxa_conversao')
                top_conversion = top_conversion.sort_values('taxa_conversao', ascending=True)
                
                # Criar gráfico de barras
                bars = ax2.barh(top_conversion.index, top_conversion['taxa_conversao'], color='skyblue')
                ax2.set_title("Top 5 Canais - Conversão Total")
                ax2.set_xlabel("Taxa de Conversão")
                ax2.grid(axis='x', linestyle='--', alpha=0.7)
                
                # Formatar eixo X como porcentagem
                ax2.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0%}'))
                
                # Adicionar valores nas barras
                for bar in bars:
                    width = bar.get_width()
                    ax2.text(width + 0.01, bar.get_y() + bar.get_height()/2, 
                            f'{width:.1%}', ha='left', va='center')
            else:
                ax2.set_title("Sem canais com dados suficientes")
        else:
            ax2.set_title("Dados incompletos")
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=charts_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def show_origin_performance(self, df):
        """Mostra desempenho por origem de leads"""
        frame = self.summary_frame

        # Usar a mesma preparação de dados do PDF
        table_data = self.get_origin_performance_data(df)
        if not table_data:
            ctk.CTkLabel(frame, text="Dados de origem não disponíveis").pack()
            return

        # Botão de exportação para Excel
        export_btn = ctk.CTkButton(
            frame, text="Exportar Excel",
            command=lambda: self.export_origin_performance_excel(df)
        )
        export_btn.pack(pady=(0, 10))

        # Criar frame para tabela
        table_frame = ctk.CTkScrollableFrame(frame)
        table_frame.pack(fill="both", expand=True, pady=10, padx=10)

        # Cabeçalho
        headers = table_data[0]
        for j, header in enumerate(headers):
            label = ctk.CTkLabel(table_frame, text=header, font=("Arial", 10, "bold"))
            label.grid(row=0, column=j, padx=2, pady=2, sticky="nsew")

        # Dados
        for i, row in enumerate(table_data[1:]):
            for j, value in enumerate(row):
                ctk.CTkLabel(table_frame, text=value).grid(
                    row=i+1, column=j, padx=2, pady=2, sticky="nsew")

        # Tornar colunas expansíveis
        for j in range(len(headers)):
            table_frame.grid_columnconfigure(j, weight=1)
    
    def show_conversion_by_channel(self, df):
        """Mostra análise de conversão por canal"""
        frame = self.summary_frame
        
        if 'origem' not in df.columns:
            ctk.CTkLabel(frame, text="Dados de origem não disponíveis").pack()
            return
        
        # Criar gráfico de comparação de conversões
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        # Agrupar por origem
        grouped = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum',
            'aproveitados': 'sum'
        }).reset_index()
        
        # Calcular taxas
        grouped['taxa_conversao'] = grouped['vendas'] / grouped['contatos']
        grouped['conversao_ap'] = grouped['vendas'] / grouped['aproveitados']
        
        # Filtrar origens com pelo menos 100 contatos
        grouped = grouped[grouped['contatos'] >= 100]
        
        # Gráfico 1: Conversão total
        if not grouped.empty:
            grouped_sorted = grouped.sort_values('taxa_conversao', ascending=False)
            bars1 = ax1.bar(grouped_sorted['origem'], grouped_sorted['taxa_conversao'], color='skyblue')
            ax1.set_title("Conversão Total por Origem")
            ax1.set_ylabel("Taxa de Conversão")
            ax1.tick_params(axis='x', rotation=90, labelsize=8)
            ax1.grid(axis='y', linestyle='--', alpha=0.7)
            
            # Formatar eixo Y como porcentagem
            ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.0%}'))
            
            # Adicionar valores nas barras como porcentagem
            for bar in bars1:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height + 0.01,
                        f'{height:.1%}', ha='center', va='bottom', fontsize=6)
            
            # Gráfico 2: Conversão de leads
            grouped_sorted = grouped.sort_values('conversao_ap', ascending=False)
            bars2 = ax2.bar(grouped_sorted['origem'], grouped_sorted['conversao_ap'], color='lightgreen')
            ax2.set_title("Conversão de Leads por Origem")
            ax2.set_ylabel("Taxa de Conversão")
            ax2.tick_params(axis='x', rotation=90, labelsize=8)
            ax2.grid(axis='y', linestyle='--', alpha=0.7)
            
            # Formatar eixo Y como porcentagem
            ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.0%}'))
            
            # Adicionar valores nas barras como porcentagem
            for bar in bars2:
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height + 0.01,
                        f'{height:.1%}', ha='center', va='bottom', fontsize=6)
        
        plt.tight_layout()
        
        chart_frame = ctk.CTkFrame(frame)
        chart_frame.pack(fill="both", expand=True, pady=10)
        
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def show_monthly_trend(self, df):
        """Mostra evolução mensal das métricas principais"""
        frame = self.summary_frame
        
        if 'periodo_dt' not in df.columns:
            ctk.CTkLabel(frame, text="Dados temporais não disponíveis").pack()
            return
        
        # Agrupar por período
        monthly = df.groupby('periodo_dt').agg({
            'contatos': 'sum',
            'aproveitados': 'sum',
            'vendas': 'sum'
        }).sort_index()
        
        # Calcular métricas
        monthly['taxa_conversao'] = monthly['vendas'] / monthly['contatos']
        monthly['conversao_ap'] = monthly['vendas'] / monthly['aproveitados']
        
        # Formatar datas para exibição
        monthly['periodo_formatado'] = monthly.index.map(self.format_period_display)
        
        # Criar gráfico
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
        
        # Gráfico 1: Volume
        monthly[['contatos', 'aproveitados', 'vendas']].plot(ax=ax1, marker='o')
        ax1.set_title("Evolução Mensal - Volume")
        ax1.set_ylabel("Quantidade")
        ax1.grid(True, linestyle='--', alpha=0.7)
        ax1.legend(['Contatos', 'Aproveitados', 'Vendas'])
        ax1.set_xlabel('')
        
        # Gráfico 2: Taxas
        monthly[['taxa_conversao', 'conversao_ap']].plot(ax=ax2, marker='o')
        ax2.set_title("Evolução Mensal - Taxas de Conversão")
        ax2.set_ylabel("Taxa")
        ax2.grid(True, linestyle='--', alpha=0.7)
        ax2.legend(['Conversão Total', 'Conversão de Aproveitados'])
        ax2.set_xlabel('')

        # Formatar eixos Y como porcentagem
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.0%}'))

        plt.tight_layout()
        
        chart_frame = ctk.CTkFrame(frame)
        chart_frame.pack(fill="both", expand=True, pady=10)
        
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def show_top_channels(self, df):
        """Mostra os canais mais eficientes"""
        frame = self.summary_frame
        
        if 'origem' not in df.columns:
            ctk.CTkLabel(frame, text="Dados de origem não disponíveis").pack()
            return
        
        # Calcular eficiência dos canais
        channel_efficiency = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum',
            'aproveitados': 'sum'
        }).reset_index()
        
        # Calcular métricas
        channel_efficiency['taxa_conversao'] = channel_efficiency['vendas'] / channel_efficiency['contatos']
        channel_efficiency['conversao_ap'] = channel_efficiency['vendas'] / channel_efficiency['aproveitados']
        channel_efficiency = channel_efficiency[channel_efficiency['contatos'] >= 50]
        
        # Criar gráficos
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        # Top 5 por conversão total
        if not channel_efficiency.empty:
            top_conversion = channel_efficiency.nlargest(5, 'taxa_conversao').sort_values('taxa_conversao', ascending=True)
            bars1 = ax1.barh(top_conversion['origem'], top_conversion['taxa_conversao'], color='skyblue')
            ax1.set_title("Top 5 Canais - Conversão Total")
            ax1.set_xlabel("Taxa de Conversão")
            ax1.grid(axis='x', linestyle='--', alpha=0.7)
            
            # Formatar eixo X como porcentagem
            ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0%}'))
            
            # Adicionar valores nas barras
            for bar in bars1:
                width = bar.get_width()
                ax1.text(width + 0.01, bar.get_y() + bar.get_height()/2, 
                        f'{width:.1%}', ha='left', va='center')
            
            # Top 5 por conversão de leads
            top_lead_conversion = channel_efficiency.nlargest(5, 'conversao_ap').sort_values('conversao_ap', ascending=True)
            bars2 = ax2.barh(top_lead_conversion['origem'], top_lead_conversion['conversao_ap'], color='lightgreen')
            ax2.set_title("Top 5 Canais - Conversão de Leads")
            ax2.set_xlabel("Taxa de Conversão")
            ax2.grid(axis='x', linestyle='--', alpha=0.7)
            
            # Formatar eixo X como porcentagem
            ax2.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0%}'))
            
            # Adicionar valores nas barras
            for bar in bars2:
                width = bar.get_width()
                ax2.text(width + 0.01, bar.get_y() + bar.get_height()/2, 
                        f'{width:.1%}', ha='left', va='center')
        
        plt.tight_layout()
        
        chart_frame = ctk.CTkFrame(frame)
        chart_frame.pack(fill="both", expand=True, pady=10)
        
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def show_sales_efficiency(self, df):
        """Mostra eficiência de vendas"""
        frame = self.summary_frame
        
        # Verificar se as colunas necessárias existem
        if 'contatos' not in df.columns or 'vendas' not in df.columns:
            ctk.CTkLabel(frame, text="Dados de eficiência não disponíveis").pack()
            return
        
        # Agrupar dados por origem e calcular totais
        grouped = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum'
        }).reset_index()
        
        # Calcular eficiência como porcentagem
        grouped['eficiencia'] = (grouped['vendas'] / grouped['contatos']) * 100
        
        # Filtrar origens com pelo menos 10 vendas
        filtered = grouped[grouped['vendas'] >= 10]
        
        if filtered.empty:
            ctk.CTkLabel(frame, text="Nenhuma origem com dados suficientes (mínimo 10 vendas)").pack()
            return
        
        # Selecionar top 15
        top_15 = filtered.nlargest(15, 'eficiencia').sort_values('eficiencia', ascending=True)
        
        # Criar gráfico
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Criar gráfico de barras horizontais
        bars = ax.barh(top_15['origem'], top_15['eficiencia'], color='skyblue')
        ax.set_title("Eficiência de Vendas (Vendas por Lead - Top 15)")
        ax.set_xlabel("Eficiência (%)")
        ax.grid(axis='x', linestyle='--', alpha=0.7)
        
        # Adicionar valores nas barras
        for bar in bars:
            width = bar.get_width()
            ax.text(width + 0.5, bar.get_y() + bar.get_height()/2, 
                    f'{width:.1f}%', ha='left', va='center')
        
        plt.tight_layout()
        
        chart_frame = ctk.CTkFrame(frame)
        chart_frame.pack(fill="both", expand=True, pady=10)
        
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def calculate_correlations(self, df):
        """Calcula correlação entre Leads e Vendas para cada origem"""
        correlations = []
        
        for origem in df['origem'].unique():
            df_origem = df[df['origem'] == origem]
            
            # Garantir que temos dados suficientes
            if len(df_origem) >= 3 and 'contatos' in df_origem.columns and 'vendas' in df_origem.columns:
                # Remover zeros para evitar problemas
                df_clean = df_origem[(df_origem['contatos'] > 0) & (df_origem['vendas'] > 0)]
                
                if len(df_clean) >= 3:
                    try:
                        corr = df_clean['contatos'].corr(df_clean['vendas'])
                        correlations.append({
                            'origem': origem,
                            'correlacao': corr,
                            'n_periodos': len(df_clean)
                        })
                    except:
                        continue
        
        # Criar DataFrame e ordenar
        if correlations:
            df_corr = pd.DataFrame(correlations)
            return df_corr.sort_values('correlacao', ascending=False)
        return pd.DataFrame()
    
    def show_correlation(self, df):
        """Mostra análise de correlação entre leads e vendas"""
        frame = self.summary_frame
        
        df_corr = self.calculate_correlations(df)
        
        if df_corr.empty:
            ctk.CTkLabel(frame, text="Dados insuficientes para calcular correlações (mínimo 3 períodos por origem)").pack()
            return
        
        # Separar em melhores e piores
        top_15 = df_corr.head(15).sort_values('correlacao', ascending=False)
        bottom_15 = df_corr.tail(15).sort_values('correlacao', ascending=True)
        
        # Criar gráficos
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 7))
        
        # Gráfico 1: Melhores correlações
        bars1 = ax1.barh(top_15['origem'], top_15['correlacao'], color='#2ecc71')
        ax1.set_title('TOP 15 - Melhores Correlações (Leads x Vendas)')
        ax1.set_xlabel("Coeficiente de Correlação")
        ax1.set_xlim(-1.1, 1.1)
        ax1.grid(axis='x', linestyle='--', alpha=0.7)
        
        # Adicionar valores nas barras
        for bar in bars1:
            width = bar.get_width()
            ax1.text(width + 0.02 if width > 0 else width - 0.1, 
                    bar.get_y() + bar.get_height()/2, 
                    f'{width:.2f}', ha='left' if width > 0 else 'right', va='center')
        
        # Gráfico 2: Piores correlações
        bars2 = ax2.barh(bottom_15['origem'], bottom_15['correlacao'], color='#e74c3c')
        ax2.set_title('TOP 15 - Piores Correlações (Leads x Vendas)')
        ax2.set_xlabel("Coeficiente de Correlação")
        ax2.set_xlim(-1.1, 1.1)
        ax2.grid(axis='x', linestyle='--', alpha=0.7)
        
        # Adicionar valores nas barras
        for bar in bars2:
            width = bar.get_width()
            ax2.text(width + 0.02 if width > 0 else width - 0.1, 
                    bar.get_y() + bar.get_height()/2, 
                    f'{width:.2f}', ha='left' if width > 0 else 'right', va='center')
        
        plt.tight_layout()
        
        chart_frame = ctk.CTkFrame(frame)
        chart_frame.pack(fill="both", expand=True, pady=10)
        
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def show_scatter_plots(self, df):
        """Mostra gráficos de dispersão entre leads e vendas"""
        frame = self.summary_frame
        
        required_columns = ['contatos', 'aproveitados', 'vendas', 'origem']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            ctk.CTkLabel(frame, text=f"Colunas obrigatórias ausentes: {', '.join(missing_columns)}").pack()
            return
        
        # Agregar dados por origem (somar todos os períodos)
        df_agg = df.groupby('origem', as_index=False).agg({
            'contatos': 'sum',
            'aproveitados': 'sum',
            'vendas': 'sum'
        })
        
        # Filtrar origens com dados válidos
        df_agg = df_agg[
            (df_agg['contatos'] > 0) & 
            (df_agg['vendas'] > 0) & 
            (df_agg['aproveitados'] > 0)
        ]
        
        if df_agg.empty:
            ctk.CTkLabel(frame, text="Sem dados válidos para análise").pack()
            return
        
        # Limitar a 15 canais para melhor visualização
        top_origins = df_agg.nlargest(15, 'contatos')['origem']
        df_filtered = df_agg[df_agg['origem'].isin(top_origins)]
        
        # Criar gráficos
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        
        # Gráfico 1: Leads vs Vendas
        sns.scatterplot(
            data=df_filtered,
            x='contatos',
            y='vendas',
            hue='origem',
            size='vendas',
            sizes=(50, 300),
            ax=ax1
        )
        ax1.set_title('Relação Leads vs Vendas por Canal')
        ax1.set_xlabel("Total de Leads")
        ax1.set_ylabel("Total de Vendas")
        ax1.grid(True, linestyle='--', alpha=0.7)
        
        # Adicionar linha de tendência
        try:
            sns.regplot(data=df_filtered, x='contatos', y='vendas', 
                        scatter=False, ax=ax1, color='gray', line_kws={'alpha': 0.5})
        except:
            pass
        
        # Gráfico 2: Leads Aproveitados vs Vendas
        sns.scatterplot(
            data=df_filtered,
            x='aproveitados',
            y='vendas',
            hue='origem',
            size='vendas',
            sizes=(50, 300),
            ax=ax2
        )
        ax2.set_title('Relação Leads Aproveitados vs Vendas por Canal')
        ax2.set_xlabel("Leads Aproveitados")
        ax2.set_ylabel("Total de Vendas")
        ax2.grid(True, linestyle='--', alpha=0.7)
        
        # Adicionar linha de tendência
        try:
            sns.regplot(data=df_filtered, x='aproveitados', y='vendas', 
                        scatter=False, ax=ax2, color='gray', line_kws={'alpha': 0.5})
        except:
            pass
        
        # Ajustar layout da legenda
        plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        
        plt.tight_layout()
        
        chart_frame = ctk.CTkFrame(frame)
        chart_frame.pack(fill="both", expand=True, pady=10)
        
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def show_detailed_data(self, df):
        """Mostra dados detalhados na segunda aba"""
        frame = self.data_frame
        
        # Criar Treeview para exibir dados tabulares
        tree_frame = ctk.CTkFrame(frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Criar barra de rolagem
        tree_scroll = ctk.CTkScrollbar(tree_frame)
        tree_scroll.pack(side="right", fill="y")
        
        # Criar Treeview
        tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set)
        
        # Configurar colunas
        tree["columns"] = list(df.columns)
        tree.column("#0", width=0, stretch=False)  # Coluna fantasma
        
        for col in df.columns:
            tree.column(col, anchor="w", width=100)
            tree.heading(col, text=col, anchor="w")
        
        # Adicionar dados
        for i, row in df.iterrows():
            tree.insert("", "end", values=list(row))
        
        tree.pack(fill="both", expand=True)
        tree_scroll.configure(command=tree.yview)
    
    def export_report(self):
        """Exporta relatório em PDF"""
        if self.data is None:
            messagebox.showwarning("Exportar", "Carregue os dados primeiro")
            return
            
        try:
            # Obter dados atuais
            df = self.get_current_data()
            if df is None or df.empty:
                messagebox.showwarning("Exportar", "Nenhum dado disponível para exportação")
                return
            
            # Solicitar local para salvar
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")]
            )
            if not file_path:
                return
            
            # Gerar relatório
            self.generate_pdf_report(df, file_path)
            messagebox.showinfo("Exportar", f"Relatório exportado com sucesso!\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar relatório:\n{str(e)}")

    def create_summary_figures(self, df):
        """Cria figuras para a visão geral"""
        if df.empty:
            return None
            
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        # Gráfico 1: Distribuição de origens
        if 'origem' in df.columns and 'contatos' in df.columns:
            origin_counts = df.groupby('origem')['contatos'].sum()
            
            if len(origin_counts) > 5:
                top_origins = origin_counts.nlargest(5)
                other = origin_counts.sum() - top_origins.sum()
                top_origins = pd.concat([top_origins, pd.Series({'Outros': other})])
            else:
                top_origins = origin_counts
            
            top_origins.plot.pie(autopct='%1.1f%%', ax=ax1, startangle=90, ylabel='')
            ax1.set_title("Distribuição por Origem (Top 5)")
        
        # Gráfico 2: Top 5 Canais por Conversão
        if 'origem' in df.columns and 'contatos' in df.columns and 'vendas' in df.columns:
            channel_efficiency = df.groupby('origem').agg({
                'contatos': 'sum',
                'vendas': 'sum'
            })
            channel_efficiency['taxa_conversao'] = channel_efficiency['vendas'] / channel_efficiency['contatos']
            channel_efficiency = channel_efficiency[channel_efficiency['contatos'] >= 50]
            
            if not channel_efficiency.empty:
                top_conversion = channel_efficiency.nlargest(5, 'taxa_conversao')
                top_conversion = top_conversion.sort_values('taxa_conversao', ascending=True)
                
                bars = ax2.barh(top_conversion.index, top_conversion['taxa_conversao'], color='skyblue')
                ax2.set_title("Top 5 Canais - Conversão Total")
                ax2.set_xlabel("Taxa de Conversão")
                ax2.grid(axis='x', linestyle='--', alpha=0.7)
                ax2.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0%}'))
                
                for bar in bars:
                    width = bar.get_width()
                    ax2.text(width + 0.01, bar.get_y() + bar.get_height()/2, 
                            f'{width:.1%}', ha='left', va='center')
        
        plt.tight_layout()
        return fig

    def get_origin_performance_data(self, df):
        """Prepara dados para tabela de desempenho por origem"""
        if 'origem' not in df.columns or df.empty:
            return None
            
        grouped = df.groupby('origem').agg({
            'contatos': 'sum',
            'aproveitados': 'sum',
            'vendas': 'sum'
        }).reset_index()
        
        grouped['%_aproveitamento'] = grouped['aproveitados'] / grouped['contatos']
        grouped['taxa_conversao'] = grouped['vendas'] / grouped['contatos']
        grouped['conversao_ap'] = grouped['vendas'] / grouped['aproveitados']
        grouped['lead_por_venda'] = grouped['contatos'] / grouped['vendas']
        grouped.replace([np.inf, -np.inf], np.nan, inplace=True)
        grouped = grouped.sort_values('vendas', ascending=False)
        
        # Preparar dados para tabela
        table_data = [["Origem", "Contatos", "Aproveitados", "Vendas", "% Aproveit.", "Conv. Total", "Conv. Leads", "Lead/Venda"]]
        
        for _, row in grouped.iterrows():
            lead_per_sale = "N/A" if pd.isna(row['lead_por_venda']) or np.isinf(row['lead_por_venda']) else f"{row['lead_por_venda']:,.1f}"
            table_data.append([
                row['origem'],
                f"{row['contatos']:,.0f}",
                f"{row['aproveitados']:,.0f}",
                f"{row['vendas']:,.0f}",
                f"{row['%_aproveitamento']:.1%}",
                f"{row['taxa_conversao']:.1%}",
                f"{row['conversao_ap']:.1%}" if not pd.isna(row['conversao_ap']) else "N/A",
                lead_per_sale
            ])
        
        return table_data

    def create_conversion_by_channel_figure(self, df):
        """Cria figura para conversão por canal"""
        if 'origem' not in df.columns or df.empty:
            return None
            
        grouped = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum',
            'aproveitados': 'sum'
        }).reset_index()
        
        grouped['taxa_conversao'] = grouped['vendas'] / grouped['contatos']
        grouped['conversao_ap'] = grouped['vendas'] / grouped['aproveitados']
        grouped = grouped[grouped['contatos'] >= 100]
        
        if grouped.empty:
            return None
            
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        # Gráfico 1: Conversão total
        grouped_sorted = grouped.sort_values('taxa_conversao', ascending=False)
        bars1 = ax1.bar(grouped_sorted['origem'], grouped_sorted['taxa_conversao'], color='skyblue')
        ax1.set_title("Conversão Total por Origem")
        ax1.set_ylabel("Taxa de Conversão")
        ax1.tick_params(axis='x', rotation=90, labelsize=8)
        ax1.grid(axis='y', linestyle='--', alpha=0.7)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.0%}'))
        
        for bar in bars1:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.01,
                    f'{height:.1%}', ha='center', va='bottom', fontsize=6)
        
        # Gráfico 2: Conversão de leads
        grouped_sorted = grouped.sort_values('conversao_ap', ascending=False)
        bars2 = ax2.bar(grouped_sorted['origem'], grouped_sorted['conversao_ap'], color='lightgreen')
        ax2.set_title("Conversão de Leads por Origem")
        ax2.set_ylabel("Taxa de Conversão")
        ax2.tick_params(axis='x', rotation=90, labelsize=8)
        ax2.grid(axis='y', linestyle='--', alpha=0.7)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.0%}'))
        
        for bar in bars2:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 0.01,
                    f'{height:.1%}', ha='center', va='bottom', fontsize=6)
        
        plt.tight_layout()
        return fig

    def create_monthly_trend_figure(self, df):
        """Cria figura para evolução mensal"""
        if 'periodo_dt' not in df.columns or df.empty:
            return None
            
        monthly = df.groupby('periodo_dt').agg({
            'contatos': 'sum',
            'aproveitados': 'sum',
            'vendas': 'sum'
        }).sort_index()
        
        monthly['taxa_conversao'] = monthly['vendas'] / monthly['contatos']
        monthly['conversao_ap'] = monthly['vendas'] / monthly['aproveitados']
        monthly['periodo_formatado'] = monthly.index.map(self.format_period_display)
        
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
        
        # Gráfico 1: Volume
        ax1.plot(monthly['periodo_formatado'], monthly['contatos'], marker='o', label='Contatos')
        ax1.plot(monthly['periodo_formatado'], monthly['aproveitados'], marker='o', label='Aproveitados')
        ax1.plot(monthly['periodo_formatado'], monthly['vendas'], marker='o', label='Vendas')
        ax1.set_title("Evolução Mensal - Volume")
        ax1.set_ylabel("Quantidade")
        ax1.grid(True, linestyle='--', alpha=0.7)
        ax1.legend()
        ax1.set_xticks(range(len(monthly)))
        ax1.set_xticklabels(monthly['periodo_formatado'], rotation=45, ha='right')
        
        # Gráfico 2: Taxas
        ax2.plot(monthly['periodo_formatado'], monthly['taxa_conversao'], marker='o', label='Conversão Total')
        ax2.plot(monthly['periodo_formatado'], monthly['conversao_ap'], marker='o', label='Conversão de Aproveitados')
        ax2.set_title("Evolução Mensal - Taxas de Conversão")
        ax2.set_ylabel("Taxa")
        ax2.grid(True, linestyle='--', alpha=0.7)
        ax2.legend()
        ax2.set_xticks(range(len(monthly)))
        ax2.set_xticklabels(monthly['periodo_formatado'], rotation=45, ha='right')
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f'{y:.0%}'))

        plt.tight_layout()
        return fig

    def create_top_channels_figure(self, df):
        """Cria figura para top canais"""
        if 'origem' not in df.columns or df.empty:
            return None
            
        channel_efficiency = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum',
            'aproveitados': 'sum'
        }).reset_index()
        
        channel_efficiency['taxa_conversao'] = channel_efficiency['vendas'] / channel_efficiency['contatos']
        channel_efficiency['conversao_ap'] = channel_efficiency['vendas'] / channel_efficiency['aproveitados']
        channel_efficiency = channel_efficiency[channel_efficiency['contatos'] >= 50]
        
        if channel_efficiency.empty:
            return None
            
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        # Top 5 por conversão total
        top_conversion = channel_efficiency.nlargest(5, 'taxa_conversao').sort_values('taxa_conversao', ascending=True)
        bars1 = ax1.barh(top_conversion['origem'], top_conversion['taxa_conversao'], color='skyblue')
        ax1.set_title("Top 5 Canais - Conversão Total")
        ax1.set_xlabel("Taxa de Conversão")
        ax1.grid(axis='x', linestyle='--', alpha=0.7)
        ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0%}'))
        
        for bar in bars1:
            width = bar.get_width()
            ax1.text(width + 0.01, bar.get_y() + bar.get_height()/2, 
                    f'{width:.1%}', ha='left', va='center')
        
        # Top 5 por conversão de leads
        top_lead_conversion = channel_efficiency.nlargest(5, 'conversao_ap').sort_values('conversao_ap', ascending=True)
        bars2 = ax2.barh(top_lead_conversion['origem'], top_lead_conversion['conversao_ap'], color='lightgreen')
        ax2.set_title("Top 5 Canais - Conversão de Leads")
        ax2.set_xlabel("Taxa de Conversão")
        ax2.grid(axis='x', linestyle='--', alpha=0.7)
        ax2.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0%}'))
        
        for bar in bars2:
            width = bar.get_width()
            ax2.text(width + 0.01, bar.get_y() + bar.get_height()/2, 
                    f'{width:.1%}', ha='left', va='center')
        
        plt.tight_layout()
        return fig

    def create_sales_efficiency_figure(self, df):
        """Cria figura para eficiência de vendas"""
        if 'contatos' not in df.columns or 'vendas' not in df.columns or df.empty:
            return None
            
        grouped = df.groupby('origem').agg({
            'contatos': 'sum',
            'vendas': 'sum'
        }).reset_index()
        
        grouped['eficiencia'] = (grouped['vendas'] / grouped['contatos']) * 100
        filtered = grouped[grouped['vendas'] >= 10]
        
        if filtered.empty:
            return None
            
        top_15 = filtered.nlargest(15, 'eficiencia').sort_values('eficiencia', ascending=True)
        
        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.barh(top_15['origem'], top_15['eficiencia'], color='skyblue')
        ax.set_title("Eficiência de Vendas (Top 15)")
        ax.set_xlabel("Eficiência (%)")
        ax.grid(axis='x', linestyle='--', alpha=0.7)
        
        for bar in bars:
            width = bar.get_width()
            ax.text(width + 0.5, bar.get_y() + bar.get_height()/2, 
                    f'{width:.1f}%', ha='left', va='center')
        
        plt.tight_layout()
        return fig

    def create_correlation_figure(self, df):
        """Cria figura para correlação"""
        df_corr = self.calculate_correlations(df)
        
        if df_corr.empty:
            return None
            
        top_15 = df_corr.head(15).sort_values('correlacao', ascending=False)
        bottom_15 = df_corr.tail(15).sort_values('correlacao', ascending=True)
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 7))
        
        # Melhores correlações
        bars1 = ax1.barh(top_15['origem'], top_15['correlacao'], color='#2ecc71')
        ax1.set_title('TOP 15 - Melhores Correlações')
        ax1.set_xlabel("Coeficiente de Correlação")
        ax1.set_xlim(-1.1, 1.1)
        ax1.grid(axis='x', linestyle='--', alpha=0.7)
        
        for bar in bars1:
            width = bar.get_width()
            ax1.text(width + 0.02 if width > 0 else width - 0.1, 
                    bar.get_y() + bar.get_height()/2, 
                    f'{width:.2f}', ha='left' if width > 0 else 'right', va='center')
        
        # Piores correlações
        bars2 = ax2.barh(bottom_15['origem'], bottom_15['correlacao'], color='#e74c3c')
        ax2.set_title('TOP 15 - Piores Correlações')
        ax2.set_xlabel("Coeficiente de Correlação")
        ax2.set_xlim(-1.1, 1.1)
        ax2.grid(axis='x', linestyle='--', alpha=0.7)
        
        for bar in bars2:
            width = bar.get_width()
            ax2.text(width + 0.02 if width > 0 else width - 0.1, 
                    bar.get_y() + bar.get_height()/2, 
                    f'{width:.2f}', ha='left' if width > 0 else 'right', va='center')
        
        plt.tight_layout()
        return fig

    def create_scatter_figure(self, df):
        """Cria figura para dispersão"""
        required_columns = ['contatos', 'aproveitados', 'vendas', 'origem']
        if any(col not in df.columns for col in required_columns) or df.empty:
            return None
            
        df_agg = df.groupby('origem', as_index=False).agg({
            'contatos': 'sum',
            'aproveitados': 'sum',
            'vendas': 'sum'
        })
        
        df_agg = df_agg[
            (df_agg['contatos'] > 0) & 
            (df_agg['vendas'] > 0) & 
            (df_agg['aproveitados'] > 0)
        ]
        
        if df_agg.empty:
            return None
            
        top_origins = df_agg.nlargest(15, 'contatos')['origem']
        df_filtered = df_agg[df_agg['origem'].isin(top_origins)]
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        
        # Gráfico 1: Leads vs Vendas
        sns.scatterplot(
            data=df_filtered,
            x='contatos',
            y='vendas',
            hue='origem',
            size='vendas',
            sizes=(50, 300),
            ax=ax1
        )
        ax1.set_title('Leads vs Vendas por Canal')
        ax1.set_xlabel("Total de Leads")
        ax1.set_ylabel("Total de Vendas")
        ax1.grid(True, linestyle='--', alpha=0.7)
        
        # Remover a legenda do primeiro gráfico
        if ax1.get_legend() is not None:
            ax1.get_legend().remove()
        
        try:
            sns.regplot(data=df_filtered, x='contatos', y='vendas', 
                        scatter=False, ax=ax1, color='gray', line_kws={'alpha': 0.5})
        except:
            pass
        
        # Gráfico 2: Leads Aproveitados vs Vendas
        sns.scatterplot(
            data=df_filtered,
            x='aproveitados',
            y='vendas',
            hue='origem',
            size='vendas',
            sizes=(50, 300),
            ax=ax2
        )
        ax2.set_title('Leads Aproveitados vs Vendas')
        ax2.set_xlabel("Leads Aproveitados")
        ax2.set_ylabel("Total de Vendas")
        ax2.grid(True, linestyle='--', alpha=0.7)
        
        try:
            sns.regplot(data=df_filtered, x='aproveitados', y='vendas', 
                        scatter=False, ax=ax2, color='gray', line_kws={'alpha': 0.5})
        except:
            pass
        
        # Ajustar layout da legenda (apenas para o segundo gráfico)
        plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        
        plt.tight_layout()
        return fig
    
    def generate_pdf_report(self, df, file_path):
        """Gera relatório em PDF com base nos dados"""
        # Configurações do documento
        doc = SimpleDocTemplate(
            file_path,
            pagesize=letter,
            rightMargin=40,
            leftMargin=40,
            topMargin=40,
            bottomMargin=40
        )
        
        styles = getSampleStyleSheet()
        
        # Definir estilos personalizados
        styles.add(ParagraphStyle(
            name='TitleStyle',
            fontSize=18,
            alignment=TA_CENTER,
            spaceAfter=12
        ))
        
        styles.add(ParagraphStyle(
            name='SubtitleStyle',
            fontSize=12,
            alignment=TA_CENTER,
            textColor=colors.grey,
            spaceAfter=20
        ))
        
        styles.add(ParagraphStyle(
            name='SectionStyle',
            fontSize=14,
            spaceBefore=20,
            spaceAfter=10
        ))
        
        styles.add(ParagraphStyle(
            name='BodyStyle',
            fontSize=10,
            alignment=TA_JUSTIFY,
            leading=14
        ))
        
        styles.add(ParagraphStyle(
            name='MetricStyle',
            fontSize=14,
            textColor=colors.darkblue,
            spaceAfter=5
        ))
        
        styles.add(ParagraphStyle(
            name='FooterStyle',
            fontSize=8,
            textColor=colors.grey,
            alignment=TA_CENTER
        ))
        
        elements = []
        fig_paths = []
        city = self.city_var.get()
        period_start = self.period_var_start.get()
        period_end = self.period_var_end.get()
        
        try:
            # Cabeçalho
            elements.append(Paragraph("Relatório Completo de Performance de Leads", styles['TitleStyle']))
            
            # Formatar período
            start_dt = self.parse_period(period_start)
            end_dt = self.parse_period(period_end)
            formatted_start = self.format_period_display(start_dt)
            formatted_end = self.format_period_display(end_dt)
            
            elements.append(Paragraph(f"{city} | {formatted_start} - {formatted_end}", styles['SubtitleStyle']))
            elements.append(Spacer(1, 12))
            
            # Seção: Visão Geral
            elements.append(Paragraph("1. Visão Geral", styles['SectionStyle']))
            
            # Calcular métricas
            total_contacts = df['contatos'].sum() if 'contatos' in df.columns else 0
            total_leads = df['aproveitados'].sum() if 'aproveitados' in df.columns else 0
            total_sales = df['vendas'].sum() if 'vendas' in df.columns else 0
            
            conversion_rate = total_sales / total_contacts if total_contacts > 0 else 0
            lead_conversion_rate = total_sales / total_leads if total_leads > 0 else 0
            
            # Criar tabela de métricas
            metrics_data = [
                ["Métrica", "Valor", "Insight"],
                ["Total de Contatos", f"{total_contacts:,.0f}", "Volume total de oportunidades geradas"],
                ["Leads Aproveitados", f"{total_leads:,.0f}", f"({total_leads/total_contacts:.1%} dos contatos)" if total_contacts > 0 else ""],
                ["Vendas Fechadas", f"{total_sales:,.0f}", f"({conversion_rate:.1%} de conversão geral)" if total_contacts > 0 else ""],
                ["Conversão de Leads", f"{lead_conversion_rate:.1%}" if total_leads > 0 else "N/A", "Eficiência no aproveitamento de oportunidades"]
            ]
            
            metrics_table = Table(metrics_data)
            metrics_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ]))
            elements.append(metrics_table)
            elements.append(Spacer(1, 20))
            
            # Gráficos da Visão Geral
            fig1 = self.create_summary_figures(df)
            if fig1:
                fig_path1 = tempfile.mktemp(suffix='.png')
                fig1.savefig(fig_path1, bbox_inches='tight')
                fig_paths.append(fig_path1)
                elements.append(Image(fig_path1, width=500, height=300))
                elements.append(Spacer(1, 20))
            
            # Seção: Desempenho por Origem
            elements.append(PageBreak())
            elements.append(Paragraph("2. Desempenho por Origem", styles['SectionStyle']))
            
            # Tabela de desempenho
            table_data = self.get_origin_performance_data(df)
            if table_data:
                origin_table = Table(table_data)
                origin_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                    ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                ]))
                elements.append(origin_table)
                elements.append(Spacer(1, 20))
            
            # Seção: Conversão por Canal
            elements.append(PageBreak())
            elements.append(Paragraph("3. Conversão por Canal", styles['SectionStyle']))
            
            fig2 = self.create_conversion_by_channel_figure(df)
            if fig2:
                fig_path2 = tempfile.mktemp(suffix='.png')
                fig2.savefig(fig_path2, bbox_inches='tight')
                fig_paths.append(fig_path2)
                elements.append(Image(fig_path2, width=500, height=300))
                elements.append(Spacer(1, 20))
            
            # Seção: Evolução Mensal
            elements.append(PageBreak())
            elements.append(Paragraph("4. Evolução Mensal", styles['SectionStyle']))
            elements.append(Paragraph("Análise: Acompanhe a evolução dos principais indicadores ao longo do tempo. Tendências de crescimento em contatos e vendas indicam eficácia nas estratégias. Quedas consistentes podem sinalizar problemas operacionais ou de mercado.", styles['BodyStyle']))
    
            fig3 = self.create_monthly_trend_figure(df)
            if fig3:
                fig_path3 = tempfile.mktemp(suffix='.png')
                fig3.savefig(fig_path3, bbox_inches='tight')
                fig_paths.append(fig_path3)
                elements.append(Image(fig_path3, width=500, height=400))
                elements.append(Spacer(1, 20))
            
            # Seção: Top Canais
            elements.append(PageBreak())
            elements.append(Paragraph("5. Top Canais", styles['SectionStyle']))
            elements.append(Paragraph("Análise: Identifique os canais com melhor desempenho. Canais com alta conversão representam oportunidades de investimento. Canais com baixa conversão podem precisar de otimização ou realocação de recursos.", styles['BodyStyle']))
    
            fig4 = self.create_top_channels_figure(df)
            if fig4:
                fig_path4 = tempfile.mktemp(suffix='.png')
                fig4.savefig(fig_path4, bbox_inches='tight')
                fig_paths.append(fig_path4)
                elements.append(Image(fig_path4, width=500, height=300))
                elements.append(Spacer(1, 20))
            
            # Seção: Eficiência de Vendas
            elements.append(PageBreak())
            elements.append(Paragraph("6. Eficiência de Vendas", styles['SectionStyle']))
            elements.append(Paragraph("Análise: Mede a porcentagem de vendas por lead. Quanto maior o valor, melhor será para investir", styles['BodyStyle']))
    
            fig5 = self.create_sales_efficiency_figure(df)
            if fig5:
                fig_path5 = tempfile.mktemp(suffix='.png')
                fig5.savefig(fig_path5, bbox_inches='tight')
                fig_paths.append(fig_path5)
                elements.append(Image(fig_path5, width=500, height=350))
                elements.append(Spacer(1, 20))
            
            # Seção: Correlação Leads-Vendas
            elements.append(PageBreak())
            elements.append(Paragraph("7. Correlação Leads-Vendas", styles['SectionStyle']))
            elements.append(Paragraph("Análise: Correlação mede a relação entre leads e vendas. Valores próximos a 1 indicam que o aumento de leads acompanha o aumento de vendas. Valores próximos a -1 indicam relação inversa. Valores próximos a 0 indicam pouca relação entre as variáveis.", styles['BodyStyle']))
            elements.append(Paragraph("Interpretação: Correlação > 0.7 = Forte relação positiva | 0.3-0.7 = Relação moderada | < 0.3 = Fraca relação", styles['BodyStyle']))
    
            fig6 = self.create_correlation_figure(df)
            if fig6:
                fig_path6 = tempfile.mktemp(suffix='.png')
                fig6.savefig(fig_path6, bbox_inches='tight')
                fig_paths.append(fig_path6)
                elements.append(Image(fig_path6, width=500, height=350))
                elements.append(Spacer(1, 20))
            
            # Seção: Dispersão Leads x Vendas
            elements.append(PageBreak())
            elements.append(Paragraph("8. Dispersão Leads x Vendas", styles['SectionStyle']))
            elements.append(Paragraph("Análise: Mostra a relação entre volume de leads e vendas geradas. Canais no canto superior direito (muitos leads e vendas) são os mais eficientes. Canais com muitos leads e poucas vendas precisam de otimização.", styles['BodyStyle']))

            fig7 = self.create_scatter_figure(df)
            if fig7:
                fig_path7 = tempfile.mktemp(suffix='.png')
                fig7.savefig(fig_path7, bbox_inches='tight')
                fig_paths.append(fig_path7)
                elements.append(Image(fig_path7, width=500, height=300))
                elements.append(Spacer(1, 20))
            
            # Rodapé
            elements.append(Spacer(1, 20))
            elements.append(Paragraph(f"Relatório gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['FooterStyle']))
            elements.append(Paragraph("Sistema de Análise de Leads | Dados Confidenciais", styles['FooterStyle']))
            
            # Construir PDF
            doc.build(elements)
            
        finally:
            # Remover arquivos temporários
            for path in fig_paths:
                try:
                    os.remove(path)
                except:
                    pass

if __name__ == "__main__":
    root = ctk.CTk()
    app = LeadAnalyzerApp(root)
    root.mainloop()