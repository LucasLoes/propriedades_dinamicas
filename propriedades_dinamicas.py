import tkinter as tk
from tkinter import ttk
import win32gui
import win32com.client
import pythoncom
import os
import stat
from datetime import datetime
import threading
import time
import logging
import webbrowser # Módulo importado para os links do rodapé

# --- CONFIGURAÇÃO DO LOG ROBUSTA ---
# Salva o log na pasta do usuário para evitar problemas de permissão
log_path = os.path.join(os.path.expanduser("~"), "propriedades_dinamicas_log.log")

logging.basicConfig(
    level=logging.DEBUG, # Nível de log máximo, como solicitado
    format='%(asctime)s [%(levelname)s | %(funcName)s] - %(message)s',
    filename=log_path,
    filemode='w' # 'w' para um log limpo a cada execução
)

# --- MÓDULO DE FUNÇÕES UTILITÁRIAS ---
def formatar_tamanho(tamanho_bytes):
    if tamanho_bytes is None: return "..."
    try:
        tamanho_bytes = int(tamanho_bytes)
        if tamanho_bytes >= 1024**3: valor = f"{tamanho_bytes / 1024**3:.2f} GB"
        elif tamanho_bytes >= 1024**2: valor = f"{tamanho_bytes / 1024**2:.2f} MB"
        elif tamanho_bytes >= 1024: valor = f"{tamanho_bytes / 1024:.2f} KB"
        else: valor = f"{tamanho_bytes} bytes"
        return valor
    except (ValueError, TypeError): return "N/A"

def obter_selecao_explorer():
    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        hwnd = win32gui.GetForegroundWindow()
        try: classe_janela = win32gui.GetClassName(hwnd)
        except pythoncom.pywintypes.error: return None, None
        if classe_janela == "CabinetWClass":
            for janela in shell.Windows():
                try:
                    if janela.HWND == hwnd:
                        selecao = janela.Document.SelectedItems()
                        caminhos = [item.Path for item in selecao]
                        if not caminhos and hasattr(janela.Document.Folder, 'Self'):
                            caminhos = [janela.Document.Folder.Self.Path]
                        return caminhos, janela.Document.Folder.Title
                except pythoncom.com_error: continue
        return None, None
    except Exception as e:
        logging.error(f"Erro inesperado em obter_selecao_explorer: {e}", exc_info=True)
        return None, None
    finally: pythoncom.CoUninitialize()

# --- SCANNER EM SEGUNDO PLANO ---
class FolderScannerThread(threading.Thread):
    def __init__(self, paths, callback, stop_event):
        super().__init__()
        self.paths, self.callback, self.stop_event = paths if isinstance(paths, list) else [paths], callback, stop_event
        self.daemon = True

    def run(self):
        total_size, file_count, folder_count = 0, 0, 0
        try:
            for path in self.paths:
                if self.stop_event.is_set(): self.callback({'status': 'cancelled'}); return
                if os.path.isdir(path):
                    if len(self.paths) > 1: folder_count += 1
                    for dirpath, dirnames, filenames in os.walk(path):
                        if self.stop_event.is_set(): self.callback({'status': 'cancelled'}); return
                        file_count += len(filenames); folder_count += len(dirnames)
                        for f in filenames:
                            try: total_size += os.path.getsize(os.path.join(dirpath, f))
                            except OSError: continue
                else:
                    file_count += 1
                    try: total_size += os.path.getsize(path)
                    except OSError: continue
        except Exception: self.callback({'status': 'error'}); return
        self.callback({'status': 'done', 'size': total_size, 'files': file_count, 'folders': folder_count})

# --- WIDGET PERSONALIZADO ---
class SelectableLabel(ttk.Entry):
    def __init__(self, parent, style_name='Value.TEntry', **kwargs):
        super().__init__(parent, style=style_name, **kwargs)
        self.text_variable = tk.StringVar()
        self.config(textvariable=self.text_variable, state='readonly')

    def set_text(self, text):
        self.text_variable.set(text)

# --- INTERFACE GRÁFICA ---
class AppPropriedadesDinamicas(tk.Tk):
    def __init__(self):
        super().__init__()
        logging.info("--- Aplicação iniciada ---")
        logging.info(f"Log sendo salvo em: {log_path}")
        self.title("Propriedades Dinâmicas")
        self.geometry("380x675+150+150"); self.minsize(380, 600)
        
        self._setup_styles_and_fonts()
        
        # Frame principal para o conteúdo dinâmico
        self.main_frame = ttk.Frame(self, style='App.TFrame', padding="20 25 20 25")
        self.main_frame.pack(expand=True, fill="both")
        
        # Cria o rodapé, que ficará fixo na parte inferior
        self._criar_rodape()
        
        self.is_scanning = False
        self.scanner_thread, self.stop_scanner_event = None, None
        self.ultima_selecao_vista = None
        self.monitoramento_ativo = True
        
        self.thread_monitor = threading.Thread(target=self.monitorar_selecao, daemon=True); self.thread_monitor.start()
        self.protocol("WM_DELETE_WINDOW", self.ao_fechar)
        self.mostrar_view_inicial()

    def _setup_styles_and_fonts(self):
        self.BG_COLOR, self.CARD_COLOR, self.SHADOW_COLOR, self.TEXT_COLOR, self.ACCENT_COLOR, self.SUBTLE_TEXT_COLOR = "#f0f2f5", "#ffffff", "#d0d0d0", "#050505", "#0078d4", "#65676b"
        self.PULSATE_COLOR = "#aaaaaa"
        self.style = ttk.Style(self); self.style.theme_use('clam')
        
        self.style.configure('.', background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=('Segoe UI', 10), borderwidth=0, relief='flat')
        self.style.configure('App.TFrame', background=self.BG_COLOR)
        self.style.configure('Card.TFrame', background=self.CARD_COLOR)
        self.style.configure('Shadow.TFrame', background=self.SHADOW_COLOR)
        self.style.configure('TLabel', background=self.CARD_COLOR, foreground=self.TEXT_COLOR)
        self.style.configure('BG.TLabel', background=self.BG_COLOR)
        self.style.configure('Title.TLabel', font=('Segoe UI', 12, 'bold'), foreground=self.ACCENT_COLOR, background=self.CARD_COLOR)
        self.style.configure('Key.TLabel', font=('Segoe UI', 10), background=self.CARD_COLOR, foreground=self.SUBTLE_TEXT_COLOR)
        self.style.configure('Hero.Icon.TLabel', font=('Segoe UI Symbol', 48), background=self.BG_COLOR, foreground=self.ACCENT_COLOR)
        self.style.configure('Hero.Text.TLabel', font=('Segoe UI', 14, 'bold'), background=self.BG_COLOR, foreground=self.TEXT_COLOR)

        common_entry_opts = {'borderwidth': 0, 'fieldbackground': self.CARD_COLOR, 'insertwidth': 0}
        self.style.configure('Value.TEntry', **common_entry_opts)
        self.style.map('Value.TEntry', foreground=[('readonly', self.TEXT_COLOR)])
        self.style.configure('Pulsate.TEntry', **common_entry_opts)
        self.style.map('Pulsate.TEntry', foreground=[('readonly', self.PULSATE_COLOR)])
        
        self.style.configure('FileName.TEntry', **common_entry_opts)
        self.style.map('FileName.TEntry', foreground=[('readonly', self.TEXT_COLOR)], font=[('readonly', ('Segoe UI', 10, 'bold'))])
        self.style.configure('Path.TEntry', **common_entry_opts)
        self.style.map('Path.TEntry', foreground=[('readonly', self.SUBTLE_TEXT_COLOR)], font=[('readonly', ('Segoe UI', 9))])

        # Estilos específicos do rodapé
        self.style.configure('Footer.TLabel', background=self.BG_COLOR, foreground=self.SUBTLE_TEXT_COLOR, font=('Segoe UI', 9))
        self.style.configure('Link.TLabel', background=self.BG_COLOR, foreground=self.ACCENT_COLOR, font=('Segoe UI', 9, 'underline'))

    def ao_fechar(self):
        self.monitoramento_ativo = False
        self._stop_current_scanner()
        self.destroy()

    def _stop_current_scanner(self):
        self.is_scanning = False
        if self.scanner_thread and self.scanner_thread.is_alive(): self.stop_scanner_event.set()

    def _limpar_frame_principal(self):
        for widget in self.main_frame.winfo_children(): widget.destroy()

    def _criar_card(self, title):
        shadow = ttk.Frame(self.main_frame, style='Shadow.TFrame')
        shadow.pack(fill='x', pady=(0, 10), padx=2)
        card = ttk.Frame(shadow, style='Card.TFrame', padding=15)
        card.pack(fill='x', pady=(0, 2), padx=(0, 2))
        ttk.Label(card, text=title, style='Title.TLabel').pack(anchor='w', pady=(0, 15))
        return card
    
    def _criar_rodape(self):
        """
        Cria e posiciona o frame do rodapé com informações e links.
        Este frame é filho da janela principal (self) para ficar fixo.
        """
        footer_frame = ttk.Frame(self, style='App.TFrame', padding="10 20 10 20")
        footer_frame.pack(side='bottom', fill='x')

        ttk.Separator(footer_frame).pack(fill='x', pady=(0, 10))

        ttk.Label(footer_frame, text="Desenvolvido por Lucas Ladeira Loes.", style='Footer.TLabel').pack(anchor='center')
        ttk.Label(footer_frame, text="Todos os direitos reservados. 2025", style='Footer.TLabel').pack(anchor='center')
        
        def abrir_link(url):
            webbrowser.open_new_tab(url)

        links_frame = ttk.Frame(footer_frame, style='App.TFrame')
        links_frame.pack(anchor='center', pady=5)

        resume_link = ttk.Label(links_frame, text="Currículo Interativo", style='Link.TLabel', cursor="hand2")
        resume_link.pack(side='left', padx=5)
        resume_link.bind("<Button-1>", lambda e: abrir_link("https://lucasloes.github.io/interactive-resume/"))

        ttk.Label(links_frame, text="|", style='Footer.TLabel').pack(side='left')

        linkedin_link = ttk.Label(links_frame, text="LinkedIn", style='Link.TLabel', cursor="hand2")
        linkedin_link.pack(side='left', padx=5)
        linkedin_link.bind("<Button-1>", lambda e: abrir_link("https://www.linkedin.com/in/lucas-ladeira-loes/"))

    def mostrar_view_inicial(self):
        self._limpar_frame_principal()
        label = ttk.Label(self.main_frame, text="Selecione um item no Explorer...", style='BG.TLabel', font=('Segoe UI', 12))
        label.pack(expand=True)

    def atualizar_interface(self, caminhos, titulo_pasta):
        self._limpar_frame_principal()
        self._stop_current_scanner()

        if not caminhos:
            self.mostrar_view_inicial()
            return

        is_dir = os.path.isdir(caminhos[0]) if len(caminhos) == 1 else False
        nome = os.path.basename(caminhos[0]) if len(caminhos) == 1 else ""
        local = os.path.dirname(caminhos[0]) if len(caminhos) == 1 else titulo_pasta
        
        if len(caminhos) > 1:
            tipo_texto, icon = f"{len(caminhos)} Itens Selecionados", "🗂️"
        else:
            tipo_texto, icon = ("Pasta", "📁") if is_dir else ("Arquivo", "📄")

        hero_frame = ttk.Frame(self.main_frame, style='App.TFrame')
        hero_frame.pack(fill='x', pady=(0, 20))
        ttk.Label(hero_frame, text=icon, style='Hero.Icon.TLabel').pack(anchor='center')
        ttk.Label(hero_frame, text=tipo_texto, style='Hero.Text.TLabel').pack(anchor='center', pady=5)

        card_details = self._criar_card("Detalhes")
        
        if len(caminhos) == 1:
            nome_frame = ttk.Frame(card_details, style='Card.TFrame')
            nome_frame.pack(fill='x', expand=True, pady=2)
            ttk.Label(nome_frame, text="Nome:", style='Key.TLabel', width=9).pack(side='left', anchor='n')
            nome_label = SelectableLabel(nome_frame, style_name='FileName.TEntry')
            nome_label.set_text(nome)
            nome_label.pack(side='left', fill='x', expand=True)

        caminho_frame = ttk.Frame(card_details, style='Card.TFrame')
        caminho_frame.pack(fill='x', expand=True, pady=2)
        label_caminho = "Caminho:" if len(caminhos) == 1 else "Local:"
        ttk.Label(caminho_frame, text=label_caminho, style='Key.TLabel', width=9).pack(side='left', anchor='n')
        local_label = SelectableLabel(caminho_frame, style_name='Path.TEntry')
        local_label.set_text(local)
        local_label.pack(side='left', fill='x', expand=True)

        card_props = self._criar_card("Propriedades")
        
        def criar_linha_prop(parent, key):
            prop_frame = ttk.Frame(parent, style='Card.TFrame')
            prop_frame.pack(fill='x', expand=True, pady=3)
            ttk.Label(prop_frame, text=key, style='Key.TLabel', width=15).pack(side='left', anchor='n')
            value_label = SelectableLabel(prop_frame)
            value_label.pack(side='left', fill='x', expand=True)
            return value_label
        
        if len(caminhos) == 1 and not is_dir:
            try:
                stats = os.stat(caminhos[0])
                criar_linha_prop(card_props, "Tamanho:").set_text(formatar_tamanho(stats.st_size))
                criar_linha_prop(card_props, "Modificado em:").set_text(datetime.fromtimestamp(stats.st_mtime).strftime('%d/%m/%Y %H:%M:%S'))
                criar_linha_prop(card_props, "Criado em:").set_text(datetime.fromtimestamp(stats.st_ctime).strftime('%d/%m/%Y %H:%M:%S'))
            except Exception as e:
                criar_linha_prop(card_props, "Erro:").set_text(str(e))
        else:
            self.label_tamanho = criar_linha_prop(card_props, "Tamanho Total:")
            self.label_arquivos = criar_linha_prop(card_props, "Total de Arquivos:")
            self.label_pastas = criar_linha_prop(card_props, "Total de Pastas:")
            
            calculando_labels = [self.label_tamanho, self.label_arquivos, self.label_pastas]
            for label in calculando_labels:
                label.set_text("Calculando...")
            
            self.is_scanning = True
            self._animate_pulsate(calculando_labels)
            
            self.stop_scanner_event = threading.Event()
            self.scanner_thread = FolderScannerThread(caminhos, self.atualizar_contagem_async, self.stop_scanner_event)
            self.scanner_thread.start()

        if len(caminhos) == 1 and not is_dir:
            card_attrs = self._criar_card("Atributos")
            
            try:
                file_attributes = os.stat(caminhos[0]).st_mode
                readonly = not (file_attributes & stat.S_IWRITE)
                criar_linha_prop(card_attrs, "Somente Leitura:").set_text("Sim" if readonly else "Não")
            except Exception as e:
                criar_linha_prop(card_attrs, "Erro ao ler:").set_text(str(e))

    def _animate_pulsate(self, widgets, direction='down'):
        if not self.is_scanning:
            for widget in widgets:
                if widget.winfo_exists():
                    widget.configure(style='Value.TEntry')
            return
        
        next_style = 'Pulsate.TEntry' if direction == 'down' else 'Value.TEntry'
        next_direction = 'up' if direction == 'down' else 'down'
        
        for widget in widgets:
            if widget.winfo_exists():
                widget.configure(style=next_style)
        
        self.after(700, self._animate_pulsate, widgets, next_direction)

    def atualizar_contagem_async(self, data):
        self.is_scanning = False
        if data['status'] == 'done':
            self.after(0, self.label_tamanho.set_text, formatar_tamanho(data['size']))
            self.after(0, self.label_arquivos.set_text, f"{data['files']:,}")
            self.after(0, self.label_pastas.set_text, f"{data['folders']:,}")

    def monitorar_selecao(self):
        try:
            logging.debug("Thread de monitoramento iniciada.")
            while self.monitoramento_ativo:
                caminhos, titulo = obter_selecao_explorer()
                if caminhos is None:
                    time.sleep(0.5); continue
                id_selecao_atual = tuple(sorted(caminhos))
                if id_selecao_atual != self.ultima_selecao_vista:
                    logging.info(f"MUDANÇA DE SELEÇÃO: {id_selecao_atual}")
                    self.ultima_selecao_vista = id_selecao_atual
                    self.after(0, self.atualizar_interface, caminhos, titulo)
                time.sleep(0.3)
        except Exception as e:
            logging.critical("ERRO FATAL NO THREAD DE MONITORAMENTO.", exc_info=True)

if __name__ == "__main__":
    app = AppPropriedadesDinamicas()
    app.mainloop()
