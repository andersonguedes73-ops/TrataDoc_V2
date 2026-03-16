import fitz  # PyMuPDF
import pytesseract
import re
import spacy
import os
import sys
import threading
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk  # MOTOR VISUAL
from PIL import Image, ImageTk
import cv2        # Inteligência Visual (Borrão)
import numpy as np # Manipulação de matrizes de imagem

# --- DETECÇÃO DE SISTEMA ---
SISTEMA = platform.system() # 'Windows' ou 'Darwin' (Mac)

# --- CONFIGURAÇÃO DE AMBIENTE ROBUSTA ---
def obter_raiz():
    if getattr(sys, 'frozen', False): 
        # O pulo do gato: Forçar a usar o _MEIPASS no Windows e o diretório do app no Mac
        if SISTEMA == "Windows": return sys._MEIPASS
        else: return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))
PASTA_RAIZ = obter_raiz()

# Configuração Tesseract (Apenas Windows)
if SISTEMA == "Windows":
    # Busca inteligente pela pasta do Tesseract
    locais_tess = [
        os.path.join(PASTA_RAIZ, "Tesseract-OCR", "tesseract.exe"),
        os.path.join(PASTA_RAIZ, "_internal", "Tesseract-OCR", "tesseract.exe")
    ]
    CAMINHO_TESS = None
    for loc in locais_tess:
        if os.path.exists(loc):
            CAMINHO_TESS = loc
            break
            
    if CAMINHO_TESS:
        pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESS
        os.environ['TESSDATA_PREFIX'] = os.path.join(os.path.dirname(CAMINHO_TESS), "tessdata")

# --- CARREGAMENTO DO MOTOR IA (CAÇADOR AUTOMÁTICO) ---
def carregar_ia():
    # 1. Tenta carregar o padrão do sistema
    try: return spacy.load("pt_core_news_md")
    except: pass

    # 2. Varredura inteligente: procura em todas as subpastas (inclusive a _internal)
    for root, dirs, files in os.walk(PASTA_RAIZ):
        # Se ele achar uma pasta que tenha "pt_core" no nome E o arquivo "config.cfg" dentro dela
        if "pt_core" in root and "config.cfg" in files:
            try: 
                return spacy.load(root)
            except: 
                continue # Se der erro, continua procurando
    return None

nlp = carregar_ia()

# Configuração Visual
ctk.set_appearance_mode("Light")  
ctk.set_default_color_theme("blue")  

class TrataDocApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Título dinâmico nacionalizado
        desc_sis = "Processamento (MacOS)" if SISTEMA == "Darwin" else "Processamento (Windows)"
        self.title(f"TrataDoc - Ferramentas de Documentos v9.0 - {desc_sis}")
        self.geometry("1450x950")
        
        # Carregamento do ícone da janela (PNG para ser universal)
        caminho_icone_janela = os.path.join(PASTA_RAIZ, "icone.png")
        if os.path.exists(caminho_icone_janela):
            try:
                img_icon = ImageTk.PhotoImage(file=caminho_icone_janela)
                self.wm_iconphoto(True, img_icon)
            except: pass

        # --- ESTRUTURA DE DADOS ---
        self.dados = {
            "ocr": {"entrada": [], "prontos": []},
            "merge": {"entrada": [], "prontos": []},
            "tarja": {"entrada": [], "prontos": []}
        }
        self.pagina_atual = 0
        self.zoom_level = 0.8 
        self.doc_aberto = None
        self.caminho_atual = None
        
        # Controle de edição visual
        self.modo_edicao = "Desativado"
        self.rect_id = None
        self.start_x = 0
        self.start_y = 0
        self.historico_estados = []

        # Grid Principal
        self.grid_rowconfigure(0, weight=0) 
        self.grid_rowconfigure(1, weight=1) 
        self.grid_rowconfigure(2, weight=0) 
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)

        self.setup_topbar()
        self.area_trabalho = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self.area_trabalho.grid(row=1, column=1, sticky="nsew")
        self.area_trabalho.grid_rowconfigure(0, weight=1)
        self.area_trabalho.grid_columnconfigure(0, weight=0, minsize=450) # Painel fixo para evitar esmagamento
        self.area_trabalho.grid_columnconfigure(1, weight=1) # Visualizador expande
        
        self.setup_sidebar()
        self.frame_ocr = ctk.CTkScrollableFrame(self.area_trabalho, corner_radius=0, fg_color="transparent")
        self.frame_merge = ctk.CTkScrollableFrame(self.area_trabalho, corner_radius=0, fg_color="transparent")
        self.frame_tarja = ctk.CTkScrollableFrame(self.area_trabalho, corner_radius=0, fg_color="transparent")

        self.setup_ocr_tab()
        self.setup_merge_tab()
        self.setup_tarja_tab()

        self.frame_view = ctk.CTkFrame(self.area_trabalho, corner_radius=0, fg_color="#f2f2f2", border_width=1, border_color="#dee2e6")
        self.frame_view.grid(row=0, column=1, sticky="nsew", padx=2, pady=2)
        self.setup_viewer()

        # Barra de Status
        self.f_status = tk.Frame(self, bd=1, relief=tk.SUNKEN, bg="#e9ecef")
        self.f_status.grid(row=2, column=0, columnspan=2, sticky="ew")
        
        msg_ia = "🧠 Motor IA: ATIVO" if nlp else "⚠️ Motor IA: DESATIVADO"
        self.lbl_status_ia = tk.Label(self.f_status, text=msg_ia, font=("Arial", 9, "bold"), fg="#28a745" if nlp else "#dc3545", bg="#e9ecef")
        self.lbl_status_ia.pack(side="left", padx=10)
        
        self.lbl_dica = tk.Label(self.f_status, text="", font=("Arial", 9, "italic"), fg="#555", bg="#e9ecef")
        self.lbl_dica.pack(side="left", padx=20)
        
        self.progress = ttk.Progressbar(self.f_status, orient="horizontal", mode="determinate")
        self.progress.pack(side="right", fill="x", expand=True, padx=5)

        self.select_frame_by_name("tarja")
        self.sidebar_frame.grid_remove() 
        self.sidebar_visible = False
        self.bind_all("<Motion>", self.check_mouse_position)

    # --- JANELA DE CARREGAMENTO ---
    def mostrar_carregamento(self, mensagem="Processando... Por favor, aguarde."):
        self.janela_load = tk.Toplevel(self)
        self.janela_load.title("Aguarde")
        self.janela_load.geometry("350x100")
        self.janela_load.resizable(False, False)
        self.janela_load.transient(self) # Mantém a janela sempre na frente
        self.janela_load.grab_set() # Bloqueia cliques na janela principal
        
        # Centraliza a janelinha
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 175
        y = self.winfo_y() + (self.winfo_height() // 2) - 50
        self.janela_load.geometry(f"+{x}+{y}")
        
        f = tk.Frame(self.janela_load, bg="white")
        f.pack(expand=True, fill="both")
        
        tk.Label(f, text="⏳", font=("Segoe UI", 24), bg="white").pack(pady=(10,0))
        tk.Label(f, text=mensagem, font=("Segoe UI", 12, "bold"), bg="white").pack()
        self.update()

    def fechar_carregamento(self):
        if hasattr(self, 'janela_load') and self.janela_load:
            self.janela_load.destroy()
            self.janela_load = None

    # --- FUNÇÕES DE NAVEGAÇÃO ---
    def set_dica(self, texto): self.lbl_dica.config(text=texto)
    def limpar_dica(self, event=None): self.lbl_dica.config(text="")

    def setup_topbar(self):
        self.topbar = ctk.CTkFrame(self, height=50, corner_radius=0, fg_color="#102a43")
        self.topbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.btn_menu = ctk.CTkButton(self.topbar, text="☰", width=50, font=("Segoe UI", 24), fg_color="transparent", text_color="white", command=self.show_sidebar)
        self.btn_menu.pack(side="left", padx=5, pady=5)
        ctk.CTkLabel(self.topbar, text="TrataDoc - Ferramentas de Documentos", font=("Segoe UI", 22, "bold"), text_color="#00d2ff").pack(side="left", padx=10)
        ctk.CTkLabel(self.topbar, text="|  Corregedoria MPO", font=("Segoe UI", 16), text_color="#e0e0e0").pack(side="left")

    def setup_sidebar(self):
        self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color="#f8f9fa", border_width=1, border_color="#dee2e6")
        self.sidebar_frame.grid(row=1, column=0, sticky="nsew")

        menus = [
            ("⬛ Ocultação de Dados", "tarja", "Proteção de dados sensíveis e anonimização de rostos."),
            ("🔗 Unificar Arquivos", "merge", "Organização e junção de múltiplos PDFs em um volume único."),
            ("📄 Texto Pesquisável", "ocr", "Conversão de documentos digitalizados em arquivos pesquisáveis.")
        ]
        for t, n, d in menus:
            b = ctk.CTkButton(self.sidebar_frame, text=t, height=55, anchor="w", fg_color="transparent", text_color="black", font=("Segoe UI", 14), command=lambda x=n: self.select_frame_by_name(x))
            b.pack(fill="x")
            b.bind("<Enter>", lambda e, msg=d: self.set_dica(msg))
            b.bind("<Leave>", self.limpar_dica)

    def show_sidebar(self, event=None):
        if not self.sidebar_visible:
            self.sidebar_frame.grid(); self.sidebar_frame.lift(); self.sidebar_visible = True

    def check_mouse_position(self, event):
        if self.sidebar_visible:
            x = self.winfo_pointerx() - self.winfo_rootx()
            if x > 250: self.sidebar_frame.grid_remove(); self.sidebar_visible = False

    def select_frame_by_name(self, name):
        for f in [self.frame_ocr, self.frame_merge, self.frame_tarja]: f.grid_forget()
        if name == "tarja": self.frame_tarja.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        elif name == "merge": self.frame_merge.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        elif name == "ocr": self.frame_ocr.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    def criar_lista_ui(self, parent, h=5):
        f = tk.Frame(parent, bg="white", bd=1, relief="solid")
        lb = tk.Listbox(f, height=h, font=("Segoe UI", 10), borderwidth=0, highlightthickness=0, selectbackground="#3B8ED0")
        lb.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        f.pack(fill="x", pady=5); return lb

    # --- ABAS ---
    def setup_tarja_tab(self):
        ctk.CTkLabel(self.frame_tarja, text="Ocultação de Dados", font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(10, 20))
        ctk.CTkButton(self.frame_tarja, text="📂 Selecionar Documentos", font=("Segoe UI", 12, "bold"), command=lambda: self.importar("tarja", self.lst_tarja)).pack(anchor="w", pady=5)
        self.lst_tarja = self.criar_lista_ui(self.frame_tarja)
        self.lst_tarja.bind('<<ListboxSelect>>', lambda e: self.preview_selecao(e, "tarja"))
        f_btns = ctk.CTkFrame(self.frame_tarja, fg_color="transparent"); f_btns.pack(anchor="w", pady=5)
        ctk.CTkButton(f_btns, text="Remover", fg_color="#dc3545", width=100, font=("Segoe UI", 12, "bold"), command=lambda: self.excluir_um(self.lst_tarja, "tarja")).pack(side="left", padx=(0, 5))
        ctk.CTkButton(f_btns, text="Limpar Lista", fg_color="gray", width=100, font=("Segoe UI", 12, "bold"), command=lambda: self.limpar_aba("tarja", self.lst_tarja)).pack(side="left")
        f_filtros = ctk.CTkFrame(self.frame_tarja, corner_radius=8, fg_color="#f8f9fa", border_width=1, border_color="#dee2e6"); f_filtros.pack(fill="x", pady=15)
        self.vars = {"Identificação (CPF / RG)": ctk.BooleanVar(value=True), "Financeiros": ctk.BooleanVar(value=True), "Contatos (Email / Tel)": ctk.BooleanVar(value=True), "Localização": ctk.BooleanVar(value=True), "Nomes (IA)": ctk.BooleanVar(value=True)}
        for l, v in self.vars.items(): ctk.CTkCheckBox(f_filtros, text=l, variable=v).pack(anchor="w", padx=20, pady=5)
        self.var_borrao_auto = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(f_filtros, text="📷 Borrão Automático (Rostos)", variable=self.var_borrao_auto, text_color="#0dcaf0", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=20, pady=10)
        ctk.CTkButton(self.frame_tarja, text="🔍 INICIAR MAPEAMENTO", command=self.thread_analise, font=("Segoe UI", 13, "bold"), height=35).pack(fill="x")
        self.lbl_tarja_status = ctk.CTkLabel(self.frame_tarja, text="Aguardando..."); self.lbl_tarja_status.pack()
        self.caixa_rev = ctk.CTkTextbox(self.frame_tarja, height=120, font=("Consolas", 13)); self.caixa_rev.pack(fill="x", pady=5)
        ctk.CTkButton(self.frame_tarja, text="🔒 EXECUTAR TARJAMENTO", fg_color="#198754", font=("Segoe UI", 14, "bold"), height=45, command=self.thread_tarjar).pack(fill="x", pady=10)
        self.lst_prontos_tarja = self.criar_lista_ui(self.frame_tarja, h=3)
        self.lst_prontos_tarja.bind('<<ListboxSelect>>', lambda e: self.preview_pronto(e, "tarja"))

    def setup_merge_tab(self):
        ctk.CTkLabel(self.frame_merge, text="Unificar Arquivos", font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(10, 20))
        ctk.CTkButton(self.frame_merge, text="📂 Adicionar PDFs", font=("Segoe UI", 12, "bold"), command=lambda: self.importar("merge", self.lst_merge)).pack(anchor="w")
        f_ajuste = ctk.CTkFrame(self.frame_merge, fg_color="transparent"); f_ajuste.pack(fill="x")
        self.lst_merge = self.criar_lista_ui(f_ajuste, h=12)
        self.lst_merge.bind('<<ListboxSelect>>', lambda e: self.preview_selecao(e, "merge"))
        f_setas = ctk.CTkFrame(f_ajuste, fg_color="transparent"); f_setas.pack(side="right", padx=5)
        ctk.CTkButton(f_setas, text="▲", width=40, command=lambda: self.mover_item(-1)).pack(pady=2); ctk.CTkButton(f_setas, text="▼", width=40, command=lambda: self.mover_item(1)).pack(pady=2)
        f_btns = ctk.CTkFrame(self.frame_merge, fg_color="transparent"); f_btns.pack(anchor="w", pady=5)
        ctk.CTkButton(f_btns, text="Remover", fg_color="#dc3545", width=100, font=("Segoe UI", 12, "bold"), command=lambda: self.excluir_um(self.lst_merge, "merge")).pack(side="left", padx=(0, 5))
        ctk.CTkButton(f_btns, text="Limpar Lista", fg_color="gray", width=100, font=("Segoe UI", 12, "bold"), command=lambda: self.limpar_aba("merge", self.lst_merge)).pack(side="left")
        ctk.CTkButton(self.frame_merge, text="🔗 UNIFICAR DOCUMENTOS", font=("Segoe UI", 14, "bold"), height=45, command=self.thread_merge).pack(fill="x", pady=20)
        self.lst_prontos_merge = self.criar_lista_ui(self.frame_merge); self.lst_prontos_merge.bind('<<ListboxSelect>>', lambda e: self.preview_pronto(e, "merge"))

    def setup_ocr_tab(self):
        ctk.CTkLabel(self.frame_ocr, text="Conversão para Texto", font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(10, 20))
        ctk.CTkButton(self.frame_ocr, text="📂 Selecionar Arquivos", font=("Segoe UI", 12, "bold"), command=lambda: self.importar("ocr", self.lst_ocr)).pack(anchor="w", pady=5)
        self.lst_ocr = self.criar_lista_ui(self.frame_ocr)
        self.lst_ocr.bind('<<ListboxSelect>>', lambda e: self.preview_selecao(e, "ocr"))
        f_btns = ctk.CTkFrame(self.frame_ocr, fg_color="transparent"); f_btns.pack(anchor="w", pady=5)
        ctk.CTkButton(f_btns, text="Remover", fg_color="#dc3545", width=100, font=("Segoe UI", 12, "bold"), command=lambda: self.excluir_um(self.lst_ocr, "ocr")).pack(side="left", padx=(0, 5))
        ctk.CTkButton(f_btns, text="Limpar Lista", fg_color="gray", width=100, font=("Segoe UI", 12, "bold"), command=lambda: self.limpar_aba("ocr", self.lst_ocr)).pack(side="left")
        ctk.CTkButton(self.frame_ocr, text="⚙️ INICIAR CONVERSÃO", fg_color="#0dcaf0", text_color="black", font=("Segoe UI", 14, "bold"), height=45, command=self.thread_ocr).pack(fill="x", pady=20)
        self.lbl_ocr_status = ctk.CTkLabel(self.frame_ocr, text=""); self.lbl_ocr_status.pack()
        self.lst_prontos_ocr = self.criar_lista_ui(self.frame_ocr); self.lst_prontos_ocr.bind('<<ListboxSelect>>', lambda e: self.preview_pronto(e, "ocr"))

    def setup_viewer(self):
        v_tb = ctk.CTkFrame(self.frame_view, height=50, fg_color="#e9ecef"); v_tb.pack(fill="x")
        ctk.CTkButton(v_tb, text="📂 Abrir", width=70, command=self.abrir_avulso).pack(side="left", padx=5)
        ctk.CTkButton(v_tb, text="🖨️ Imprimir", width=100, command=self.imprimir_avulso).pack(side="left", padx=2)
        ctk.CTkButton(v_tb, text="📠 Digitalizar", width=110, command=self.chamar_scanner).pack(side="left", padx=2)
        ctk.CTkButton(v_tb, text="💾 Salvar", width=90, fg_color="#0d6efd", command=self.salvar_manual).pack(side="right", padx=10)
        ctk.CTkButton(v_tb, text="↩️ Desfazer", width=110, fg_color="#ffc107", text_color="black", command=self.desfazer_manual).pack(side="right", padx=5)
        self.seg_modo = ctk.CTkSegmentedButton(v_tb, values=["Desativado", "⬛ Tarja", "💧 Borrão"], command=self.mudar_modo_edicao)
        self.seg_modo.set("Desativado"); self.seg_modo.pack(side="right", padx=15)
        c_nav = ctk.CTkFrame(self.frame_view, fg_color="transparent"); c_nav.pack(pady=5, fill="x")
        ctk.CTkButton(c_nav, text="◀", width=30, command=self.pag_ant).pack(side="left", padx=10); self.lbl_pag = ctk.CTkLabel(c_nav, text="Pág: 0/0", font=("Segoe UI", 12, "bold")); self.lbl_pag.pack(side="left", padx=10); ctk.CTkButton(c_nav, text="▶", width=30, command=self.pag_prox).pack(side="left", padx=10)
        self.z_scale = ctk.CTkSlider(c_nav, from_=0.1, to=2.0, command=self.att_zoom, width=150); self.z_scale.set(0.8); self.z_scale.pack(side="right", padx=20)
        
        f_can = tk.Frame(self.frame_view, bg="gray"); f_can.pack(expand=True, fill="both", padx=5, pady=5)
        
        # Adicionando scrollbars para garantir que a imagem grande seja desenhada
        self.v_scroll = ttk.Scrollbar(f_can, orient="vertical")
        self.h_scroll = ttk.Scrollbar(f_can, orient="horizontal")
        
        self.can_view = tk.Canvas(f_can, bg="#cccccc", cursor="crosshair", highlightthickness=0, 
                                  yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        
        self.v_scroll.config(command=self.can_view.yview)
        self.h_scroll.config(command=self.can_view.xview)
        
        self.v_scroll.pack(side="right", fill="y")
        self.h_scroll.pack(side="bottom", fill="x")
        self.can_view.pack(side="left", expand=True, fill="both")
        
        self.can_view.bind("<ButtonPress-1>", self.on_press); self.can_view.bind("<B1-Motion>", self.on_drag); self.can_view.bind("<ButtonRelease-1>", self.on_release)

    # --- PROCESSAMENTO CORE ---
    def auto_borrar_rostos_pdf(self, path):
        try:
            doc = fitz.open(path)
            
            # 1. Tenta carregar o cascade nativo da biblioteca cv2
            caminho_xml = None
            try:
                if hasattr(cv2, 'data') and hasattr(cv2.data, 'haarcascades'):
                    caminho_xml = os.path.join(cv2.data.haarcascades, 'haarcascade_frontalface_default.xml')
            except: pass

            # 2. Se falhar, tenta usar o caminho relativo
            if not caminho_xml or not os.path.exists(caminho_xml):
                caminho_xml = 'haarcascade_frontalface_default.xml'
            
            # Carrega o classificador
            detector = cv2.CascadeClassifier(caminho_xml)
            
            # Verifica se carregou com sucesso
            if detector.empty():
                messagebox.showwarning("Aviso", "Não foi possível carregar o modelo de rostos (CascadeClassifier vazio).\nO tarjamento continuará, mas os rostos não serão borrados.")
                doc.close()
                return

            for p in doc:
                pix = p.get_pixmap(colorspace=fitz.csRGB)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)
                faces = detector.detectMultiScale(cv2.cvtColor(img, cv2.COLOR_RGB2GRAY), 1.1, 5)
                for (x, y, w, h) in faces:
                    rect = fitz.Rect(x*(p.rect.width/pix.width), y*(p.rect.height/pix.height), (x+w)*(p.rect.width/pix.width), (y+h)*(p.rect.height/pix.height))
                    c_pix = p.get_pixmap(clip=rect, colorspace=fitz.csRGB)
                    borr = cv2.GaussianBlur(cv2.cvtColor(np.frombuffer(c_pix.samples, dtype=np.uint8).reshape(c_pix.h, c_pix.w, 3), cv2.COLOR_RGB2BGR), (61, 61), 40)
                    p.insert_image(rect, stream=cv2.imencode('.png', borr)[1].tobytes())
            doc.saveIncr()
            doc.close()
        except Exception as e:
            messagebox.showwarning("Aviso", f"Ocorreu um erro ao tentar borrar rostos:\n{e}\n\nO processo de tarjamento continuará sem o borrão de rostos.")

    def thread_analise(self): threading.Thread(target=self.analisar, daemon=True).start()
    def analisar(self):
        if not self.dados["tarja"]["entrada"] or not nlp: return
        self.after(0, lambda: self.lbl_tarja_status.configure(text="⏳ Mapeando...", text_color="#ffc107"))
        termos = set()
        
        # Conectando os checkboxes da interface para ativar os filtros dinamicamente
        chk_id = self.vars["Identificação (CPF / RG)"].get()
        chk_fin = self.vars["Financeiros"].get()
        chk_ctt = self.vars["Contatos (Email / Tel)"].get()
        chk_loc = self.vars["Localização"].get() 
        chk_ia = self.vars["Nomes (IA)"].get()

        for cam in self.dados["tarja"]["entrada"]:
            try:
                with fitz.open(cam) as d:
                    for p in d:
                        txt = p.get_text("text")
                        # Cria uma versão sem quebras para o regex achar dados de tabela quebrados
                        txt_limpo = txt.replace("\n", " ").replace("\r", " ")
                        txt_limpo = " ".join(txt_limpo.split())

                        # 1. Contatos (Email e Telefones)
                        if chk_ctt:
                            termos.update(re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', txt_limpo))
                            termos.update(re.findall(r'(?:\+?55\s?)?(?:\(?\d{2}\)?[\s-]?)?\d{4,5}[\s-]?\d{4}', txt_limpo))

                        # 2. Identificação (CPF, CNPJ, RG, Processos)
                        if chk_id:
                            termos.update(re.findall(r'\b\d{3}\.\d{3}\.\d{3}-\d{2}\b', txt_limpo))
                            termos.update(re.findall(r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b', txt_limpo))
                            termos.update(re.findall(r'\b[A-Z]{2}-\d{1,2}\.\d{3}\.\d{3}\b', txt_limpo))
                            termos.update(re.findall(r'\b\d{3,}[\.\-\/\s]+\d+[\.\-\/\s]*\d*\b', txt_limpo))

                        # 3. Financeiros (Contas e Cartões)
                        if chk_fin:
                            termos.update(re.findall(r'\b(?:\d{4}|[Xx]{4})[\-\s]?(?:\d{4}|[Xx]{4})[\-\s]?(?:\d{4}|[Xx]{4})[\-\s]?\d{4}\b', txt_limpo))
                            termos.update(re.findall(r'\b\d{4,10}[\-\s][0-9Xx]\b', txt_limpo))
                            termos.update(re.findall(r'(?i)Agência\s\d+', txt_limpo))
                            termos.update(re.findall(r'(?i)Conta Corrente\s\d+[\-\s]?[0-9Xx]?', txt_limpo))

                        # 4. Localização (Placas, CEP, Endereço básico)
                        if chk_loc:
                            termos.update(re.findall(r'\b[A-Za-z]{3}[\-\s]?[0-9][A-Za-z0-9][0-9]{2}\b', txt_limpo))
                            termos.update(re.findall(r'\b\d{5}-\d{3}\b', txt_limpo))
                            termos.update(re.findall(r'(?i)(?:Rua|Avenida|Travessa)\s+[A-Za-zÀ-ÖØ-öø-ÿ\s]+,\s*(?:nº\s*)?\d+', txt_limpo))

                        # 5. Nomes (Motor de IA)
                        if chk_ia:
                            for ent in nlp(txt).ents:
                                if ent.label_ == "PER" and len(ent.text) > 3: 
                                    termos.add(ent.text.strip())

            except Exception as e:
                self.after(0, lambda err=e: messagebox.showerror("Erro na Análise", f"Falha:\n{err}"))
                
        self.caixa_rev.delete("0.0", "end")
        # Garante que os termos vão pra caixa já limpinhos e sem duplicatas
        [self.caixa_rev.insert("end", f"{t.strip()}\n") for t in sorted(list(termos)) if t.strip()]
        self.after(0, lambda: self.lbl_tarja_status.configure(text="✅ Concluído!", text_color="#28a745"))

    def thread_tarjar(self): threading.Thread(target=self.tarjar, daemon=True).start()
    def tarjar(self):
        if not self.dados["tarja"]["entrada"]: return
        
        # Chama a tela de carregamento (tem que ser pelo after para não travar a thread)
        self.after(0, lambda: self.mostrar_carregamento("Aplicando Tarjas e Borrões..."))
        
        # Pega os termos da caixa, garante que não são vazios e ordena do maior pro menor
        termos_brutos = self.caixa_rev.get("0.0", "end").split("\n")
        termos = [t.strip() for t in termos_brutos if t.strip()]
        termos = sorted(list(set(termos)), key=len, reverse=True)
        
        for cam in self.dados["tarja"]["entrada"]:
            try:
                saida = os.path.splitext(cam)[0] + "_TARJADO.pdf"; doc = fitz.open(cam)
                for p in doc:
                    for t in termos:
                        try:
                            # Tenta achar o texto na página
                            marcacoes = p.search_for(t)
                            if marcacoes:
                                for inst in marcacoes: 
                                    try:
                                        # Tenta pintar o retângulo
                                        p.add_redact_annot(inst, fill=(0,0,0))
                                    except Exception:
                                        # Se o retângulo for inválido, pula essa anotação específica
                                        continue
                                # Aplica TODAS as redações que deram certo para ESTE termo
                                p.apply_redactions()
                        except Exception:
                            # Se a busca em si travar, pula pro próximo termo
                            continue
                            
                doc.save(saida); doc.close()
                if self.var_borrao_auto.get(): self.auto_borrar_rostos_pdf(saida)
                if saida not in self.dados["tarja"]["prontos"]: self.dados["tarja"]["prontos"].append(saida)
                self.after(0, lambda: self.atualizar_lb(self.lst_prontos_tarja, self.dados["tarja"]["prontos"]))
            except Exception as e:
                self.after(0, lambda err=e: messagebox.showerror("Erro ao Tarjar", f"Falha no documento:\n{err}"))
                
        # Fecha a tela de carregamento no final
        self.after(0, self.fechar_carregamento)
        self.after(0, lambda: messagebox.showinfo("Concluído", "Processo de tarjamento finalizado!"))

    def thread_ocr(self): threading.Thread(target=self.exec_ocr, daemon=True).start()
    def exec_ocr(self):
        if not self.dados["ocr"]["entrada"]: return
        
        self.after(0, lambda: self.mostrar_carregamento("Convertendo Imagens para Texto (OCR)..."))
        
        for cam in self.dados["ocr"]["entrada"]:
            try:
                doc = fitz.open(cam); res_p = fitz.open(); tot = len(doc)
                for i in range(tot):
                    pix = doc[i].get_pixmap(dpi=200); res = pytesseract.image_to_pdf_or_hocr(Image.frombytes("RGB", [pix.width, pix.height], pix.samples), extension='pdf', lang='por', timeout=60)
                    with fitz.open("pdf", res) as p: res_p.insert_pdf(p)
                    self.after(0, lambda v=((i+1)/tot)*100: self.progress.configure(value=v))
                saida = f"{os.path.splitext(cam)[0]}_PRONTO.pdf"; res_p.save(saida); res_p.close(); doc.close()
                self.dados["ocr"]["prontos"].append(saida); self.after(0, lambda: self.atualizar_lb(self.lst_prontos_ocr, self.dados["ocr"]["prontos"]))
            except Exception as e:
                self.after(0, lambda err=e: messagebox.showerror("Erro no OCR", f"O OCR falhou:\n{err}"))
                
        self.after(0, self.fechar_carregamento)
        self.after(0, lambda: messagebox.showinfo("Concluído", "Conversão OCR finalizada!"))

    def thread_merge(self): threading.Thread(target=self.exec_merge, daemon=True).start()
    def exec_merge(self):
        if not self.dados["merge"]["entrada"]: return
        
        self.after(0, lambda: self.mostrar_carregamento("Unificando Documentos PDF..."))
        
        try:
            res = fitz.open(); [res.insert_pdf(fitz.open(f)) for f in self.dados["merge"]["entrada"]]
            saida = os.path.join(os.path.dirname(self.dados["merge"]["entrada"][0]), "UNIFICADO.pdf")
            res.save(saida); res.close(); self.dados["merge"]["prontos"].append(saida); self.after(0, lambda: self.atualizar_lb(self.lst_prontos_merge, self.dados["merge"]["prontos"]))
        except Exception as e:
             self.after(0, lambda err=e: messagebox.showerror("Erro ao Unificar", f"Falha:\n{err}"))
             
        self.after(0, self.fechar_carregamento)
        self.after(0, lambda: messagebox.showinfo("Concluído", "Unificação finalizada!"))

    # --- HANDLERS ---
    def carregar_pdf(self, c):
        if not c or not os.path.exists(c): return
        if self.doc_aberto: self.doc_aberto.close()
        self.caminho_atual = c; self.doc_aberto = fitz.open(c); self.pagina_atual = 0; self.renderizar()

    def renderizar(self):
        if not self.doc_aberto: return
        pix = self.doc_aberto.load_page(self.pagina_atual).get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level))
        
        # Cria a imagem e salva a referência
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.img_tk = ImageTk.PhotoImage(img) 
        
        self.can_view.delete("all")
        self.can_view.create_image(0, 0, anchor="nw", image=self.img_tk)
        self.can_view.config(scrollregion=self.can_view.bbox("all")) 
        
        self.lbl_pag.configure(text=f"Pág: {self.pagina_atual+1}/{len(self.doc_aberto)}")

    def on_press(self, e):
        if self.modo_edicao == "Desativado" or not self.doc_aberto: return
        self.start_x, self.start_y = self.can_view.canvasx(e.x), self.can_view.canvasy(e.y)
        self.rect_id = self.can_view.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", dash=(4,4))

    def on_drag(self, e):
        if self.rect_id: self.can_view.coords(self.rect_id, self.start_x, self.start_y, self.can_view.canvasx(e.x), self.can_view.canvasy(e.y))

    def on_release(self, e):
        if not self.rect_id: return
        self.historico_estados.append(self.doc_aberto.tobytes()); p = self.doc_aberto.load_page(self.pagina_atual)
        x0, x1 = min(self.start_x, self.can_view.canvasx(e.x))/self.zoom_level, max(self.start_x, self.can_view.canvasx(e.x))/self.zoom_level
        y0, y1 = min(self.start_y, self.can_view.canvasy(e.y))/self.zoom_level, max(self.start_y, self.can_view.canvasy(e.y))/self.zoom_level
        rect = fitz.Rect(x0, y0, x1, y1)
        if "Tarja" in self.modo_edicao: p.add_redact_annot(rect, fill=(0,0,0)); p.apply_redactions()
        elif "Borrão" in self.modo_edicao:
            pix = p.get_pixmap(clip=rect, colorspace=fitz.csRGB); borr = cv2.GaussianBlur(cv2.cvtColor(np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3), cv2.COLOR_RGB2BGR), (61, 61), 40)
            p.insert_image(rect, stream=cv2.imencode('.png', borr)[1].tobytes())
        self.can_view.delete(self.rect_id); self.rect_id = None; self.renderizar()

    def mover_item(self, d):
        idx = self.lst_merge.curselection()
        if idx:
            i = idx[0]; j = i + d
            if 0 <= j < len(self.dados["merge"]["entrada"]):
                self.dados["merge"]["entrada"][i], self.dados["merge"]["entrada"][j] = self.dados["merge"]["entrada"][j], self.dados["merge"]["entrada"][i]
                self.atualizar_lb(self.lst_merge, self.dados["merge"]["entrada"]); self.lst_merge.select_set(j); self.carregar_pdf(self.dados["merge"]["entrada"][j])

    def importar(self, s, lb):
        f = filedialog.askopenfilenames(); 
        if f: self.dados[s]["entrada"].extend(f); self.atualizar_lb(lb, self.dados[s]["entrada"]); self.carregar_pdf(f[0])
    def atualizar_lb(self, lb, l): lb.delete(0, tk.END); [lb.insert(tk.END, os.path.basename(x)) for x in l]
    def preview_selecao(self, e, s):
        idx = e.widget.curselection(); 
        if idx: self.carregar_pdf(self.dados[s]["entrada"][idx[0]])
    def preview_pronto(self, e, s):
        idx = e.widget.curselection(); 
        if idx: self.carregar_pdf(self.dados[s]["prontos"][idx[0]])
    def excluir_um(self, lb, s):
        idx = lb.curselection(); 
        if idx: self.dados[s]["entrada"].pop(idx[0]); self.atualizar_lb(lb, self.dados[s]["entrada"])
    def salvar_manual(self): 
        f = filedialog.asksaveasfilename(defaultextension=".pdf"); 
        if f: self.doc_aberto.save(f); messagebox.showinfo("Sucesso", "Edição gravada!")
    def abrir_avulso(self): self.carregar_pdf(filedialog.askopenfilename())
    def imprimir_avulso(self):
        if self.caminho_atual:
            if SISTEMA == "Windows": os.startfile(self.caminho_atual)
            else: os.system(f"open '{self.caminho_atual}'")
    def chamar_scanner(self):
        if SISTEMA == "Windows": os.system("start wiaacmgr")
        else: messagebox.showinfo("Scanner", "Utilize o 'Captura de Imagem' nativo.")
    def att_zoom(self, v): self.zoom_level = float(v); self.renderizar()
    def pag_ant(self): 
        if self.pagina_atual > 0: self.pagina_atual -= 1; self.renderizar()
    def pag_prox(self): 
        if self.doc_aberto and self.pagina_atual < len(self.doc_aberto)-1: self.pagina_atual += 1; self.renderizar()
    def limpar_aba(self, s, lb): self.dados[s]["entrada"] = []; self.atualizar_lb(lb, [])
    def desfazer_manual(self):
        if self.historico_estados:
            self.doc_aberto.close(); self.doc_aberto = fitz.open("pdf", self.historico_estados.pop()); self.renderizar()
    def mudar_modo_edicao(self, v): self.modo_edicao = v

if __name__ == "__main__":
    app = TrataDocApp(); app.mainloop()