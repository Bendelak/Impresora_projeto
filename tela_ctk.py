"""Interface gráfica (customtkinter) do sistema de impressão — Posto 44.

Requer:  pip install customtkinter
"""

import csv
import json
import logging
import queue
import secrets
import threading
import time
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

from Impresora_projeto.main import (
    NOME_IMPRESSORA,
    ErroImpressora,
    Impressora,
    calcular_carta_frete,
    calcular_operacoes,
    formatar_numero,
    formatar_valor,
    listar_impressoras,
    parse_numero,
    parse_operacao,
    parse_valor,
)

# ============================================================
# CONFIGURAÇÃO
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
PASTA_CSV = BASE_DIR / "CSV"
LOG_VALES = BASE_DIR / "logs" / "vales.csv"
ARQUIVO_CONFIG = BASE_DIR / "config.json"

MAX_SENHAS = 500
MAX_COPIAS_WIFI = 20
LIMITE_TEXTO = 1000
PAUSA_ENTRE_IMPRESSOES = 0.3

# Texto da coluna "tempo" do arquivo de importação do GDC
TEMPO_BANHO = "7 min e 0 seg"

# --- Tela "Sobre" (edite à vontade) ---
NOME_PROJETO = "Sistema de Impressão"
VERSAO = "2.1"
AUTOR = "Warlen kauã"        

COMENTARIOS = """
O QUE O SISTEMA FAZ
• Texto: imprime linhas de texto com cabeçalho e tamanho de fonte à escolha.
• Banho: gera senhas únicas de 6 dígitos, imprime uma por vez e salva o CSV no formato de importação do GDC (pasta CSV).
• Vale: imprime vales e registra cada um em logs/vales.csv.
• Calculadora: fita de cálculo com + - * / impressa no cupom.
• Carta Frete: apura a quebra de peso e o saldo (acessada pela Calculadora).
• WiFi: imprime nome da rede, senha e QR Code de conexão.
• Config: escolhe a impressora e as fontes padrão (salvas em config.json).

CARTA FRETE — COMO É CALCULADO (baseado na planilha CALCULO_FRETE_MANUAL.xlsx)
• Quebra = peso líquido - peso da nota (negativo = faltou peso).
• Isenção de quebra: informada em % do peso da nota ou direto em kg.
• Se a quebra (sem sinal) passar da isenção: valor da quebra = (quebra - isenção) x valor por kg.
• Valor saldo = valor total - adiantamento 1 - adiantamento 2.
• Valor a receber = valor saldo - valor da quebra.

ARQUIVOS GERADOS
• config.json (preferências), CSV/ (senhas de banho), logs/vales.csv e erros.log.

TECNOLOGIAS
• Python, customtkinter (interface) e pywin32 (impressão ESC/POS na Epson TM-T20X).
"""

PLACEHOLDER_TEXTO = (
    "Digite o texto que será impresso…\n"
    "(cada linha vira uma impressão separada)"
)

CONFIG_PADRAO = {
    "impressora": NOME_IMPRESSORA,
    "cabecalho": "POSTO 44",
    "fonte_texto": "Média (2x2)",
    "fonte_banho": "Média (2x2)",
    "wifi_titulo": "WIFI",
    "wifi_rede": "Clientes Posto 44",
    "wifi_senha": "posto442025",
}

# --- cores (tema escuro com vermelho do Posto 44) ---
BG = "#0D141F"
SIDEBAR = "#101823"
CARD = "#121B28"
CARD_HOVER = "#182335"
BORDA = "#243142"
INPUT = "#17212F"
TEXTO = "#E6EAF0"
TEXTO_SEC = "#8A97A8"
HEADER = "#A30D1B"
VERMELHO = "#CF1424"
VERMELHO_HOVER = "#E01F31"
NAV_ATIVO = "#C0121F"
NAV_HOVER = "#1C2736"
CINZA_OCUPADO = "#5B6574"
VERDE = "#22C55E"
ERRO = "#EF4444"
AMARELO = "#FBBF24"

STATUS_COR = {"ok": VERDE, "ocupado": AMARELO, "erro": ERRO}


# ============================================================
# FUNÇÕES AUXILIARES (sem interface)
# ============================================================

def carregar_config():
    config = dict(CONFIG_PADRAO)
    try:
        config.update(json.loads(ARQUIVO_CONFIG.read_text(encoding="utf-8")))
    except FileNotFoundError:
        pass
    except (OSError, ValueError):
        logging.warning("config.json inválido; usando padrões", exc_info=True)
    return config


def salvar_config(config):
    try:
        ARQUIVO_CONFIG.write_text(
            json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except OSError:
        logging.warning("Não foi possível salvar config.json", exc_info=True)


def gerar_senhas(quantidade):
    """Senhas únicas de 6 dígitos (sem zero à esquerda), geradas com `secrets`."""
    sorteio = secrets.SystemRandom()
    return [str(n) for n in sorteio.sample(range(100000, 1000000), quantidade)]


def salvar_csv_senhas(senhas):
    PASTA_CSV.mkdir(parents=True, exist_ok=True)
    arquivo = PASTA_CSV / f"senhas_{datetime.now():%d-%m-%Y_%H-%M-%S}.csv"

    # Formato de importação do GDC (igual ao criarsenhas.py): sem cabeçalho,
    # separado por ";" e em UTF-8 SEM BOM (o BOM estragaria a primeira célula).
    with open(arquivo, "w", newline="", encoding="utf-8") as f:
        escritor = csv.writer(f, delimiter=";")
        for numero, senha in enumerate(senhas, start=1):
            escritor.writerow([numero, senha, 0, TEMPO_BANHO, 1])

    return arquivo


def registrar_vale(valor):
    """Guarda cada vale impresso (auditoria). Falha aqui não impede a impressão."""
    try:
        LOG_VALES.parent.mkdir(parents=True, exist_ok=True)
        novo = not LOG_VALES.exists()

        with open(LOG_VALES, "a", newline="", encoding="utf-8") as f:
            escritor = csv.writer(f, delimiter=";")
            if novo:
                escritor.writerow(["Data/Hora", "Valor"])
            escritor.writerow([
                datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                f"{valor:.2f}".replace(".", ","),
            ])
    except OSError:
        logging.warning("Não foi possível registrar o vale no log", exc_info=True)


# ============================================================
# TELA
# ============================================================

class Tela:

    def __init__(self):
        self.config = carregar_config()
        self.imp = Impressora(self.config["impressora"])

        ctk.set_appearance_mode("dark")
        self.janela = ctk.CTk(fg_color=BG)

        self.f = {
            "logo": ctk.CTkFont(family="Segoe UI", size=40, weight="bold"),
            "cabecalho": ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            "titulo": ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            "negrito": ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            "normal": ctk.CTkFont(family="Segoe UI", size=14),
            "pequena": ctk.CTkFont(family="Segoe UI", size=12),
            "nav": ctk.CTkFont(family="Segoe UI", size=16),
            "botao": ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            "resultado": ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            "icone": ctk.CTkFont(family="Segoe UI Emoji", size=30),
        }

        tamanhos = self.imp.tamanhos
        padrao = CONFIG_PADRAO["fonte_texto"]
        self.cabecalho_var = ctk.StringVar(value=self.config["cabecalho"])
        self.fonte_texto_var = ctk.StringVar(
            value=self.config["fonte_texto"] if self.config["fonte_texto"] in tamanhos else padrao
        )
        self.fonte_banho_var = ctk.StringVar(
            value=self.config["fonte_banho"] if self.config["fonte_banho"] in tamanhos else padrao
        )

        self.fila = queue.Queue()     # worker -> interface
        self.botoes = []              # botões de imprimir (desativam durante o trabalho)
        self.paginas = {}
        self.nav = {}
        self.ocupado = False
        self._avanco = None
        self._detalhe_falha = ""

        self._configurar_janela()
        self._criar_cabecalho()
        self._criar_corpo()
        self._criar_rodape()

        self.janela.report_callback_exception = self._erro_inesperado
        self.janela.protocol("WM_DELETE_WINDOW", self._fechar)
        self.janela.after(100, self._processar_fila)

        self._mostrar("inicio")
        self._atualizar_relogio()
        self._verificar_impressora()

    def executar(self):
        self.janela.mainloop()

    # --------------------------------------------------------
    # Janela, cabeçalho, corpo e rodapé
    # --------------------------------------------------------

    def _configurar_janela(self):
        self.janela.title("Posto 44 — Sistema de Impressão")

        largura, altura = 1080, 700
        pos_x = (self.janela.winfo_screenwidth() - largura) // 2
        pos_y = max(0, (self.janela.winfo_screenheight() - altura) // 2 - 20)

        self.janela.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
        self.janela.minsize(980, 640)

        self.janela.grid_columnconfigure(0, weight=1)
        self.janela.grid_rowconfigure(1, weight=1)

    def _criar_cabecalho(self):
        topo = ctk.CTkFrame(self.janela, fg_color=HEADER, corner_radius=0, height=88)
        topo.grid(row=0, column=0, sticky="ew")
        topo.pack_propagate(False)

        ctk.CTkLabel(
            topo, text="⛽ POSTO 44", font=self.f["logo"], text_color="white"
        ).pack(side="left", padx=(28, 18))

        ctk.CTkFrame(topo, width=1, height=50, fg_color="#C9505C").pack(side="left")

        ctk.CTkLabel(
            topo,
            text="🖨  SISTEMA DE IMPRESSÃO",
            font=self.f["cabecalho"],
            text_color="white",
        ).pack(side="left", padx=22)

        # estado da impressora (canto direito)
        direita = ctk.CTkFrame(topo, fg_color="transparent")
        direita.pack(side="right", padx=28)

        self.ponto_impressora = ctk.CTkLabel(
            direita, text="●", font=ctk.CTkFont(size=18), text_color=AMARELO
        )
        self.ponto_impressora.pack(side="left", padx=(0, 10))

        textos = ctk.CTkFrame(direita, fg_color="transparent")
        textos.pack(side="left")

        ctk.CTkLabel(
            textos, text="Impressora", font=self.f["pequena"], text_color="#F2C4C9"
        ).pack(anchor="w")

        self.nome_impressora = ctk.CTkLabel(
            textos, text=self.imp.nome, font=self.f["negrito"], text_color="white"
        )
        self.nome_impressora.pack(anchor="w")

    def _criar_corpo(self):
        corpo = ctk.CTkFrame(self.janela, fg_color="transparent", corner_radius=0)
        corpo.grid(row=1, column=0, sticky="nsew")
        corpo.grid_columnconfigure(1, weight=1)
        corpo.grid_rowconfigure(0, weight=1)

        # --- menu lateral ---
        lateral = ctk.CTkFrame(corpo, fg_color=SIDEBAR, corner_radius=0, width=250)
        lateral.grid(row=0, column=0, sticky="ns")
        lateral.pack_propagate(False)   # mantém a largura de 250 px

        itens = [
            ("inicio", "🏠", "Início"),
            ("banho", "🚿", "Banho"),
            ("vale", "🎫", "Vale"),
            ("calculadora", "🧮", "Calculadora"),
            ("wifi", "📶", "WiFi"),
            ("config", "⚙", "Config"),
        ]
        for chave, icone, texto in itens:
            botao = ctk.CTkButton(
                lateral,
                text=f"{icone}   {texto}",
                anchor="w",
                height=50,
                corner_radius=8,
                font=self.f["nav"],
                fg_color="transparent",
                hover_color=NAV_HOVER,
                text_color=TEXTO,
                command=lambda c=chave: self._mostrar(c),
            )
            botao.pack(fill="x", padx=16, pady=(18 if chave == "inicio" else 3, 3))
            self.nav[chave] = botao

        botao_sobre = ctk.CTkButton(
            lateral,
            text="ℹ   Sobre",
            anchor="w",
            height=50,
            corner_radius=8,
            font=self.f["nav"],
            fg_color="transparent",
            hover_color=NAV_HOVER,
            text_color=TEXTO,
            command=lambda: self._mostrar("sobre"),
        )
        botao_sobre.pack(side="bottom", fill="x", padx=16, pady=(3, 18))
        self.nav["sobre"] = botao_sobre

        # --- área de conteúdo ---
        self.conteudo = ctk.CTkFrame(corpo, fg_color="transparent")
        self.conteudo.grid(row=0, column=1, sticky="nsew", padx=20, pady=18)
        self.conteudo.grid_columnconfigure(0, weight=1)
        self.conteudo.grid_rowconfigure(0, weight=1)

        self._montar_pagina_inicio()
        self._montar_pagina_banho()
        self._montar_pagina_vale()
        self._montar_pagina_calculadora()
        self._montar_pagina_wifi()
        self._montar_pagina_config()
        self._montar_pagina_carta_frete()
        self._montar_pagina_sobre()

    def _criar_rodape(self):
        rodape = ctk.CTkFrame(self.janela, fg_color=SIDEBAR, corner_radius=0, height=46)
        rodape.grid(row=2, column=0, sticky="ew")
        rodape.pack_propagate(False)

        self.ponto_status = ctk.CTkLabel(
            rodape, text="●", font=ctk.CTkFont(size=18), text_color=VERDE
        )
        self.ponto_status.pack(side="left", padx=(24, 8))

        self.texto_status = ctk.CTkLabel(
            rodape, text="Sistema pronto.", font=self.f["normal"], text_color=VERDE
        )
        self.texto_status.pack(side="left")

        self.relogio = ctk.CTkLabel(
            rodape, text="", font=self.f["normal"], text_color=TEXTO_SEC
        )
        self.relogio.pack(side="right", padx=24)

    def _status(self, texto, tipo="ok"):
        cor = STATUS_COR[tipo]
        self.ponto_status.configure(text_color=cor)
        self.texto_status.configure(text=texto, text_color=cor)

    def _atualizar_relogio(self):
        self.relogio.configure(text=f"🕒  {datetime.now():%d/%m/%Y   %H:%M}")
        self.janela.after(15000, self._atualizar_relogio)

    # --------------------------------------------------------
    # Estado da impressora
    # --------------------------------------------------------

    def _verificar_impressora(self, avisar=False):
        self.nome_impressora.configure(text=self.imp.nome)
        self.ponto_impressora.configure(text_color=AMARELO)

        def trabalho():
            ok = self.imp.disponivel()
            self.fila.put(lambda: self._mostrar_impressora(ok, avisar))

        threading.Thread(target=trabalho, daemon=True).start()

    def _mostrar_impressora(self, ok, avisar=False):
        self.ponto_impressora.configure(text_color=VERDE if ok else ERRO)

        if avisar:
            if ok:
                self._status(f"Impressora '{self.imp.nome}' encontrada.")
            else:
                self._status("Impressora não encontrada.", "erro")
                messagebox.showwarning(
                    "Impressora",
                    f"Não encontrei '{self.imp.nome}' no Windows.\n"
                    "Confira o nome na lista de impressoras.",
                )

    # --------------------------------------------------------
    # Navegação
    # --------------------------------------------------------

    def _mostrar(self, chave):
        self.paginas[chave].tkraise()
        destaque = {"carta_frete": "calculadora"}.get(chave, chave)
        for nome, botao in self.nav.items():
            ativo = nome == destaque
            botao.configure(
                fg_color=NAV_ATIVO if ativo else "transparent",
                hover_color=NAV_ATIVO if ativo else NAV_HOVER,
            )

    def _nova_pagina(self, chave, icone, titulo, expandir=True, rolavel=False):
        pagina = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        pagina.grid(row=0, column=0, sticky="nsew")

        classe = ctk.CTkScrollableFrame if rolavel else ctk.CTkFrame
        cartao = classe(
            pagina, fg_color=CARD, border_width=1, border_color=BORDA, corner_radius=12
        )
        if expandir or rolavel:
            cartao.pack(fill="both", expand=True)
        else:
            cartao.pack(fill="x", anchor="n")

        ctk.CTkLabel(
            cartao, text=f"{icone}   {titulo}", font=self.f["titulo"], text_color=TEXTO
        ).pack(anchor="w", padx=24, pady=(20, 12))

        self.paginas[chave] = pagina
        return pagina, cartao

    # --------------------------------------------------------
    # Fábricas de widgets
    # --------------------------------------------------------

    def _rotulo(self, pai, texto):
        return ctk.CTkLabel(
            pai, text=texto, font=self.f["normal"], text_color=TEXTO_SEC, anchor="w"
        )

    def _dica(self, pai, texto):
        ctk.CTkLabel(
            pai, text=texto, font=self.f["pequena"], text_color=TEXTO_SEC, anchor="w"
        ).pack(anchor="w", padx=24, pady=(0, 8))

    def _entrada(self, pai, variavel=None, **opcoes):
        return ctk.CTkEntry(
            pai,
            textvariable=variavel,
            height=44,
            corner_radius=8,
            border_width=1,
            border_color=BORDA,
            fg_color=INPUT,
            text_color=TEXTO,
            placeholder_text_color=TEXTO_SEC,
            font=self.f["normal"],
            **opcoes,
        )

    def _combo(self, pai, variavel, valores, readonly=True):
        return ctk.CTkComboBox(
            pai,
            values=valores,
            variable=variavel,
            state="readonly" if readonly else "normal",
            height=44,
            corner_radius=8,
            border_width=1,
            border_color=BORDA,
            fg_color=INPUT,
            button_color=BORDA,
            button_hover_color=NAV_HOVER,
            dropdown_fg_color=CARD,
            dropdown_hover_color=NAV_HOVER,
            dropdown_text_color=TEXTO,
            text_color=TEXTO,
            font=self.f["normal"],
            dropdown_font=self.f["normal"],
        )

    def _caixa_texto(self, pai, altura):
        return ctk.CTkTextbox(
            pai,
            height=altura,
            corner_radius=8,
            border_width=1,
            border_color=BORDA,
            fg_color=INPUT,
            text_color=TEXTO,
            font=self.f["normal"],
        )

    def _botao(self, pai, texto, comando, largura=340, registrar=True):
        """Botão vermelho; se `registrar`, fica desativado enquanto imprime."""
        botao = ctk.CTkButton(
            pai,
            text=texto,
            command=comando,
            width=largura,
            height=56,
            corner_radius=8,
            fg_color=VERMELHO,
            hover_color=VERMELHO_HOVER,
            text_color="white",
            text_color_disabled="#D5D9E0",
            font=self.f["botao"],
        )
        if registrar:
            self.botoes.append(botao)
        return botao

    def _botao_secundario(self, pai, texto, comando, largura=150):
        return ctk.CTkButton(
            pai,
            text=texto,
            command=comando,
            width=largura,
            height=44,
            corner_radius=8,
            fg_color=BORDA,
            hover_color=NAV_HOVER,
            text_color=TEXTO,
            font=self.f["normal"],
        )

    # --------------------------------------------------------
    # Caixa de texto com placeholder e contador
    # --------------------------------------------------------

    def _preparar_caixa_texto(self, caixa, placeholder, contador, limite):
        caixa.tag_config("ph", foreground=TEXTO_SEC)
        estado = {"ativo": False, "contador": contador, "limite": limite}
        caixa._ph = estado

        def mostrar():
            if not estado["ativo"] and not caixa.get("1.0", "end-1c"):
                caixa.insert("1.0", placeholder, "ph")
                estado["ativo"] = True

        def esconder(_e=None):
            if estado["ativo"]:
                caixa.delete("1.0", "end")
                estado["ativo"] = False

        def atualizar(_e=None):
            if estado["ativo"]:
                return
            texto = caixa.get("1.0", "end-1c")
            if len(texto) > limite:
                caixa.delete(f"1.0+{limite}c", "end-1c")
                texto = caixa.get("1.0", "end-1c")
            contador.configure(text=f"{len(texto)}/{limite}")

        caixa.bind("<FocusIn>", esconder)
        caixa.bind("<FocusOut>", lambda _e: (atualizar(), mostrar()))
        caixa.bind("<KeyRelease>", atualizar)
        caixa.bind("<ButtonRelease-2>", atualizar)
        estado["mostrar"] = mostrar
        mostrar()

    @staticmethod
    def _ler_caixa(caixa):
        """Texto digitado (ignora o placeholder)."""
        ph = getattr(caixa, "_ph", None)
        if ph and ph["ativo"]:
            return ""
        return caixa.get("1.0", "end-1c")

    def _limpar_caixa(self, caixa):
        caixa.delete("1.0", "end")
        ph = getattr(caixa, "_ph", None)
        if ph:
            ph["ativo"] = False
            ph["contador"].configure(text=f"0/{ph['limite']}")
            if self.janela.focus_get() is not getattr(caixa, "_textbox", None):
                ph["mostrar"]()

    # --------------------------------------------------------
    # Páginas
    # --------------------------------------------------------

    def _montar_pagina_inicio(self):
        pagina, cartao = self._nova_pagina("inicio", "🖨", "Impressão de Texto")

        # atalhos (embaixo)
        linha = ctk.CTkFrame(pagina, fg_color="transparent")
        linha.pack(side="bottom", fill="x", pady=(14, 0))

        atalhos = [
            ("🚿", "#3B82F6", "Banho", "Imprimir senha de banho", "banho"),
            ("🎫", "#22C55E", "Vale", "Imprimir vale", "vale"),
            ("🧮", "#8B5CF6", "Calculadora", "Imprimir cálculos", "calculadora"),
            ("📶", "#06B6D4", "WiFi", "QR Code da rede", "wifi"),
            ("⚙", "#94A3B8", "Configurações", "Impressora e fontes", "config"),
        ]
        for coluna, (icone, cor, titulo, sub, destino) in enumerate(atalhos):
            linha.grid_columnconfigure(coluna, weight=1, uniform="atalho")
            self._atalho(linha, coluna, icone, cor, titulo, sub, destino)

        # campos
        grade = ctk.CTkFrame(cartao, fg_color="transparent")
        grade.pack(fill="x", padx=24)
        grade.grid_columnconfigure((0, 1), weight=1, uniform="campo")

        self._rotulo(grade, "Cabeçalho").grid(row=0, column=0, sticky="w")
        self._entrada(grade, self.cabecalho_var).grid(
            row=1, column=0, sticky="ew", padx=(0, 8)
        )

        self._rotulo(grade, "Fonte").grid(row=0, column=1, sticky="w", padx=(8, 0))
        self._combo(grade, self.fonte_texto_var, list(self.imp.tamanhos)).grid(
            row=1, column=1, sticky="ew", padx=(8, 0)
        )

        self.caixa_texto = self._caixa_texto(cartao, altura=110)
        self.caixa_texto.pack(fill="both", expand=True, padx=24, pady=(16, 4))

        self.contador_texto = ctk.CTkLabel(
            cartao, text=f"0/{LIMITE_TEXTO}", font=self.f["pequena"], text_color=TEXTO_SEC
        )
        self.contador_texto.pack(anchor="e", padx=24)

        self._preparar_caixa_texto(
            self.caixa_texto, PLACEHOLDER_TEXTO, self.contador_texto, LIMITE_TEXTO
        )

        self._botao(cartao, "🖨   IMPRIMIR", self.imprimir_texto).pack(pady=(8, 22))

    def _atalho(self, pai, coluna, icone, cor, titulo, subtitulo, destino):
        card = ctk.CTkFrame(
            pai, fg_color=CARD, border_width=1, border_color=BORDA, corner_radius=10
        )
        card.grid(row=0, column=coluna, sticky="nsew", padx=5)

        ctk.CTkLabel(card, text=icone, font=self.f["icone"], text_color=cor) \
            .pack(pady=(14, 2))
        ctk.CTkLabel(card, text=titulo, font=self.f["negrito"], text_color=TEXTO).pack()
        ctk.CTkLabel(
            card,
            text=subtitulo,
            font=self.f["pequena"],
            text_color=TEXTO_SEC,
            wraplength=120,
        ).pack(padx=8, pady=(2, 14))

        self._tornar_clicavel(card, lambda: self._mostrar(destino))

    @staticmethod
    def _tornar_clicavel(card, comando):
        def entrar(_e=None):
            card.configure(fg_color=CARD_HOVER)

        def sair(_e=None):
            try:
                x, y = card.winfo_pointerxy()
                sob = card.winfo_containing(x, y)
                if sob is not None and str(sob).startswith(str(card)):
                    return          # o mouse só passou para um filho do card
            except Exception:
                pass
            card.configure(fg_color=CARD)

        for widget in (card, *card.winfo_children()):
            widget.bind("<Button-1>", lambda _e: comando())
            widget.bind("<Enter>", entrar)
            widget.bind("<Leave>", sair)

    def _montar_pagina_banho(self):
        _, cartao = self._nova_pagina("banho", "🚿", "Senhas de Banho", expandir=False)
        self._dica(
            cartao,
            f"Gera senhas únicas e imprime uma por vez (1 a {MAX_SENHAS}). "
            "Os códigos também são salvos na pasta CSV.",
        )

        grade = ctk.CTkFrame(cartao, fg_color="transparent")
        grade.pack(fill="x", padx=24)
        grade.grid_columnconfigure((0, 1), weight=1, uniform="campo")

        self._rotulo(grade, "Quantidade de senhas").grid(row=0, column=0, sticky="w")
        self.entrada_banho = self._entrada(grade, placeholder_text="Ex.: 100")
        self.entrada_banho.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.entrada_banho.bind("<Return>", lambda _e: self.imprimir_banho())

        self._rotulo(grade, "Tamanho da senha").grid(row=0, column=1, sticky="w", padx=(8, 0))
        self._combo(grade, self.fonte_banho_var, list(self.imp.tamanhos)).grid(
            row=1, column=1, sticky="ew", padx=(8, 0)
        )

        self._botao(cartao, "🖨   IMPRIMIR", self.imprimir_banho).pack(pady=(26, 26))

    def _montar_pagina_vale(self):
        _, cartao = self._nova_pagina("vale", "🎫", "Vale")
        self._dica(cartao, "Um valor por linha — cada linha vira um vale. Ex.: 50 ou 50,00")

        self.caixa_vale = self._caixa_texto(cartao, altura=150)
        self.caixa_vale.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        self._botao(cartao, "🖨   IMPRIMIR", self.imprimir_vale).pack(pady=(8, 22))

    def _montar_pagina_calculadora(self):
        _, cartao = self._nova_pagina("calculadora", "🧮", "Calculadora")
        self._dica(cartao, "Um por linha. Use + - * / no início (ex.: *2, /4). Sem sinal, soma.")

        self.caixa_calculadora = self._caixa_texto(cartao, altura=130)
        self.caixa_calculadora.pack(fill="both", expand=True, padx=24, pady=(0, 6))

        self.resultado_calculadora = ctk.CTkLabel(
            cartao, text="Total: 0,00", font=self.f["resultado"], text_color=TEXTO
        )
        self.resultado_calculadora.pack(pady=6)

        linha = ctk.CTkFrame(cartao, fg_color="transparent")
        linha.pack(pady=(6, 22))
        self._botao(linha, "CALCULAR E IMPRIMIR", self.calcular, largura=300).pack(
            side="left", padx=6
        )
        self._botao_secundario(
            linha, "🚚  Carta Frete", lambda: self._mostrar("carta_frete"), largura=190
        ).pack(side="left", padx=6)

    def _montar_pagina_carta_frete(self):
        _, cartao = self._nova_pagina("carta_frete", "🚚", "Carta Frete", rolavel=True)
        self._dica(cartao, "Mesma conta da planilha de cálculo manual. "
                           "Use vírgula nos decimais.")

        grade = ctk.CTkFrame(cartao, fg_color="transparent")
        grade.pack(fill="x", padx=18)
        grade.grid_columnconfigure((0, 1, 2), weight=1, uniform="cf")

        def campo(linha, coluna, rotulo, exemplo):
            self._rotulo(grade, rotulo).grid(
                row=linha, column=coluna, sticky="w", padx=6,
                pady=(12 if linha else 0, 0),
            )
            entrada = self._entrada(grade, placeholder_text=exemplo)
            entrada.grid(row=linha + 1, column=coluna, sticky="ew", padx=6)
            entrada.bind("<Return>", lambda _e: self.calcular_frete())
            return entrada

        self.cf_nota = campo(0, 0, "Peso da nota (kg)", "Ex.: 32.050")
        self.cf_liquido = campo(0, 1, "Peso líquido (kg)", "Ex.: 31.900")
        self.cf_total = campo(0, 2, "Valor total (R$)", "Ex.: 7.948,40")
        self.cf_adiant1 = campo(2, 0, "Adiantamento 1 (R$)", "Opcional")
        self.cf_adiant2 = campo(2, 1, "Adiantamento 2 (R$)", "Opcional")
        self.cf_valor_kg = campo(2, 2, "Valor por kg da quebra (R$)", "Ex.: 1,29")
        self.cf_isencao = campo(4, 0, "Isenção de quebra", "Ex.: 0,3")

        self._rotulo(grade, "Isenção informada em:").grid(
            row=4, column=1, sticky="w", padx=6, pady=(12, 0)
        )
        self.cf_unidade = ctk.StringVar(value="%")
        ctk.CTkSegmentedButton(
            grade,
            values=["%", "kg"],
            variable=self.cf_unidade,
            height=44,
            font=self.f["normal"],
            fg_color=INPUT,
            selected_color=VERMELHO,
            selected_hover_color=VERMELHO_HOVER,
            unselected_color=INPUT,
            unselected_hover_color=NAV_HOVER,
            text_color=TEXTO,
        ).grid(row=5, column=1, sticky="ew", padx=6)

        botoes = ctk.CTkFrame(cartao, fg_color="transparent")
        botoes.pack(pady=(18, 10))
        self._botao(
            botoes, "🧮   CALCULAR", self.calcular_frete, largura=240, registrar=False
        ).pack(side="left", padx=6)
        self._botao_secundario(botoes, "Limpar", self.limpar_frete, largura=110).pack(
            side="left", padx=6
        )
        self._botao_secundario(
            botoes, "← Calculadora", lambda: self._mostrar("calculadora"), largura=160
        ).pack(side="left", padx=6)

        resultado = ctk.CTkFrame(
            cartao, fg_color=INPUT, border_width=1, border_color=BORDA, corner_radius=10
        )
        resultado.pack(fill="x", padx=24, pady=(8, 22))
        resultado.grid_columnconfigure(1, weight=1)

        self.res_cf = {}
        itens = [
            ("quebra", "Quebra (líquido − nota)"),
            ("isencao", "Isenção de quebra"),
            ("exc", "Quebra acima da isenção"),
            ("valor_quebra", "Valor da quebra (desconto)"),
            ("saldo", "Valor saldo (total − adiantamentos)"),
        ]
        for linha, (chave, rotulo) in enumerate(itens):
            pad = (12 if linha == 0 else 3, 3)
            ctk.CTkLabel(
                resultado, text=rotulo, font=self.f["normal"],
                text_color=TEXTO_SEC, anchor="w",
            ).grid(row=linha, column=0, sticky="w", padx=18, pady=pad)
            valor = ctk.CTkLabel(
                resultado, text="—", font=self.f["negrito"],
                text_color=TEXTO, anchor="e",
            )
            valor.grid(row=linha, column=1, sticky="e", padx=18, pady=pad)
            self.res_cf[chave] = valor

        ctk.CTkFrame(resultado, height=1, fg_color=BORDA).grid(
            row=5, column=0, columnspan=2, sticky="ew", padx=18, pady=8
        )
        ctk.CTkLabel(
            resultado, text="VALOR A RECEBER", font=self.f["negrito"], text_color=TEXTO_SEC
        ).grid(row=6, column=0, sticky="w", padx=18)
        self.res_cf["receber"] = ctk.CTkLabel(
            resultado, text="R$ 0,00", font=self.f["resultado"],
            text_color=TEXTO, anchor="e",
        )
        self.res_cf["receber"].grid(row=6, column=1, sticky="e", padx=18)

        self.res_cf["aviso"] = ctk.CTkLabel(
            resultado, text="", font=self.f["pequena"], text_color=AMARELO,
            anchor="w", justify="left", wraplength=600,
        )
        self.res_cf["aviso"].grid(
            row=7, column=0, columnspan=2, sticky="w", padx=18, pady=(4, 14)
        )

    def _montar_pagina_sobre(self):
        _, cartao = self._nova_pagina("sobre", "ℹ", "Sobre o projeto")

        topo = ctk.CTkFrame(cartao, fg_color="transparent")
        topo.pack(fill="x", padx=24)

        ctk.CTkLabel(
            topo, text=NOME_PROJETO, font=self.f["cabecalho"],
            text_color=TEXTO, anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            topo, text=f"Versão {VERSAO}", font=self.f["normal"],
            text_color=TEXTO_SEC, anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            topo, text=f"Criado por: {AUTOR}", font=self.f["negrito"],
            text_color=TEXTO, anchor="w",
        ).pack(anchor="w", pady=(10, 0))

        ctk.CTkLabel(
            cartao, text="Comentários do projeto", font=self.f["negrito"],
            text_color=TEXTO_SEC,
        ).pack(anchor="w", padx=24, pady=(16, 6))

        caixa = self._caixa_texto(cartao, altura=200)
        caixa.pack(fill="both", expand=True, padx=24, pady=(0, 22))
        caixa.insert("1.0", COMENTARIOS.strip())
        caixa.configure(state="disabled")      # somente leitura

    def _montar_pagina_wifi(self):
        _, cartao = self._nova_pagina("wifi", "📶", "QR Code do WiFi", expandir=False)
        self._dica(cartao, "Imprime o nome da rede, a senha e o QR Code de conexão.")

        grade = ctk.CTkFrame(cartao, fg_color="transparent")
        grade.pack(fill="x", padx=24)
        grade.grid_columnconfigure((0, 1), weight=1, uniform="campo")

        self._rotulo(grade, "Rede").grid(row=0, column=0, sticky="w")
        self.entrada_wifi_rede = self._entrada(grade)
        self.entrada_wifi_rede.insert(0, self.config["wifi_rede"])
        self.entrada_wifi_rede.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        self._rotulo(grade, "Senha").grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.entrada_wifi_senha = self._entrada(grade)
        self.entrada_wifi_senha.insert(0, self.config["wifi_senha"])
        self.entrada_wifi_senha.grid(row=1, column=1, sticky="ew", padx=(8, 0))

        self._rotulo(grade, "Título").grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.entrada_wifi_titulo = self._entrada(grade)
        self.entrada_wifi_titulo.insert(0, self.config["wifi_titulo"])
        self.entrada_wifi_titulo.grid(row=3, column=0, sticky="ew", padx=(0, 8))

        self._rotulo(grade, "Cópias").grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(12, 0))
        self.entrada_wifi_copias = self._entrada(grade)
        self.entrada_wifi_copias.insert(0, "1")
        self.entrada_wifi_copias.grid(row=3, column=1, sticky="ew", padx=(8, 0))

        self._botao(cartao, "🖨   IMPRIMIR", self.imprimir_wifi).pack(pady=(26, 26))

    def _montar_pagina_config(self):
        _, cartao = self._nova_pagina("config", "⚙", "Configurações", expandir=False)

        corpo = ctk.CTkFrame(cartao, fg_color="transparent")
        corpo.pack(fill="x", padx=24)
        corpo.grid_columnconfigure(0, weight=1)

        self._rotulo(corpo, "Impressora (nome no Windows)").grid(row=0, column=0, sticky="w")
        self.combo_impressora = self._combo(
            corpo, None, listar_impressoras(), readonly=False
        )
        self.combo_impressora.set(self.imp.nome)
        self.combo_impressora.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        self._botao_secundario(
            corpo, "Atualizar lista", self._atualizar_lista_impressoras
        ).grid(row=1, column=1)

        self._rotulo(corpo, "Fonte padrão — Texto").grid(
            row=2, column=0, sticky="w", pady=(16, 0)
        )
        self._combo(corpo, self.fonte_texto_var, list(self.imp.tamanhos)).grid(
            row=3, column=0, sticky="ew", padx=(0, 8)
        )

        self._rotulo(corpo, "Fonte padrão — Senha de banho").grid(
            row=4, column=0, sticky="w", pady=(16, 0)
        )
        self._combo(corpo, self.fonte_banho_var, list(self.imp.tamanhos)).grid(
            row=5, column=0, sticky="ew", padx=(0, 8)
        )

        ctk.CTkButton(
            cartao,
            text="💾   SALVAR E TESTAR",
            command=self.salvar_configuracoes,
            width=340,
            height=56,
            corner_radius=8,
            fg_color=VERMELHO,
            hover_color=VERMELHO_HOVER,
            font=self.f["botao"],
        ).pack(pady=(26, 26))

    def _atualizar_lista_impressoras(self):
        self.combo_impressora.configure(values=listar_impressoras())
        self._status("Lista de impressoras atualizada.")

    def salvar_configuracoes(self):
        nome = self.combo_impressora.get().strip()
        if not nome:
            messagebox.showwarning("Aviso", "Escolha ou digite o nome da impressora.")
            return

        self.imp.nome = nome
        self._salvar_config()
        self._verificar_impressora(avisar=True)

    def _salvar_config(self):
        self.config.update({
            "impressora": self.imp.nome,
            "cabecalho": self.cabecalho_var.get(),
            "fonte_texto": self.fonte_texto_var.get(),
            "fonte_banho": self.fonte_banho_var.get(),
            "wifi_titulo": self.entrada_wifi_titulo.get().strip() or "WIFI",
            "wifi_rede": self.entrada_wifi_rede.get().strip(),
            "wifi_senha": self.entrada_wifi_senha.get().strip(),
        })
        salvar_config(self.config)

    # --------------------------------------------------------
    # Execução em segundo plano (a janela não congela)
    # --------------------------------------------------------

    def _executar(self, tarefa, ao_concluir):
        """Roda `tarefa(progresso)` numa thread e chama `ao_concluir(resultado)`."""
        self._avanco = None
        self._detalhe_falha = ""
        self._definir_ocupado(True)
        self._status("Imprimindo...", "ocupado")

        def trabalho():
            try:
                resultado = tarefa(self._registrar_avanco)
            except Exception as erro:
                self.fila.put(lambda e=erro: self._falha(e))
            else:
                self.fila.put(lambda r=resultado: self._sucesso(ao_concluir, r))

        threading.Thread(target=trabalho, daemon=True).start()

    def _registrar_avanco(self, atual, total):
        # chamado pela thread de trabalho
        self._avanco = (atual, total)
        self.fila.put(
            lambda: self._status(f"Imprimindo {atual} de {total}...", "ocupado")
        )

    def _processar_fila(self):
        try:
            while True:
                self.fila.get_nowait()()
        except queue.Empty:
            pass
        finally:
            self.janela.after(100, self._processar_fila)

    def _definir_ocupado(self, ocupado):
        self.ocupado = ocupado
        for botao in self.botoes:
            botao.configure(
                state="disabled" if ocupado else "normal",
                fg_color=CINZA_OCUPADO if ocupado else VERMELHO,
            )

    def _sucesso(self, ao_concluir, resultado):
        self._definir_ocupado(False)
        self.ponto_impressora.configure(text_color=VERDE)
        ao_concluir(resultado)

    def _falha(self, erro):
        self._definir_ocupado(False)
        logging.error("Falha na impressão", exc_info=erro)

        if isinstance(erro, ErroImpressora):
            self.ponto_impressora.configure(text_color=ERRO)

        mensagem = str(erro)
        detalhe = self._detalhe_falha
        if not detalhe and self._avanco:
            atual, total = self._avanco
            detalhe = f"Impressos antes do erro: {atual} de {total}."
        if detalhe:
            mensagem += f"\n\n{detalhe}"

        self._status("Falha na impressão.", "erro")
        messagebox.showerror("Erro", mensagem)

    def _erro_inesperado(self, tipo, valor, tb):
        logging.error("Erro inesperado na interface", exc_info=(tipo, valor, tb))
        self._definir_ocupado(False)
        self._status("Erro inesperado (veja erros.log).", "erro")
        messagebox.showerror("Erro", f"Erro inesperado:\n{valor}")

    def _fechar(self):
        if self.ocupado and not messagebox.askyesno(
            "Impressão em andamento",
            "Ainda há uma impressão em andamento.\nFechar mesmo assim?",
        ):
            return
        self.janela.destroy()

    # --------------------------------------------------------
    # Leitura de campos
    # --------------------------------------------------------

    def _ler_valores(self, caixa):
        """Um valor positivo por linha. Valida tudo ANTES de imprimir."""
        valores = []

        for numero, linha in enumerate(self._ler_caixa(caixa).splitlines(), start=1):
            linha = linha.strip()
            if not linha:
                continue

            try:
                valor = parse_valor(linha)
            except ValueError:
                raise ValueError(f"Linha {numero}: valor inválido ({linha}).") from None

            if valor <= 0:
                raise ValueError(f"Linha {numero}: o valor deve ser maior que zero.")

            valores.append(valor)

        if not valores:
            raise ValueError("Digite pelo menos um valor.")

        return valores

    def _ler_operacoes(self, caixa):
        """Lê as linhas da calculadora (operador opcional + valor)."""
        operacoes = []

        for numero, linha in enumerate(self._ler_caixa(caixa).splitlines(), start=1):
            linha = linha.strip()
            if not linha:
                continue

            try:
                operacoes.append(parse_operacao(linha))
            except ValueError as erro:
                raise ValueError(f"Linha {numero}: {erro}.") from None

        if not operacoes:
            raise ValueError("Digite pelo menos um valor.")

        return operacoes

    # --------------------------------------------------------
    # Ações
    # --------------------------------------------------------

    def imprimir_texto(self):
        if self.ocupado:
            return

        linhas = [x.strip() for x in self._ler_caixa(self.caixa_texto).splitlines()]
        linhas = [x for x in linhas if x]

        if not linhas:
            messagebox.showwarning("Aviso", "Digite pelo menos um texto.")
            return

        tamanho = self.imp.tamanhos[self.fonte_texto_var.get()]
        cabecalho = self.cabecalho_var.get().strip()
        textos = [f"{cabecalho}\n{x}" if cabecalho else x for x in linhas]

        def tarefa(progresso):
            for i, texto in enumerate(textos, start=1):
                self.imp.imprimir(texto, tamanho)
                progresso(i, len(textos))
                time.sleep(PAUSA_ENTRE_IMPRESSOES)
            return len(textos)

        def concluido(qtd):
            self._limpar_caixa(self.caixa_texto)
            self._salvar_config()
            self._status(f"{qtd} impressão(ões) concluída(s).")

        self._executar(tarefa, concluido)

    def imprimir_banho(self):
        if self.ocupado:
            return

        texto = self.entrada_banho.get().strip()

        if not texto.isdecimal() or not 1 <= int(texto) <= MAX_SENHAS:
            messagebox.showwarning(
                "Aviso", f"Digite uma quantidade de 1 a {MAX_SENHAS}."
            )
            return

        quantidade = int(texto)

        if not messagebox.askyesno(
            "Confirmar", f"Gerar e imprimir {quantidade} senha(s)?"
        ):
            return

        senhas = gerar_senhas(quantidade)
        tamanho = self.imp.tamanhos[self.fonte_banho_var.get()]

        def tarefa(progresso):
            impressas = []
            arquivo = None
            try:
                for senha in senhas:
                    self.imp.imprimir_banho(senha, tamanho)
                    impressas.append(senha)
                    progresso(len(impressas), len(senhas))
                    time.sleep(PAUSA_ENTRE_IMPRESSOES)
            finally:
                # o CSV guarda exatamente o que saiu impresso, mesmo se falhar no meio
                if impressas:
                    arquivo = salvar_csv_senhas(impressas)
                    self._detalhe_falha = (
                        f"{len(impressas)} de {len(senhas)} senhas foram impressas "
                        f"e salvas em:\n{arquivo}"
                    )
            return arquivo, len(impressas)

        def concluido(resultado):
            arquivo, qtd = resultado
            self.entrada_banho.delete(0, "end")
            self._salvar_config()
            self._status(f"{qtd} senhas criadas e impressas.")
            messagebox.showinfo(
                "Concluído", f"{qtd} senhas geradas.\nArquivo: {arquivo}"
            )

        self._executar(tarefa, concluido)

    def imprimir_vale(self):
        if self.ocupado:
            return

        try:
            valores = self._ler_valores(self.caixa_vale)
        except ValueError as erro:
            messagebox.showwarning("Aviso", str(erro))
            return

        total = formatar_valor(sum(valores, Decimal("0")))

        if not messagebox.askyesno(
            "Confirmar",
            f"Imprimir {len(valores)} vale(s), total R$ {total}?",
        ):
            return

        def tarefa(progresso):
            for i, valor in enumerate(valores, start=1):
                self.imp.vale_imp(valor)
                registrar_vale(valor)
                progresso(i, len(valores))
                time.sleep(PAUSA_ENTRE_IMPRESSOES)
            return len(valores)

        def concluido(qtd):
            self.caixa_vale.delete("1.0", "end")
            self._status(f"{qtd} vale(s) impresso(s).")

        self._executar(tarefa, concluido)

    def calcular(self):
        if self.ocupado:
            return

        try:
            operacoes = self._ler_operacoes(self.caixa_calculadora)
            total = calcular_operacoes(operacoes)
        except ValueError as erro:
            messagebox.showwarning("Aviso", str(erro))
            return

        texto_total = formatar_valor(total)
        self.resultado_calculadora.configure(text=f"Total: {texto_total}")

        def tarefa(progresso):
            self.imp.imprimir_calculo(operacoes, total)

        def concluido(_resultado):
            self.caixa_calculadora.delete("1.0", "end")
            self._status(f"Calculadora: total = {texto_total}")

        self._executar(tarefa, concluido)

    def imprimir_wifi(self):
        if self.ocupado:
            return

        titulo = self.entrada_wifi_titulo.get().strip() or "WIFI"
        rede = self.entrada_wifi_rede.get().strip()
        senha = self.entrada_wifi_senha.get().strip()
        copias = self.entrada_wifi_copias.get().strip()

        if not rede:
            messagebox.showwarning("Aviso", "Digite o nome da rede.")
            return

        if not copias.isdecimal() or not 1 <= int(copias) <= MAX_COPIAS_WIFI:
            messagebox.showwarning(
                "Aviso", f"Digite um número de cópias de 1 a {MAX_COPIAS_WIFI}."
            )
            return

        copias = int(copias)

        def tarefa(progresso):
            for i in range(1, copias + 1):
                self.imp.imprimir_wifi(rede, senha, titulo)
                progresso(i, copias)
                time.sleep(PAUSA_ENTRE_IMPRESSOES)
            return copias

        def concluido(qtd):
            self._salvar_config()
            self._status(f"{qtd} QR Code(s) de WiFi impresso(s).")

        self._executar(tarefa, concluido)

    # --------------------------------------------------------
    # Carta frete
    # --------------------------------------------------------

    @staticmethod
    def _ler_numero(entrada, nome, casas, obrigatorio=True):
        texto = entrada.get().strip()

        if not texto:
            if obrigatorio:
                raise ValueError(f"Preencha: {nome}.")
            return Decimal("0")

        try:
            return parse_numero(texto, casas)
        except ValueError:
            raise ValueError(f"{nome}: valor inválido ({texto}).") from None

    def calcular_frete(self):
        try:
            nota = self._ler_numero(self.cf_nota, "Peso da nota", 3)
            liquido = self._ler_numero(self.cf_liquido, "Peso líquido", 3)
            total = self._ler_numero(self.cf_total, "Valor total", 2)
            ad1 = self._ler_numero(self.cf_adiant1, "Adiantamento 1", 2, obrigatorio=False)
            ad2 = self._ler_numero(self.cf_adiant2, "Adiantamento 2", 2, obrigatorio=False)
            isencao = self._ler_numero(self.cf_isencao, "Isenção de quebra", 4)
            valor_kg = self._ler_numero(
                self.cf_valor_kg, "Valor por kg da quebra", 4, obrigatorio=False
            )
            em_percentual = self.cf_unidade.get() == "%"

            r = calcular_carta_frete(
                nota, liquido, isencao, total, ad1, ad2, valor_kg, em_percentual
            )
        except ValueError as erro:
            messagebox.showwarning("Aviso", str(erro))
            return

        isencao_txt = f"{formatar_numero(r.isencao, 2)} kg"
        if em_percentual:
            isencao_txt += f"  ({f'{isencao.normalize():f}'.replace('.', ',')}%)"

        res = self.res_cf
        res["quebra"].configure(
            text=f"{formatar_numero(r.quebra, 2)} kg",
            text_color=ERRO if r.quebra < 0 else TEXTO,
        )
        res["isencao"].configure(text=isencao_txt)
        res["exc"].configure(text=f"{formatar_numero(r.quebra_excedente, 2)} kg")
        res["valor_quebra"].configure(text=f"R$ {formatar_valor(r.valor_quebra)}")
        res["saldo"].configure(text=f"R$ {formatar_valor(r.saldo)}")
        res["receber"].configure(
            text=f"R$ {formatar_valor(r.saldo_a_receber)}",
            text_color=VERDE if r.saldo_a_receber >= 0 else ERRO,
        )

        avisos = []
        if r.quebra_excedente == 0:
            avisos.append("Quebra dentro da isenção: nenhum desconto.")
        elif valor_kg == 0:
            avisos.append(
                "A quebra passou da isenção, mas o valor por kg da quebra não foi "
                "informado: nenhum desconto foi aplicado."
            )
        if r.quebra > 0 and r.quebra_excedente > 0:
            avisos.append(
                "O peso líquido ficou ACIMA do peso da nota além da isenção. Como na "
                "planilha, esse excesso também é descontado como quebra."
            )
        if r.saldo_a_receber < 0:
            avisos.append("Valor a receber negativo: adiantamentos e quebra passam do total.")
        res["aviso"].configure(text=" ".join(avisos))

        self._status(f"Carta frete: a receber R$ {formatar_valor(r.saldo_a_receber)}")

    def limpar_frete(self):
        for entrada in (
            self.cf_nota, self.cf_liquido, self.cf_total, self.cf_adiant1,
            self.cf_adiant2, self.cf_isencao, self.cf_valor_kg,
        ):
            entrada.delete(0, "end")

        for chave in ("quebra", "isencao", "exc", "valor_quebra", "saldo"):
            self.res_cf[chave].configure(text="—", text_color=TEXTO)
        self.res_cf["receber"].configure(text="R$ 0,00", text_color=TEXTO)
        self.res_cf["aviso"].configure(text="")
        self.cf_nota.focus()