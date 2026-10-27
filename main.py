"""Comunicação com a impressora Epson (ESC/POS) e utilitários de valores."""

import re
import threading
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import NamedTuple

import pywintypes
import win32print

# ============================================================
# CONFIGURAÇÃO
# ============================================================

NOME_IMPRESSORA = "EPSON TM-T20X Receipt6"
COLUNAS = 48  # colunas da fonte normal em papel de 80 mm (32 para 58 mm)

TITULO_BANHO = "Banho de 7 minutos"

# ============================================================
# COMANDOS ESC/POS
# ============================================================

INICIALIZAR = b"\x1b\x40"
CODEPAGE_850 = b"\x1b\x74\x02"      # acentos compatíveis com cp850
ALINHAR_CENTRO = b"\x1b\x61\x01"
ALINHAR_DIREITA = b"\x1b\x61\x02"
FONTE_NORMAL = b"\x1d\x21\x00"
CORTAR = b"\x1d\x56\x00"


def fonte(tamanho):
    """Comando de tamanho de fonte (nibble alto = largura, baixo = altura)."""
    return bytes([0x1D, 0x21, tamanho])


def avancar(linhas):
    return b"\n" * linhas


def _codificar(texto):
    # errors="replace": um caractere fora do cp850 vira "?" em vez de travar
    return texto.encode("cp850", errors="replace")


QR_CORRECAO_L = 0x30
QR_CORRECAO_M = 0x31
QR_CORRECAO_Q = 0x32
QR_CORRECAO_H = 0x33


def _gs_k(funcao, parametros=b""):
    """Comando GS ( k de QR Code: cn=0x31, fn=funcao, seguido dos parâmetros."""
    tamanho = len(parametros) + 2
    return bytes([0x1D, 0x28, 0x6B, tamanho % 256, tamanho // 256, 0x31, funcao]) + parametros


def comando_qrcode(conteudo, modulo=8, correcao=QR_CORRECAO_M):
    """Bytes que configuram, armazenam e imprimem um QR Code.

    modulo: tamanho de cada quadradinho (1 a 16).
    """
    if not 1 <= modulo <= 16:
        raise ValueError("modulo deve estar entre 1 e 16")

    return b"".join([
        _gs_k(0x43, bytes([modulo])),             # tamanho
        _gs_k(0x45, bytes([correcao])),           # correção de erro
        _gs_k(0x50, b"\x30" + conteudo.encode("utf-8")),  # armazena dados
        _gs_k(0x51, b"\x30"),                     # imprime
    ])


def montar_qr_wifi(rede, senha):
    """Texto do QR no padrão WIFI:T:WPA;S:rede;P:senha;;

    Caracteres especiais (\\ ; , : ") são escapados. Senha vazia = rede aberta.
    """
    def esc(texto):
        return re.sub(r'([\\;,:"])', r"\\\1", texto)

    if senha:
        return f"WIFI:T:WPA;S:{esc(rede)};P:{esc(senha)};;"
    return f"WIFI:T:nopass;S:{esc(rede)};;"


# ============================================================
# VALORES (Decimal evita erro de arredondamento do float)
# ============================================================

CENTAVOS = Decimal("0.01")
VALOR_MAXIMO = Decimal("99999999.99")


def parse_valor(texto):
    """Converte texto em Decimal com 2 casas.

    Aceita "10", "10.50", "10,50" e "1.234,56" (se houver vírgula,
    o ponto é tratado como separador de milhar).
    Levanta ValueError se o texto não for um valor válido.
    """
    limpo = texto.strip().replace("R$", "").replace(" ", "")

    if not re.fullmatch(r"-?[\d.,]+", limpo):
        raise ValueError(f"valor inválido: {texto!r}")

    if "," in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")

    try:
        valor = Decimal(limpo)
    except InvalidOperation:
        raise ValueError(f"valor inválido: {texto!r}") from None

    if abs(valor) > VALOR_MAXIMO:
        raise ValueError(f"valor muito alto: {texto!r}")

    return valor.quantize(CENTAVOS, rounding=ROUND_HALF_UP)


NUMERO_MAXIMO = Decimal("999999999999")


def parse_numero(texto, casas=2):
    """Número NÃO negativo no padrão brasileiro, arredondado em `casas`.

    Aceita 32450, 32.450 (ponto = milhar), 32450,5 e 32.450,5.
    Ponto sem vírgula só vira milhar quando tem exatamente 3 dígitos depois
    dele (32.450); 10.50 continua sendo decimal.
    Levanta ValueError se não for um número válido.
    """
    limpo = texto.strip().replace(" ", "")

    if not re.fullmatch(r"\d[\d.,]*", limpo) or limpo.count(",") > 1:
        raise ValueError(f"número inválido: {texto!r}")

    if "," in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"[1-9]\d{0,2}(\.\d{3})+", limpo):
        limpo = limpo.replace(".", "")

    try:
        valor = Decimal(limpo)
    except InvalidOperation:
        raise ValueError(f"número inválido: {texto!r}") from None

    if valor > NUMERO_MAXIMO:
        raise ValueError(f"número muito alto: {texto!r}")

    return valor.quantize(Decimal(1).scaleb(-casas), rounding=ROUND_HALF_UP)


def formatar_numero(valor, casas=2):
    """Número no padrão brasileiro, com espaço no milhar (ex.: 32 450,50)."""
    valor = Decimal(valor).quantize(Decimal(1).scaleb(-casas), rounding=ROUND_HALF_UP)
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", " ")


def formatar_valor(valor):
    """Formata para exibição.

    1       -> 1,00
    3000    -> 3 000,00
    8278.25 -> 8 278,25
    """
    return formatar_numero(valor, 2)


OPERADORES = {
    "+": "+", "-": "-",
    "*": "*", "x": "*", "X": "*", "×": "*",
    "/": "/", "÷": "/",
}
SIMBOLOS_FITA = {"+": "+", "-": "-", "*": "x", "/": "/"}


def parse_operacao(texto):
    """Lê uma linha da calculadora e devolve (operador, valor).

    '*2' -> ('*', 2.00)   '/4' -> ('/', 4.00)
    '-30' -> ('-', 30.00)  '10' -> ('+', 10.00)  (sem sinal = soma)
    """
    texto = texto.strip()
    operador = "+"

    if texto and texto[0] in OPERADORES:
        operador = OPERADORES[texto[0]]
        texto = texto[1:]
        if texto.strip().startswith("-"):
            raise ValueError("use apenas um sinal")

    valor = parse_valor(texto)

    if operador == "/" and valor == 0:
        raise ValueError("divisão por zero")

    return operador, valor


def calcular_operacoes(operacoes):
    """Aplica as operações de cima para baixo (estilo fita, sem precedência)."""
    if operacoes[0][0] in "*/":
        raise ValueError("A primeira linha precisa ser um valor, sem * ou /.")

    total = Decimal("0")
    for operador, valor in operacoes:
        if operador == "+":
            total += valor
        elif operador == "-":
            total -= valor
        elif operador == "*":
            total *= valor
        else:
            total /= valor

    if abs(total) > VALOR_MAXIMO:
        raise ValueError("O resultado é grande demais.")

    # mantém a precisão nas etapas e arredonda só no final
    return total.quantize(CENTAVOS, rounding=ROUND_HALF_UP)


class ResultadoCartaFrete(NamedTuple):
    quebra: Decimal           # peso líquido - peso da nota (negativo = faltou peso)
    isencao: Decimal          # kg de isenção de quebra
    quebra_excedente: Decimal # kg acima da isenção (0 se dentro da isenção)
    valor_quebra: Decimal     # R$ a descontar (SEM arredondar, como na planilha)
    saldo: Decimal            # R$ valor total - adiantamentos
    saldo_a_receber: Decimal  # R$ saldo - valor da quebra


def calcular_carta_frete(
    peso_nota,
    peso_liquido,
    isencao,
    valor_total,
    adiantamento_1=Decimal("0"),
    adiantamento_2=Decimal("0"),
    valor_kg_quebra=Decimal("0"),
    isencao_em_percentual=True,
):
    """Mesma conta da planilha CALCULO_FRETE_MANUAL.xlsx.

    QUEBRA            = peso líquido - peso da nota
    ISENÇÃO           = kg informados, ou peso da nota x % (isencao_em_percentual)
    VALOR QUEBRA      = SE |quebra| > isenção: (|quebra| - isenção) x valor por kg
                        senão 0   (igual à fórmula da planilha, que usa ABS)
    VALOR SALDO       = valor total - adiantamento 1 - adiantamento 2
    A RECEBER         = valor saldo - valor da quebra

    Nada é arredondado no meio da conta; arredonde só ao exibir.
    """
    if peso_nota <= 0:
        raise ValueError("O peso da nota deve ser maior que zero.")

    if isencao_em_percentual:
        if isencao > 100:
            raise ValueError("A isenção em % não pode passar de 100.")
        isencao_kg = peso_nota * isencao / 100
    else:
        isencao_kg = isencao

    quebra = peso_liquido - peso_nota

    if abs(quebra) > isencao_kg:
        excedente = abs(quebra) - isencao_kg
        valor_quebra = excedente * valor_kg_quebra
    else:
        excedente = Decimal("0")
        valor_quebra = Decimal("0")

    saldo = valor_total - adiantamento_1 - adiantamento_2

    return ResultadoCartaFrete(
        quebra, isencao_kg, excedente, valor_quebra, saldo, saldo - valor_quebra
    )


# ============================================================
# IMPRESSORA
# ============================================================

def listar_impressoras():
    """Nomes das impressoras instaladas no Windows."""
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        return sorted(p[2] for p in win32print.EnumPrinters(flags))
    except pywintypes.error:
        return []


class ErroImpressora(Exception):
    """Falha de comunicação com a impressora (mensagem pronta para o usuário)."""


class Impressora:

    def __init__(self, nome=NOME_IMPRESSORA):
        self.nome = nome
        self._trava = threading.Lock()   # um trabalho por vez no spooler

        self.tamanhos = {
            "Normal": 0x00,
            "Média (2x2)": 0x11,
            "Grande (3x3)": 0x22,
            "Muito Grande (4x4)": 0x33,
            "Gigante (8x8)": 0x77,
        }

    # --------------------------------------------------------
    # Envio (único lugar que fala com o spooler)
    # --------------------------------------------------------

    def _enviar(self, nome_doc, dados):
        with self._trava:
            try:
                h = win32print.OpenPrinter(self.nome)
                try:
                    win32print.StartDocPrinter(h, 1, (nome_doc, None, "RAW"))
                    win32print.StartPagePrinter(h)
                    win32print.WritePrinter(h, bytes(dados))
                    win32print.EndPagePrinter(h)
                    win32print.EndDocPrinter(h)
                finally:
                    win32print.ClosePrinter(h)
            except pywintypes.error as erro:
                raise ErroImpressora(
                    f"Não foi possível imprimir em '{self.nome}'.\n"
                    "Verifique se a impressora está ligada, com papel "
                    "e conectada ao computador.\n\n"
                    f"Detalhe: {erro.strerror}"
                ) from erro

    def disponivel(self):
        """True se a impressora existe no Windows e pode ser aberta."""
        try:
            h = win32print.OpenPrinter(self.nome)
        except pywintypes.error:
            return False
        win32print.ClosePrinter(h)
        return True

    @staticmethod
    def _iniciar():
        return INICIALIZAR + CODEPAGE_850

    # --------------------------------------------------------
    # Documentos
    # --------------------------------------------------------

    def imprimir(self, texto, tamanho):
        dados = b"".join([
            self._iniciar(),
            ALINHAR_CENTRO,
            fonte(tamanho),
            _codificar(texto),
            FONTE_NORMAL,
            avancar(6),
            CORTAR,
        ])
        self._enviar("Texto", dados)

    def imprimir_banho(self, senha, tamanho):
        dados = b"".join([
            self._iniciar(),
            ALINHAR_CENTRO,
            fonte(0x11),
            _codificar(TITULO_BANHO),
            avancar(2),
            fonte(tamanho),
            _codificar(senha),
            FONTE_NORMAL,
            avancar(6),
            CORTAR,
        ])
        self._enviar("Senha", dados)

    def vale_imp(self, valor):
        data = datetime.now().strftime("%d/%m/%Y")
        separador = "=" * 32

        dados = b"".join([
            self._iniciar(),
            ALINHAR_CENTRO,
            fonte(0x11),
            _codificar("POSTO 44\n"),
            FONTE_NORMAL,
            _codificar(f"{separador}\nData: {data}\n\n"),
            fonte(0x22),
            _codificar("VALE\n"),
            _codificar(f"R$ {formatar_valor(valor)}\n\n"),
            FONTE_NORMAL,
            _codificar(
                f"{separador}\n\n"
                "Autorizado por:\n\n"
                f"{'_' * 30}\n\n"
                "Valido somente no Posto 44\n"
            ),
            avancar(4),
            CORTAR,
        ])
        self._enviar("Vale Combustivel", dados)

    def imprimir_calculo(self, operacoes, total):
        """Fita de cálculo: valor + operador à direita, linha, total."""
        linhas = [
            f"{formatar_valor(valor)} {SIMBOLOS_FITA[op]}"
            for op, valor in operacoes
        ]
        texto_total = formatar_valor(total)

        # fonte 2x2 = metade das colunas
        largura = min(
            max(len(x) for x in [*linhas, texto_total]),
            COLUNAS // 2,
        )

        dados = b"".join([
            self._iniciar(),
            ALINHAR_DIREITA,
            fonte(0x11),
            _codificar("\n".join(linhas) + "\n"),
            _codificar("-" * largura + "\n"),
            _codificar(texto_total + "\n"),
            FONTE_NORMAL,
            avancar(3),
            CORTAR,
        ])
        self._enviar("Calculadora", dados)

    def imprimir_wifi(self, rede, senha, titulo="WIFI"):
        """Imprime título, nome da rede, senha e o QR Code de conexão."""
        linhas = [titulo, rede]
        if senha:
            linhas.append(f"Senha: {senha}")

        dados = b"".join([
            self._iniciar(),
            ALINHAR_CENTRO,
            fonte(0x11),
            _codificar("\n".join(linhas) + "\n\n"),
            FONTE_NORMAL,
            comando_qrcode(montar_qr_wifi(rede, senha)),
            avancar(4),
            CORTAR,
        ])
        self._enviar("QR Code WiFi", dados)