# -*- coding: utf-8 -*-
"""
Normalização de documento, data, valor e texto.
===============================================

Tudo que vem da Vórtx e da Grafeno passa por aqui antes de qualquer comparação.
Os dois lados chegam em formatos diferentes — a Grafeno manda ``R$ 1.346,63`` e
``23/03/2026`` como texto, a Vórtx manda número e data de verdade — e a chave de
conciliação só funciona se os dois virarem a mesma coisa.

Chave canônica:  ``documento|AAAA-MM-DD|0.00``
"""

import re
import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

#  config só carrega caminhos e constantes — não importa nada de core, então
#  esta dependência não fecha ciclo.
from core import config

CENTAVO = Decimal("0.01")

#  Excel conta os dias a partir de 30/12/1899 (o bug do ano 1900 já embutido).
_EPOCA_EXCEL = date(1899, 12, 30)

_SO_DIGITOS = re.compile(r"\D+")
_MOEDA = re.compile(r"[^\d,.\-]")


# ---------------------------------------------------------------------------
#  Documento
# ---------------------------------------------------------------------------
def documento(valor) -> str:
    """``349.532.217-53`` -> ``34953221753``. Devolve '' para vazio."""
    if valor is None:
        return ""
    texto_bruto = str(valor).strip()
    if not texto_bruto:
        return ""
    digitos = _SO_DIGITOS.sub("", texto_bruto)
    #  CPF exportado como número perde o zero à esquerda ("453124780").
    if 0 < len(digitos) < 11:
        digitos = digitos.zfill(11)
    elif 11 < len(digitos) < 14:
        digitos = digitos.zfill(14)
    return digitos


def documento_formatado(digitos: str) -> str:
    """Devolve o documento pontuado, como o Termo exibe."""
    d = documento(digitos)
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return d


# ---------------------------------------------------------------------------
#  Data
# ---------------------------------------------------------------------------
def data(valor):
    """Converte para ``date``. Aceita datetime, date, serial do Excel e texto."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        #  Serial do Excel. Abaixo de 1000 é quase certo que não é data.
        if valor < 1000:
            return None
        return _EPOCA_EXCEL + timedelta(days=int(valor))

    texto_bruto = str(valor).strip()
    if not texto_bruto:
        return None
    #  Descarta a parte de hora, quando vier.
    texto_bruto = texto_bruto.split(" ")[0].split("T")[0]
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(texto_bruto, formato).date()
        except ValueError:
            continue
    return None


def iso(d) -> str:
    """Data em ``AAAA-MM-DD``; '' quando não houver data."""
    d = data(d)
    return d.isoformat() if d else ""


def br(d) -> str:
    """Data em ``dd/mm/aaaa``; '' quando não houver data."""
    d = data(d)
    return d.strftime("%d/%m/%Y") if d else ""


def serial_excel(d) -> str:
    """Número de série da data no Excel — o que ``CONCAT`` de uma data produz."""
    d = data(d)
    return str((d - _EPOCA_EXCEL).days) if d else ""


# ---------------------------------------------------------------------------
#  Valor
# ---------------------------------------------------------------------------
def valor(v):
    """``R$ 1.346,63`` -> ``Decimal('1346.63')``. Devolve None para vazio."""
    if v is None or v == "":
        return None
    if isinstance(v, Decimal):
        return v.quantize(CENTAVO, rounding=ROUND_HALF_UP)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return Decimal(str(v)).quantize(CENTAVO, rounding=ROUND_HALF_UP)

    texto_bruto = _MOEDA.sub("", str(v).strip())
    if not texto_bruto:
        return None
    #  Formato brasileiro: ponto é milhar, vírgula é decimal.
    if "," in texto_bruto:
        texto_bruto = texto_bruto.replace(".", "").replace(",", ".")
    try:
        return Decimal(texto_bruto).quantize(CENTAVO, rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return None


def valor_texto(v) -> str:
    """Valor em ``0.00``, do jeito que entra na chave."""
    d = valor(v)
    return f"{d:.2f}" if d is not None else ""


# ---------------------------------------------------------------------------
#  Texto
# ---------------------------------------------------------------------------
#  Assinatura do UTF-8 lido como latin-1 / cp1252: 'Ã©', 'Ã§', 'Ãµ', 'Âº'...
_MOJIBAKE = ("Ã", "Â")
#  Caractere de substituição — aqui o original já se perdeu na exportação.
_PERDIDO = ("�", "ï¿½")


def texto(v) -> str:
    """Limpa espaços e conserta a acentuação quando ainda dá para consertar."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    return corrigir_acentuacao(s)


def corrigir_acentuacao(s: str) -> str:
    """
    Desfaz o UTF-8 lido como latin-1 (``SAÃšDE`` -> ``SAÚDE``).

    Não há o que fazer quando o export já trocou o caractere por U+FFFD:
    nesse caso a informação se perdeu na origem. :func:`tem_perda_acentuacao`
    detecta a situação para a bateria de validações avisar.
    """
    if not any(marca in s for marca in _MOJIBAKE):
        return s
    for codificacao in ("latin-1", "cp1252"):
        try:
            candidato = s.encode(codificacao).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        #  Só aceita se realmente melhorou — não inventa acento onde não havia.
        if not any(marca in candidato for marca in _MOJIBAKE):
            return candidato
    return s


def tem_perda_acentuacao(s: str) -> bool:
    """True quando o texto traz caractere de substituição irrecuperável."""
    return any(marca in s for marca in _PERDIDO)


def sem_acento(s: str) -> str:
    """Versão sem acento e em maiúsculas, para comparar nome de pessoa."""
    s = unicodedata.normalize("NFKD", texto(s))
    return "".join(c for c in s if not unicodedata.combining(c)).upper()


# ---------------------------------------------------------------------------
#  Chave de conciliação
# ---------------------------------------------------------------------------
def chave(doc, vencimento, val) -> str:
    """``documento|AAAA-MM-DD|0.00`` — a chave exata da passada 1."""
    return f"{documento(doc)}|{iso(vencimento)}|{valor_texto(val)}"


def chave_parcial(doc, vencimento) -> str:
    """``documento|AAAA-MM-DD`` — a chave da passada 2, sem o valor."""
    return f"{documento(doc)}|{iso(vencimento)}"


def so_digitos(v) -> str:
    """Só os dígitos, para comparar identificador com identificador."""
    return re.sub(r"\D", "", str(v or ""))


def chave_boleto(numero_titulo) -> str:
    """
    O ``Nosso_Número`` que a Grafeno dá ao boleto de um título da Vórtx.

    É o ``NumeroTitulo`` com o prefixo de convênio: ``382477119`` vira
    ``403382477119``. Serve **só para desempatar** entre boletos que a chave
    composta já casou — usada sozinha ela pareia gente diferente (ver o
    comentário de ``config.PREFIXO_NOSSO_NUMERO``).

    Devolve ``''`` quando o número não é numérico, e aí não há desempate.
    """
    digitos = so_digitos(numero_titulo)
    return f"{config.PREFIXO_NOSSO_NUMERO}{digitos}" if digitos else ""
