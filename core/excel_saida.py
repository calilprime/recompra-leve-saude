# -*- coding: utf-8 -*-
"""
Escrita da planilha da rodada.
==============================

**A estrutura do arquivo é congelada** (seção 5.2). Ele alimenta os BIs da casa
e é usado em outras análises: não se cria, exclui, move, renomeia ou oculta
coluna; não se cria, exclui ou renomeia aba; não se mexe em cabeçalho,
formatação ou formato de número.

Por isso a gravação é cirúrgica, direto no XML de dentro do ``.xlsx`` — o mesmo
padrão do ``fnet_app.py``. Abrir um arquivo de 50 MB com o openpyxl em modo de
escrita reconstruiria a planilha inteira e derrubaria em silêncio o que a
biblioteca não conhece. Aqui o cabeçalho, os estilos, as demais abas e tudo o
que vem antes e depois do ``<sheetData>`` são preservados byte a byte.

O que muda no arquivo, e só isso:

* abas VORTX e GRAFENO — as linhas de dado, substituídas pela extração da rodada;
* ``VORTX!AR`` — ``PROCX`` com intervalo fixo vira ``CONT.SES`` de coluna
  inteira, testando os status na ordem de prioridade (ajuste 1);
* ``TERMO!B5:H5`` — sai a condição ``AQ=1``, que descartava título com mais de
  um boleto, e os intervalos passam a coluna inteira (ajuste 2);
* ``INFORMAÇÕES!D8:D11`` — os parâmetros da rodada.

Juros (``D13``) e multa (``D14``) não são tocados: são a fonte, não o destino.
"""

import re
import shutil
import zipfile
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from core import config
from core import normalizacao as nz

EPOCA_EXCEL = date(1899, 12, 30)


# ---------------------------------------------------------------------------
#  Fórmulas corrigidas (guardadas em inglês e com vírgula, como o XML exige)
# ---------------------------------------------------------------------------
def _formula_status(linha):
    """
    ``VORTX!AR`` — ajuste 1. Duas correções em uma fórmula só.

    **Coluna inteira.** O ``PROCX`` original lia até a linha 124.996 e a aba
    GRAFENO tem 141.977: tudo além do limite devolvia ``N/A`` e o título saía do
    Termo sem aviso. ``CONT.SES`` de coluna inteira não tem limite e o Excel o
    otimiza para o intervalo usado.

    **Qual boleto.** Testa primeiro a ligação pelo ``Nosso_Número``
    (``"403"&NumeroTitulo``), restrita aos boletos que a chave da coluna AP já
    casou; só se ela não resolver é que cai na prioridade de status. É a mesma
    ordem de ``conciliacao.escolher_boleto`` — os dois caminhos de cálculo
    precisam dar o mesmo Termo, e a validação de convergência confere isso.

    A restrição à coluna AP é o que torna a ligação segura: sozinha ela pareia
    títulos com boletos de outras pessoas (ver ``config.PREFIXO_NOSSO_NUMERO``).
    """
    numero = f"{config.VORTX_COL_NUMERO_TITULO}{linha}"
    nosso = f"GRAFENO!${config.GRAFENO_COL_NOSSO}:${config.GRAFENO_COL_NOSSO}"
    status_col = f"GRAFENO!${config.GRAFENO_COL_STATUS}:${config.GRAFENO_COL_STATUS}"
    chave = f"GRAFENO!$AE:$AE"

    testes = ""
    fechamento = ""
    #  1) o boleto identificado pelo Nosso_Número, dentro dos que a chave casou
    for status in config.PRIORIDADE_STATUS:
        testes += (f'IF(COUNTIFS({chave},AP{linha},{nosso},'
                   f'"{config.PREFIXO_NOSSO_NUMERO}"&{numero},'
                   f'{status_col},"{status}")>0,"{status}",')
        fechamento += ")"
    #  2) sem ligação, a prioridade de status decide
    for status in config.PRIORIDADE_STATUS:
        testes += (f'IF(COUNTIFS({chave},AP{linha},{status_col},'
                   f'"{status}")>0,"{status}",')
        fechamento += ")"
    return f'IF(AP{linha}="","",{testes}"N/A"{fechamento})'


FORMULAS_VORTX = {
    "AP": lambda n: f"_xlfn.CONCAT(J{n},P{n},U{n})",
    "AQ": lambda n: f'IF(AP{n}="","",COUNTIF(GRAFENO!AE:AE,VORTX!AP{n}))',
    "AR": _formula_status,
    "AS": lambda n: (f'IF(AR{n}="","",IF(AND(AR{n}="{config.STATUS_RECOMPRAVEL}",'
                     f"AT{n}<=INFORMAÇÕES!$D$11,AT{n}>=INFORMAÇÕES!$D$10),1,0))"),
    "AT": lambda n: f'IF(AP{n}="","",P{n})',
    "AU": lambda n: f'IF(AR{n}="","",INFORMAÇÕES!$D$8)',
    "AV": lambda n: f'IF(AP{n}="","",U{n})',
    "AW": lambda n: f'IFERROR(IF(AR{n}="","",DATEDIF(AT{n},AU{n},"D")*AS{n}),"-")',
    "AX": lambda n: f'IFERROR(AV{n}*INFORMAÇÕES!$D$14*AS{n},"")',
    "AY": lambda n: (f'IF(AP{n}="","",IFERROR(((1+INFORMAÇÕES!$D$13)^(AW{n})-1)'
                     f"*AV{n},0))"),
    "AZ": lambda n: f'IFERROR(SUM(AV{n},AX{n},AY{n})*AS{n},"")',
}

FORMULAS_GRAFENO = {
    "AE": lambda n: f'IF(A{n}=0,"",_xlfn.CONCAT(D{n},DATEVALUE(H{n}),VALUE(I{n})))',
}

#  ``TERMO!B5:H5`` — ajuste 2. Sem a condição AQ=1 e em coluna inteira.
FORMULAS_TERMO = {
    "B": "UPPER(_xlfn._xlws.FILTER(VORTX!$K:$K,VORTX!$AS:$AS=1))",
    "C": "UPPER(_xlfn._xlws.FILTER(VORTX!$J:$J,VORTX!$AS:$AS=1))",
    "D": "UPPER(_xlfn._xlws.FILTER(VORTX!$L:$L,VORTX!$AS:$AS=1))",
    "E": "(_xlfn._xlws.FILTER(VORTX!$P:$P,VORTX!$AS:$AS=1))",
    "F": "(_xlfn._xlws.FILTER(VORTX!$AU:$AU,VORTX!$AS:$AS=1))",
    "G": "(_xlfn._xlws.FILTER(VORTX!$AV:$AV,VORTX!$AS:$AS=1))",
    "H": "(_xlfn._xlws.FILTER(VORTX!$AZ:$AZ,VORTX!$AS:$AS=1))",
}


# ---------------------------------------------------------------------------
#  Utilidades de coluna e de XML
# ---------------------------------------------------------------------------
def indice_coluna(letras):
    n = 0
    for ch in letras.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def letra_coluna(indice):
    letras = ""
    while indice > 0:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def _escapar(texto):
    return (str(texto).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def _numero(v):
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, Decimal):
        v = float(v)
    if isinstance(v, int):
        return str(v)
    f = float(v)
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    return repr(f)


def _serial(d):
    if isinstance(d, datetime):
        base = d.date()
        fracao = (d.hour * 3600 + d.minute * 60 + d.second) / 86400
    else:
        base, fracao = d, 0
    dias = (base - EPOCA_EXCEL).days
    return str(dias) if not fracao else repr(dias + fracao)


def _celula(referencia, valor, estilo=None, formula=None, array_ref=None):
    """Monta uma tag ``<c>``. Texto vai como inlineStr, sem mexer em sharedStrings."""
    s = f' s="{estilo}"' if estilo is not None else ""
    if formula is not None:
        tag_f = (f'<f t="array" ref="{array_ref}" ca="1">{_escapar(formula)}</f>'
                 if array_ref else f"<f>{_escapar(formula)}</f>")
        cache = ""
        if valor is not None:
            if isinstance(valor, str):
                cache = f'<v>{_escapar(valor)}</v>'
                return f'<c r="{referencia}"{s} t="str">{tag_f}{cache}</c>'
            cache = f"<v>{_conteudo_numerico(valor)}</v>"
        return f'<c r="{referencia}"{s}>{tag_f}{cache}</c>'

    if valor is None or valor == "":
        return f'<c r="{referencia}"{s}/>'
    if isinstance(valor, bool):
        return f'<c r="{referencia}"{s} t="b"><v>{1 if valor else 0}</v></c>'
    if isinstance(valor, (datetime, date)):
        return f'<c r="{referencia}"{s}><v>{_serial(valor)}</v></c>'
    if isinstance(valor, (int, float, Decimal)):
        return f'<c r="{referencia}"{s}><v>{_numero(valor)}</v></c>'
    return (f'<c r="{referencia}"{s} t="inlineStr"><is>'
            f'<t xml:space="preserve">{_escapar(valor)}</t></is></c>')


def _conteudo_numerico(valor):
    if isinstance(valor, (datetime, date)):
        return _serial(valor)
    return _numero(valor)


# ---------------------------------------------------------------------------
#  Leitura do esqueleto de uma aba
# ---------------------------------------------------------------------------
class Aba:
    """O XML de uma aba, partido em antes / sheetData / depois."""

    def __init__(self, caminho_interno, xml):
        self.caminho_interno = caminho_interno
        self.xml = xml
        vazia = re.search(r"<sheetData\s*/>", xml)
        if vazia:
            self.antes = xml[:vazia.start()]
            self.corpo = ""
            self.depois = xml[vazia.end():]
        else:
            inicio = re.search(r"<sheetData[^>]*>", xml)
            fim = xml.index("</sheetData>")
            self.antes = xml[:inicio.end()]
            self.corpo = xml[inicio.end():fim]
            self.depois = xml[fim + len("</sheetData>"):]
        self.linhas = self._separar_linhas()

    def _separar_linhas(self):
        linhas = {}
        for achado in re.finditer(
                r'<row(?P<attrs>[^>]*\br="(?P<n>\d+)"[^>]*)(?:/>|>(?P<body>.*?)</row>)',
                self.corpo, flags=re.S):
            linhas[int(achado.group("n"))] = (achado.group("attrs"),
                                              achado.group("body") or "")
        return linhas

    def estilos_da_linha(self, numero):
        """
        Índice de estilo por coluna, colhido de uma linha modelo.

        Devolve ``None`` quando a linha não existe — senão o ``or`` de quem
        chama nunca cai no plano B e a linha gerada sai sem formato de número,
        o que transforma data em número de série na cara do usuário.
        """
        if numero not in self.linhas:
            return None
        attrs, corpo = self.linhas[numero]
        estilos = {}
        for celula in re.finditer(r'<c\b[^>]*r="([A-Z]+)\d+"[^>]*?(?:/>|>)', corpo):
            tag = celula.group(0)
            estilo = re.search(r'\bs="(\d+)"', tag)
            estilos[celula.group(1)] = estilo.group(1) if estilo else None
        estilo_linha = re.search(r'\bs="(\d+)"', attrs)
        estilos["__linha__"] = estilo_linha.group(1) if estilo_linha else None
        estilos["__attrs__"] = attrs
        return estilos

    def xml_com(self, corpo, dimensao=None):
        antes = self.antes
        if dimensao:
            antes = re.sub(r'<dimension ref="[^"]*"/>',
                           f'<dimension ref="{dimensao}"/>', antes, count=1)
        return antes + corpo + "</sheetData>" + self.depois


class Pacote:
    """O ``.xlsx`` aberto como zip, com as partes na memória."""

    def __init__(self, caminho):
        self.caminho = Path(caminho)
        with zipfile.ZipFile(self.caminho) as z:
            self.partes = {n: z.read(n) for n in z.namelist()}
        self._mapa_abas = self._ler_mapa_abas()

    def _ler_mapa_abas(self):
        wb = self.partes["xl/workbook.xml"].decode("utf-8")
        rels = self.partes["xl/_rels/workbook.xml.rels"].decode("utf-8")
        alvo = {}
        for rel in re.findall(r"<Relationship\b[^>]*?>", rels):
            ident = re.search(r'Id="([^"]+)"', rel)
            caminho = re.search(r'Target="([^"]+)"', rel)
            if ident and caminho:
                alvo[ident.group(1)] = caminho.group(1)
        mapa = {}
        for sh in re.findall(r"<sheet\b[^>]*?>", wb):
            nome = re.search(r'name="([^"]+)"', sh)
            rid = re.search(r'r:id="([^"]+)"', sh)
            if nome and rid and rid.group(1) in alvo:
                destino = alvo[rid.group(1)].lstrip("/")
                if not destino.startswith("xl/"):
                    destino = "xl/" + destino
                mapa[_desescapar(nome.group(1))] = destino
        return mapa

    def nome_real_aba(self, nome):
        if nome in self._mapa_abas:
            return nome
        alvo = nz.sem_acento(nome)
        for existente in self._mapa_abas:
            if nz.sem_acento(existente) == alvo:
                return existente
        raise KeyError(f"Aba {nome!r} não existe. Abas: "
                       f"{', '.join(self._mapa_abas)}")

    def aba(self, nome):
        nome = self.nome_real_aba(nome)
        caminho = self._mapa_abas[nome]
        return Aba(caminho, self.partes[caminho].decode("utf-8"))

    def gravar_aba(self, aba, xml):
        self.partes[aba.caminho_interno] = xml.encode("utf-8")

    def forcar_recalculo(self):
        """Manda o Excel recalcular tudo ao abrir e descarta a cadeia antiga."""
        wb = self.partes["xl/workbook.xml"].decode("utf-8")
        if "<calcPr" in wb:
            if "fullCalcOnLoad" not in wb:
                wb = re.sub(r"<calcPr\b([^>]*?)/>",
                            r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
        else:
            wb = re.sub(r"(</sheets>)", r'\1<calcPr calcId="0" fullCalcOnLoad="1"/>',
                        wb, count=1)
        self.partes["xl/workbook.xml"] = wb.encode("utf-8")

        #  A calcChain descreve a ordem de cálculo das fórmulas antigas.
        #  Reescrevemos as linhas: ela fica inconsistente e tem de sair.
        self.partes.pop("xl/calcChain.xml", None)
        tipos = self.partes.get("[Content_Types].xml", b"").decode("utf-8")
        tipos = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", tipos)
        self.partes["[Content_Types].xml"] = tipos.encode("utf-8")
        rels = self.partes["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rels = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "", rels)
        self.partes["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    def salvar(self, destino):
        destino = Path(destino)
        temporario = destino.with_suffix(destino.suffix + ".tmp")
        with zipfile.ZipFile(temporario, "w", zipfile.ZIP_DEFLATED,
                             compresslevel=6) as z:
            for nome, dados in self.partes.items():
                z.writestr(nome, dados)
        temporario.replace(destino)
        return destino


def _desescapar(s):
    return (s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
            .replace("&quot;", '"').replace("&apos;", "'"))


# ---------------------------------------------------------------------------
#  Escrita das abas de dado
# ---------------------------------------------------------------------------
def _escrever_aba_dados(pacote, nome_aba, registros, colunas_formula, log=None):
    """
    Substitui as linhas de dado de uma aba, mantendo cabeçalho e estilos.

    ``registros`` é uma lista de tuplas brutas, na ordem original das colunas
    do export. As colunas auxiliares recebem fórmula, uma por linha.
    """
    log = log or (lambda *a, **k: None)
    aba = pacote.aba(nome_aba)
    modelo = aba.estilos_da_linha(2) or aba.estilos_da_linha(1) or {}
    attrs_modelo = modelo.get("__attrs__", "")
    tem_spans = "spans=" in attrs_modelo

    #  Largura fixa para todas as linhas: o ``spans`` precisa ser o mesmo, e a
    #  dimensão da aba também.
    ultima_coluna = max(
        [indice_coluna(letra) for letra in list(modelo) + list(colunas_formula)
         if not letra.startswith("__")]
        + [max((len(bruto) for bruto in registros), default=0)]
    )

    partes = []
    cabecalho = aba.linhas.get(1)
    if cabecalho:
        partes.append(f"<row{cabecalho[0]}>{cabecalho[1]}</row>")

    for deslocamento, bruto in enumerate(registros):
        numero = deslocamento + 2
        celulas = []
        for posicao, valor in enumerate(bruto, start=1):
            if valor is None or valor == "":
                continue
            letra = letra_coluna(posicao)
            celulas.append(_celula(f"{letra}{numero}", valor,
                                   estilo=modelo.get(letra)))
        for letra, construir in colunas_formula.items():
            celulas.append(_celula(f"{letra}{numero}", None,
                                   estilo=modelo.get(letra),
                                   formula=construir(numero)))
        atributos = f' r="{numero}"'
        if tem_spans:
            atributos += f' spans="1:{ultima_coluna}"'
        if modelo.get("__linha__"):
            atributos += f' s="{modelo["__linha__"]}" customFormat="1"'
        partes.append(f"<row{atributos}>{''.join(celulas)}</row>")

    ultima_linha = len(registros) + 1
    dimensao = f"A1:{letra_coluna(ultima_coluna)}{max(ultima_linha, 1)}"
    pacote.gravar_aba(aba, aba.xml_com("".join(partes), dimensao))
    log(f"  Aba {nome_aba}: {len(registros)} linhas gravadas "
        f"(A1 até {letra_coluna(ultima_coluna)}{ultima_linha}).")
    return ultima_linha


# ---------------------------------------------------------------------------
#  Aba TERMO DE RECOMPRA
# ---------------------------------------------------------------------------
def _escrever_termo(pacote, linhas_termo, totais=None, log=None):
    """
    Escreve o Termo calculado pelo motor e reancora as fórmulas de matriz.

    Os valores gravados são os do motor; as fórmulas ``FILTER`` ficam na linha
    âncora, como conferência independente (seção 5.3). A validação de
    convergência já garantiu que os dois caminhos dão o mesmo resultado — se não
    dessem, a emissão nem teria chegado aqui.

    **A linha âncora é escrita sempre**, inclusive quando o Termo sai vazio.
    Isso é o defeito que a simulação de 14/08 revelou: o laço percorria
    ``linhas_termo`` e a fórmula ia só na primeira iteração, então com zero
    títulos nada era escrito — ``B5:H5`` saía em branco, as ``FILTER`` do
    template desapareciam do arquivo entregue, e a validação "Convergência
    Python ↔ Excel" comparava 0 com 0 e dizia ``OK``. Não convergia: não havia
    o que convergir. Com a âncora sempre presente, o arquivo vazio ainda leva a
    fórmula dentro, e quem abrir vê o Excel recalcular zero — que é uma
    informação, ao contrário de uma célula em branco.

    Sobre gravar valor **e** fórmula na mesma faixa: não há conflito. É assim
    que o próprio Excel guarda um intervalo derramado — a célula mestra leva
    ``<f t="array" ref="B5:H2330">`` e as demais levam só o valor em cache. O
    valor em cache não é decoração: é o que os BIs da casa e o
    ``retroativo.ler_termo_emitido`` leem, porque leem com ``data_only=True`` e
    fórmula sem cache chega como ``None``. Tirar os valores para "preservar a
    fórmula" cegaria os dois.
    """
    log = log or (lambda *a, **k: None)
    aba = pacote.aba(config.ABA_TERMO)
    primeira = config.TERMO_PRIMEIRA_LINHA
    ultima = primeira + max(len(linhas_termo), 1) - 1

    #  Todo o intervalo derramado compartilha a formatação da linha âncora.
    #  Usá-la em todas as linhas é o que mantém data como data e valor como
    #  moeda no arquivo entregue.
    modelo = (aba.estilos_da_linha(primeira)
              or aba.estilos_da_linha(primeira + 1) or {})

    #  A coluna A do template guarda um "." de marcação de área. Mantém.
    ancora_a = re.search(r'<c r="A\d+"[^>]*>.*?</c>',
                         aba.linhas.get(primeira, ("", ""))[1] or "", re.S)

    #  As linhas de cabeçalho vão inteiras, como estão — menos o cache dos
    #  totalizadores C2, G2 e H2, que ainda é o do template.
    totais = totais or {}
    partes = []
    for numero in sorted(aba.linhas):
        if numero >= primeira:
            continue
        attrs, corpo = aba.linhas[numero]
        #  A referência já carrega o número da linha, então só casa na linha
        #  certa; nas outras ``_atualizar_cache`` devolve o corpo intacto.
        for referencia, campo in config.TERMO_TOTAIS.items():
            corpo = _atualizar_cache(corpo, referencia, totais.get(campo))
        partes.append(f"<row{attrs}>{corpo}</row>")

    #  ``[None]`` é o Termo vazio: uma passada só, que escreve a âncora com a
    #  fórmula e sem valor em cache.
    for deslocamento, linha in enumerate(linhas_termo or [None]):
        numero = primeira + deslocamento
        celulas = []
        if ancora_a:
            celulas.append(re.sub(r'r="A\d+"', f'r="A{numero}"', ancora_a.group(0)))
        for letra, campo in config.TERMO_COLUNAS.items():
            valor = None if linha is None else linha[campo]
            if letra == "B" and valor is not None:
                valor = str(valor).upper()
            formula = FORMULAS_TERMO[letra] if deslocamento == 0 else None
            array_ref = (f"{letra}{primeira}:{letra}{ultima}"
                         if deslocamento == 0 else None)
            celulas.append(_celula(f"{letra}{numero}", valor,
                                   estilo=modelo.get(letra),
                                   formula=formula, array_ref=array_ref))
        partes.append(f'<row r="{numero}" spans="1:13">{"".join(celulas)}</row>')

    pacote.gravar_aba(aba, aba.xml_com("".join(partes), f"A1:M{max(ultima, 4)}"))
    log(f"  Aba {config.ABA_TERMO}: {len(linhas_termo)} linhas "
        f"({primeira} a {ultima}); fórmulas FILTER ancoradas em "
        f"B{primeira}:H{primeira}."
        + ("  ← Termo vazio: só a fórmula, sem valores."
           if not linhas_termo else ""),
        "aviso" if not linhas_termo else "info")
    return ultima


# ---------------------------------------------------------------------------
#  Aba INFORMAÇÕES
# ---------------------------------------------------------------------------
def _escrever_informacoes(pacote, parametros, totais=None, log=None):
    """
    Grava os parâmetros da rodada. Juros e multa não são tocados.

    ``D12`` (quantidade) e ``C19`` (valor total) são fórmulas que leem o Termo:
    a fórmula fica, o cache é atualizado. Sem isso o arquivo entregue anunciava
    ``QTD RECOMPRA`` = 680 sobre um Termo de 2.638 linhas.
    """
    log = log or (lambda *a, **k: None)
    aba = pacote.aba(config.ABA_INFORMACOES)
    valores = {
        config.CEL_DATA_RECOMPRA: parametros.data_recompra,
        config.CEL_NUMERO_RODADA: parametros.numero_rodada,
        config.CEL_JANELA_INICIO: parametros.janela_inicio,
        config.CEL_JANELA_FIM: parametros.janela_fim,
    }
    corpo = aba.corpo
    for referencia, valor in valores.items():
        if valor is None:
            continue
        corpo = _substituir_celula(corpo, referencia, valor)

    totais = totais or {}
    for referencia, campo in config.INFORMACOES_TOTAIS.items():
        corpo = _atualizar_cache(corpo, referencia, totais.get(campo))

    pacote.gravar_aba(aba, aba.xml_com(corpo))
    log(f"  Aba {config.ABA_INFORMACOES}: rodada {parametros.numero_rodada}, "
        f"recompra {nz.br(parametros.data_recompra)}, janela "
        f"{nz.br(parametros.janela_inicio)}–{nz.br(parametros.janela_fim)}, "
        f"QTD {totais.get('quantidade')}, total R$ "
        f"{totais.get('total_recompra') or 0:,.2f}"
        .replace(",", "@").replace(".", ",").replace("@", ".") + ".")


def _atualizar_cache(corpo, referencia, valor):
    """
    Reescreve o **valor em cache** de uma célula de fórmula, mantendo a fórmula.

    A estrutura do arquivo é congelada, então totalizadores como
    ``COUNTA(C5:C99984)`` e ``SUBTOTAL(9,H5:H1048576)`` continuam sendo fórmula.
    Só que fórmula no ``.xlsx`` guarda dois pedaços: a expressão (``<f>``) e o
    último resultado calculado (``<v>``). Ao copiar o template, a expressão vem
    certa e o resultado vem **do template** — 680 títulos e R$ 894.362,31, os
    números da rodada que serviu de molde.

    O Excel corrige isso ao abrir, porque ``forcar_recalculo`` liga o
    ``fullCalcOnLoad``. Quem lê o arquivo por fora, não: os BIs da casa e
    ``openpyxl`` com ``data_only=True`` leem o cache e acreditam nele. Foi assim
    que a rodada 14 saiu com ``QTD RECOMPRA`` = 680 no arquivo entregue.

    Devolve o corpo sem mexer em nada quando a célula não existe ou não tem
    fórmula — aí não há cache a corrigir e inventar um seria pior.
    """
    padrao = re.compile(rf'<c r="{referencia}"(?P<attrs>[^>]*?)>(?P<body>.*?)</c>',
                        re.S)
    achado = padrao.search(corpo)
    if not achado:
        return corpo
    formula = re.search(r"<f\b.*?</f>|<f\b[^>]*/>", achado.group("body"), re.S)
    if not formula:
        return corpo

    attrs = achado.group("attrs") or ""
    #  O tipo antigo sai: um cache numérico não pode ficar marcado como erro
    #  (``t="e"``) nem como texto de uma execução anterior.
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    attrs = re.sub(r'\s+vm="[^"]*"', "", attrs)
    if isinstance(valor, str):
        attrs += ' t="str"'
        cache = f"<v>{_escapar(valor)}</v>"
    elif valor is None:
        cache = ""
    else:
        cache = f"<v>{_conteudo_numerico(valor)}</v>"
    nova = f'<c r="{referencia}"{attrs}>{formula.group(0)}{cache}</c>'
    return corpo[:achado.start()] + nova + corpo[achado.end():]


def _totais_do_resultado(resultado):
    """Os números que os totalizadores do template calculam."""
    return {
        "quantidade": len(resultado.linhas_termo),
        "total_nominal": resultado.total_nominal,
        "total_recompra": resultado.total_recompra,
    }


def _substituir_celula(corpo, referencia, valor):
    """Troca o conteúdo de uma célula que já existe, preservando o estilo."""
    padrao = re.compile(rf'<c r="{referencia}"(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</c>)',
                        re.S)
    achado = padrao.search(corpo)
    estilo = None
    if achado:
        s = re.search(r'\bs="(\d+)"', achado.group("attrs") or "")
        estilo = s.group(1) if s else None
    nova = _celula(referencia, valor, estilo=estilo)
    if achado:
        return corpo[:achado.start()] + nova + corpo[achado.end():]
    return corpo


# ---------------------------------------------------------------------------
#  Ponto de entrada
# ---------------------------------------------------------------------------
def nome_arquivo(resultado):
    """Nome conforme a regra de ``INFORMAÇÕES!C6``, com o valor efetivo."""
    parametros = resultado.parametros
    milhoes = (resultado.total_recompra / Decimal(1_000_000)).quantize(Decimal("0.01"))
    nome = config.PADRAO_NOME_SAIDA.format(
        n=parametros.numero_rodada,
        valor=f"{milhoes:.2f}".replace(".", ","),
        mes_ano=parametros.data_recompra.strftime("%m-%Y")
        if parametros.data_recompra else "sem-data",
    )
    if resultado.modo_simulacao:
        nome = nome.replace(".xlsx", "_SIMULACAO.xlsx")
    return nome


def gerar(resultado, pasta_destino, log=None):
    """Copia o template, preenche as abas e grava o arquivo da rodada."""
    log = log or (lambda *a, **k: None)
    destino = Path(pasta_destino) / nome_arquivo(resultado)

    log(f"  Copiando o template {Path(resultado.template).name}…")
    pacote = Pacote(resultado.template)

    titulos = [c.titulo for c in resultado.conciliacao.casamentos]
    _escrever_aba_dados(pacote, config.ABA_VORTX,
                        [t.bruto for t in titulos], FORMULAS_VORTX, log=log)
    _escrever_aba_dados(pacote, config.ABA_GRAFENO,
                        [b.bruto for b in resultado.grafeno.registros],
                        FORMULAS_GRAFENO, log=log)
    totais = _totais_do_resultado(resultado)
    _escrever_termo(pacote, resultado.linhas_termo, totais=totais, log=log)
    _escrever_informacoes(pacote, resultado.parametros, totais=totais, log=log)
    pacote.forcar_recalculo()

    destino.parent.mkdir(parents=True, exist_ok=True)
    pacote.salvar(destino)
    tamanho = destino.stat().st_size / (1024 * 1024)
    log(f"  Arquivo gravado: {destino.name} ({tamanho:.1f} MB)", "ok")
    return destino


# ---------------------------------------------------------------------------
#  Preparação do template a partir do arquivo oficial
# ---------------------------------------------------------------------------
def preparar_template(arquivo_oficial, destino=None, log=None):
    """
    Monta ``templates/Leve_Saude_Recompra_TEMPLATE.xlsx`` a partir de um arquivo
    de recompra real.

    Mantém abas, colunas, cabeçalhos, estilos e a aba INFORMAÇÕES inteira;
    esvazia as linhas de dado das abas VORTX, GRAFENO e TERMO, guardando a
    linha 2 (e a 5, no Termo) como modelo de estilo. O arquivo original **não**
    é tocado — a leitura é sobre a cópia (regras 12-A e 15.8).
    """
    log = log or (lambda *a, **k: None)
    destino = Path(destino or config.TEMPLATE_PADRAO)
    destino.parent.mkdir(parents=True, exist_ok=True)

    log(f"Preparando template a partir de {Path(arquivo_oficial).name}…")
    pacote = Pacote(arquivo_oficial)

    for nome_aba, formulas, primeira_util in (
        (config.ABA_VORTX, FORMULAS_VORTX, 2),
        (config.ABA_GRAFENO, FORMULAS_GRAFENO, 2),
    ):
        aba = pacote.aba(nome_aba)
        partes = []
        for numero in (1, primeira_util):
            if numero not in aba.linhas:
                continue
            attrs, corpo = aba.linhas[numero]
            if numero >= primeira_util:
                for letra, construir in formulas.items():
                    corpo = _substituir_celula(corpo, f"{letra}{numero}", None)
                    corpo = _injetar_formula(corpo, letra, numero,
                                             construir(numero), aba)
            partes.append(f"<row{attrs}>{corpo}</row>")
        pacote.gravar_aba(aba, aba.xml_com("".join(partes)))
        log(f"  {nome_aba}: linhas de dado removidas, cabeçalho e estilos mantidos.")

    aba = pacote.aba(config.ABA_TERMO)
    partes = []
    for numero in sorted(aba.linhas):
        if numero > config.TERMO_PRIMEIRA_LINHA:
            continue
        attrs, corpo = aba.linhas[numero]
        if numero == config.TERMO_PRIMEIRA_LINHA:
            for letra, formula in FORMULAS_TERMO.items():
                corpo = _injetar_formula(corpo, letra, numero, formula, aba,
                                         array_ref=f"{letra}{numero}:{letra}{numero}")
        partes.append(f"<row{attrs}>{corpo}</row>")
    pacote.gravar_aba(aba, aba.xml_com("".join(partes)))
    log(f"  {config.ABA_TERMO}: linhas de dado removidas, fórmulas do ajuste 2 "
        f"aplicadas.")

    pacote.forcar_recalculo()
    pacote.salvar(destino)
    log(f"Template pronto: {destino} "
        f"({destino.stat().st_size / 1024:.0f} KB)", "ok")
    return destino


def _injetar_formula(corpo, letra, numero, formula, aba, array_ref=None):
    """Troca uma célula por sua fórmula, mantendo o estilo que ela já tinha."""
    referencia = f"{letra}{numero}"
    estilos = aba.estilos_da_linha(numero)
    nova = _celula(referencia, None, estilo=estilos.get(letra),
                   formula=formula, array_ref=array_ref)
    padrao = re.compile(rf'<c r="{referencia}"[^>]*?(?:/>|>.*?</c>)', re.S)
    if padrao.search(corpo):
        return padrao.sub(nova, corpo, count=1)
    return _inserir_ordenado(corpo, nova, letra)


def _inserir_ordenado(corpo, celula_nova, letra):
    """Insere uma célula mantendo a ordem por coluna dentro da linha."""
    celulas = re.findall(r'<c\b[^>]*?(?:/>|>.*?</c>)', corpo, re.S)
    celulas.append(celula_nova)

    def ordem(tag):
        achado = re.search(r'r="([A-Z]+)\d+"', tag)
        return indice_coluna(achado.group(1)) if achado else 0

    return "".join(sorted(celulas, key=ordem))
