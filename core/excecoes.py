# -*- coding: utf-8 -*-
"""
Relatório de exceções (seção 10.2).
===================================

Arquivo separado, conforme a decisão 2.2. Nada disso entra na planilha de
recompra — lá a estrutura é congelada e as regras teriam de virar coluna nova.

Quem opera revisa, marca a coluna de decisão na aba ``DECISAO`` e a automação
regenera o Termo considerando as rejeições.

Duas contagens sempre juntas — títulos e linhas de boleto. Foi a falta disso que
fez o Vagner contar 7 e a apuração contar 5, sendo o mesmo fato (seção 2.6).
"""

from datetime import datetime
from decimal import Decimal
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from core import calculo, config
from core import normalizacao as nz

AZUL = "001B5C"
LARANJA = "FF965A"
CINZA = "EEF3FA"

_TITULO = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_NEGRITO = Font(name="Calibri", size=11, bold=True)
_NORMAL = Font(name="Calibri", size=10)
_FUNDO_TITULO = PatternFill("solid", fgColor=AZUL)
_FUNDO_DESTAQUE = PatternFill("solid", fgColor=CINZA)
_BORDA = Border(bottom=Side(style="thin", color="DFE6F1"))


def gerar(resultado, pasta_destino, log=None):
    """Monta o arquivo de exceções da rodada."""
    log = log or (lambda *a, **k: None)
    destino = Path(pasta_destino) / _nome(resultado)
    livro = openpyxl.Workbook()
    livro.remove(livro.active)

    _aba_resumo(livro, resultado)
    _aba_multiplos_boletos(livro, resultado)
    _aba_divergencia_valor(livro, resultado)
    _aba_sem_boleto(livro, resultado)
    _aba_barrados(livro, resultado)
    _aba_grafeno_sem_vortx(livro, resultado)
    _aba_decisao(livro, resultado)

    destino.parent.mkdir(parents=True, exist_ok=True)
    livro.save(destino)
    log(f"  Relatório de exceções: {destino.name}", "ok")
    return destino


def _nome(resultado):
    numero = resultado.parametros.numero_rodada or 0
    data = resultado.parametros.data_recompra or datetime.now().date()
    sufixo = "_SIMULACAO" if resultado.modo_simulacao else ""
    return f"Excecoes_Recompra_{numero}_{data.strftime('%Y-%m-%d')}{sufixo}.xlsx"


# ---------------------------------------------------------------------------
#  Utilidades de escrita
# ---------------------------------------------------------------------------
def _nova_aba(livro, nome, cabecalho, larguras=None):
    aba = livro.create_sheet(nome)
    aba.append(cabecalho)
    for celula in aba[1]:
        celula.font = _TITULO
        celula.fill = _FUNDO_TITULO
        celula.alignment = Alignment(horizontal="center", vertical="center",
                                     wrap_text=True)
    aba.freeze_panes = "A2"
    for posicao, largura in enumerate(larguras or [], start=1):
        aba.column_dimensions[get_column_letter(posicao)].width = largura
    return aba


def _linha(aba, valores, formatos=None):
    aba.append(valores)
    numero = aba.max_row
    for posicao, valor in enumerate(valores, start=1):
        celula = aba.cell(row=numero, column=posicao)
        celula.font = _NORMAL
        celula.border = _BORDA
        if isinstance(valor, Decimal):
            celula.value = float(valor)
            celula.number_format = '#,##0.00'
        elif formatos and posicao in formatos:
            celula.number_format = formatos[posicao]
    return numero


_DATA = "dd/mm/yyyy"


# ---------------------------------------------------------------------------
#  RESUMO
# ---------------------------------------------------------------------------
def _aba_resumo(livro, r):
    aba = _nova_aba(livro, "RESUMO", ["Item", "Valor", "Observação"],
                    larguras=[46, 24, 96])
    p = r.parametros
    c = r.conciliacao.contagens

    def secao(titulo):
        numero = _linha(aba, [titulo, "", ""])
        for celula in aba[numero]:
            celula.font = _NEGRITO
            celula.fill = _FUNDO_DESTAQUE

    secao("Parâmetros da rodada")
    _linha(aba, ["N° da recompra", p.numero_rodada, ""])
    _linha(aba, ["Data da recompra", nz.br(p.data_recompra), ""])
    _linha(aba, ["Janela", f"{nz.br(p.janela_inicio)} a {nz.br(p.janela_fim)}", ""])
    _linha(aba, ["Juros mora (a.d.)", f"{(p.juros_mora or 0) * 100:.6f}%",
                 "Lido de INFORMAÇÕES!D13 — nunca fixado no código"])
    _linha(aba, ["Multa", f"{(p.multa or 0) * 100:.2f}%", "Lido de INFORMAÇÕES!D14"])
    _linha(aba, ["Modo", "SIMULAÇÃO" if r.modo_simulacao else "Emissão",
                 "Simulação não atualiza o histórico e marca o nome do arquivo"])

    secao("Extração")
    _linha(aba, ["Fonte Vórtx", r.vortx.origem, r.vortx.caminho])
    _linha(aba, ["DataGeracao da Vórtx", nz.br(r.data_extracao_vortx),
                 "Data de posição informada pela Vórtx"])
    _linha(aba, ["Fonte Grafeno", r.grafeno.origem, r.grafeno.caminho])
    _linha(aba, ["Extração Grafeno", nz.br(r.data_extracao_grafeno), ""])
    _linha(aba, ["Defasagem até a recompra", f"{r.defasagem_dias} dia(s) útil(eis)",
                 calculo.frase_defasagem(r.defasagem_dias)])
    _linha(aba, ["Snapshot", str(r.pasta_snapshot or "—"), ""])

    secao("Contagens")
    _linha(aba, ["Linhas lidas da Vórtx", c.get("titulos_lidos"), ""])
    _linha(aba, ["Linhas removidas na deduplicação", c.get("linhas_removidas_dedup"),
                 "Mesmo IdTituloVortx em mais de uma linha"])
    _linha(aba, ["Títulos após deduplicação", c.get("titulos"), ""])
    _linha(aba, ["Boletos lidos da Grafeno", c.get("boletos"),
                 "Base varrida por inteiro, sem limite de linha"])
    _linha(aba, ["Casados na passada 1 (chave exata)", c.get("casados_passada_1"), ""])
    _linha(aba, ["Casados na passada 2 (tolerância)", c.get("casados_passada_2"),
                 f"Só quando o par (documento, vencimento) é único dos dois "
                 f"lados — tolerância de R$ {config.TOLERANCIA_VALOR}"])
    _linha(aba, ["Títulos sem boleto", c.get("nao_casados"), ""])
    _linha(aba, ["Cobertura da conciliação", f"{c.get('cobertura', 0) * 100:.2f}%", ""])
    _linha(aba, ["Títulos com mais de um boleto", c.get("multiplos_boletos"),
                 f"{sum(x.qtd_boletos for x in r.conciliacao.multiplos_boletos)} "
                 f"linhas de boleto no total"])
    _linha(aba, ["Boletos sem título na Vórtx", c.get("boletos_sem_titulo"),
                 "Informativo — a Vórtx é a fonte de verdade do escopo (seção 3.2)"])

    secao("Termo")
    _linha(aba, ["Títulos elegíveis", len(r.linhas_termo), ""])
    _linha(aba, ["Valor nominal", calculo.arredondar(r.total_nominal), ""])
    _linha(aba, ["Valor de recompra", calculo.arredondar(r.total_recompra), ""])
    _linha(aba, ["Títulos barrados", len(r.selecao.barrados), "Detalhe na aba BARRADOS"])

    secao("Bateria de validações")
    for item in r.validacoes:
        numero = _linha(aba, [item.titulo, item.nivel,
                              f"{item.detalhe}  ·  {item.porque}"])
        if item.nivel == "ERRO":
            aba.cell(row=numero, column=2).font = Font(bold=True, color="C00000")
        elif item.nivel == "AVISO":
            aba.cell(row=numero, column=2).font = Font(bold=True, color="B26B00")
        else:
            aba.cell(row=numero, column=2).font = Font(bold=True, color="1B7F4B")


# ---------------------------------------------------------------------------
#  MULTIPLOS_BOLETOS
# ---------------------------------------------------------------------------
def _aba_multiplos_boletos(livro, r):
    #  Duas colunas a mais desde a Fase 3: como o empate foi resolvido e o que a
    #  prioridade de status teria escolhido sozinha. Quando as duas divergem, a
    #  linha merece conferência — é o título mudando de dono entre irmãos.
    aba = _nova_aba(livro, "MULTIPLOS_BOLETOS", [
        "IdTituloVortx", "Sacado", "CPF/CNPJ", "Vencimento", "Valor nominal",
        "Qtd boletos", "Status encontrados", "Status escolhido",
        "Resolvido por", "Só pela prioridade seria", "Entrou no Termo?",
    ], larguras=[16, 40, 20, 13, 15, 12, 34, 20, 20, 22, 16])

    no_termo = {l["id_titulo"] for l in r.linhas_termo}
    divergentes = {id(c): p for c, p in r.conciliacao.desempates_divergentes}
    rotulos = {"numero_titulo": "Nosso_Número", "prioridade": "Prioridade de status"}

    for casamento in r.conciliacao.multiplos_boletos:
        t = casamento.titulo
        encontrados = ", ".join(
            f"{s} ({n})" for s, n in sorted(_contar_status(casamento.boletos).items()))
        pela_prioridade = divergentes.get(id(casamento))
        _linha(aba, [t.id_titulo, t.nome, t.documento_formatado,
                     t.vencimento, t.valor_nominal, casamento.qtd_boletos,
                     encontrados, casamento.status,
                     rotulos.get(casamento.desempate, "—"),
                     pela_prioridade.status if pela_prioridade else "o mesmo",
                     "Sim" if t.id_titulo in no_termo else "Não"],
               formatos={4: _DATA})
    if aba.max_row == 1:
        _linha(aba, ["— nenhuma chave com mais de um boleto —"])


def _contar_status(boletos):
    contagem = {}
    for b in boletos:
        contagem[b.status] = contagem.get(b.status, 0) + 1
    return contagem


# ---------------------------------------------------------------------------
#  DIVERGENCIA_VALOR
# ---------------------------------------------------------------------------
def _aba_divergencia_valor(livro, r):
    aba = _nova_aba(livro, "DIVERGENCIA_VALOR", [
        "Classificação", "IdTituloVortx", "Sacado", "CPF/CNPJ", "Vencimento",
        "Valor Vórtx", "Valor Grafeno", "Diferença", "Emissão do título",
        "Status do boleto", "Leitura",
    ], larguras=[18, 16, 38, 20, 13, 14, 14, 12, 15, 18, 60])

    casos = (r.conciliacao.por_causa(config.CAUSA_TITULO_IRMAO)
             + r.conciliacao.por_causa(config.CAUSA_DIVERGENCIA_VALOR))
    por_parcial = {}
    for boleto in r.grafeno.registros:
        por_parcial.setdefault(boleto.chave_parcial, []).append(boleto)

    for casamento in casos:
        t = casamento.titulo
        vizinhos = por_parcial.get(t.chave_parcial, [])
        vizinho = min(vizinhos, key=lambda b: abs((b.valor or Decimal(0))
                                                  - (t.valor_nominal or Decimal(0))),
                      default=None)
        leitura = (
            "Título irmão: o sacado tem outro título neste mesmo vencimento que "
            "casou com o boleto. O órfão costuma ser o de emissão mais antiga — "
            "provável contrato reajustado ou trocado (seção 3)."
            if casamento.causa == config.CAUSA_TITULO_IRMAO else
            "Valor de face divergente sem título irmão. Padrão observado: valor "
            "redondo na Vórtx contra valor preciso na Grafeno."
        )
        _linha(aba, [
            casamento.causa, t.id_titulo, t.nome, t.documento_formatado,
            t.vencimento, t.valor_nominal,
            vizinho.valor if vizinho else None,
            casamento.diferenca_valor, t.data_emissao,
            vizinho.status if vizinho else "", leitura,
        ], formatos={5: _DATA, 9: _DATA})
    if aba.max_row == 1:
        _linha(aba, ["— nenhuma divergência de valor —"])


# ---------------------------------------------------------------------------
#  SEM_BOLETO
# ---------------------------------------------------------------------------
def _aba_sem_boleto(livro, r):
    aba = _nova_aba(livro, "SEM_BOLETO", [
        "Causa", "IdTituloVortx", "Sacado", "CPF/CNPJ", "Vencimento",
        "Valor nominal", "Na janela?", "Emissão", "Situação Vórtx",
    ], larguras=[26, 16, 38, 20, 13, 14, 12, 13, 22])

    inicio, fim = r.parametros.janela_inicio, r.parametros.janela_fim
    causas_sem_boleto = (config.CAUSA_SACADO_INEXISTENTE,
                         config.CAUSA_DIVERGENCIA_VENCIMENTO,
                         config.CAUSA_SEM_CORRESPONDENCIA)
    for casamento in r.conciliacao.casamentos:
        if casamento.casou or casamento.causa not in causas_sem_boleto:
            continue
        t = casamento.titulo
        na_janela = bool(t.vencimento and inicio <= t.vencimento <= fim)
        _linha(aba, [casamento.causa, t.id_titulo, t.nome, t.documento_formatado,
                     t.vencimento, t.valor_nominal, "Sim" if na_janela else "Não",
                     t.data_emissao, t.situacao], formatos={5: _DATA, 8: _DATA})
    if aba.max_row == 1:
        _linha(aba, ["— todos os títulos encontraram boleto —"])


# ---------------------------------------------------------------------------
#  BARRADOS
# ---------------------------------------------------------------------------
def _aba_barrados(livro, r):
    aba = _nova_aba(livro, "BARRADOS", [
        "Motivo", "IdTituloVortx", "Sacado", "CPF/CNPJ", "Vencimento",
        "Valor nominal", "Status do boleto", "Rodada anterior", "O que fazer",
    ], larguras=[44, 16, 38, 20, 13, 14, 18, 15, 62])

    for casamento, motivo in r.selecao.barrados:
        t = casamento.titulo
        rodada = r.historico.rodada_de(t.id_titulo)
        acao = (
            "Confirmar com Operações se o boleto foi baixado na Grafeno depois "
            "da rodada anterior. Enquanto continuar 'Aberta (Vencida)', a "
            "trava impede a terceira cobrança."
            if "histórico" in motivo else
            "Refazer a extração da Vórtx com uma única DataGeracao e rodar de novo."
        )
        _linha(aba, [motivo, t.id_titulo, t.nome, t.documento_formatado,
                     t.vencimento, t.valor_nominal, casamento.status,
                     f"{rodada['numero']}ª em {nz.br(rodada['data_recompra'])}"
                     if rodada else "—", acao], formatos={5: _DATA})

    for mantido, descartado in r.conciliacao.duplicados_vortx:
        _linha(aba, ["Linha repetida removida na deduplicação",
                     descartado.id_titulo, descartado.nome,
                     descartado.documento_formatado, descartado.vencimento,
                     descartado.valor_nominal, "",
                     f"linha {descartado.linha} (mantida a {mantido.linha})",
                     "Duas extrações na mesma aba. Refazer a extração."],
               formatos={5: _DATA})
    if aba.max_row == 1:
        _linha(aba, ["— nenhum título barrado —"])


# ---------------------------------------------------------------------------
#  GRAFENO_SEM_VORTX
# ---------------------------------------------------------------------------
def _aba_grafeno_sem_vortx(livro, r):
    aba = _nova_aba(livro, "GRAFENO_SEM_VORTX", [
        "Pagador", "CPF/CNPJ", "Vencimento", "Valor", "Status", "Nosso número",
        "Seu número", "Na janela?",
    ], larguras=[40, 20, 13, 14, 18, 20, 20, 12])

    inicio, fim = r.parametros.janela_inicio, r.parametros.janela_fim
    #  Só o que interessa olhar: vencido em aberto. O resto é volume histórico.
    interessantes = [b for b in r.conciliacao.boletos_sem_titulo
                     if b.status == config.STATUS_RECOMPRAVEL]
    for boleto in sorted(interessantes,
                         key=lambda b: (b.vencimento or nz.data("1900-01-01"))):
        na_janela = bool(boleto.vencimento and inicio <= boleto.vencimento <= fim)
        _linha(aba, [boleto.nome, nz.documento_formatado(boleto.documento),
                     boleto.vencimento, boleto.valor, boleto.status,
                     boleto.nosso_numero, boleto.seu_numero,
                     "Sim" if na_janela else "Não"], formatos={3: _DATA})
    if aba.max_row == 1:
        _linha(aba, ["— nenhum boleto vencido em aberto sem título na Vórtx —"])


# ---------------------------------------------------------------------------
#  DECISAO
# ---------------------------------------------------------------------------
def _aba_decisao(livro, r):
    """
    A aba que quem opera preenche.

    Preenche ``Decisão`` com ``rejeita`` para tirar o título do Termo, ou
    ``aceita`` para mantê-lo, e roda de novo apontando este arquivo no campo
    correspondente da interface.
    """
    aba = _nova_aba(livro, "DECISAO", [
        "IdTituloVortx", "Sacado", "CPF/CNPJ", "Vencimento", "Valor nominal",
        "Valor de recompra", "Motivo de estar aqui", "Decisão (aceita/rejeita)",
        "Observação", "Quem decidiu", "Quando",
    ], larguras=[16, 38, 20, 13, 14, 16, 46, 24, 40, 22, 14])

    from openpyxl.worksheet.datavalidation import DataValidation
    validacao = DataValidation(type="list", formula1='"aceita,rejeita"',
                               allow_blank=True)
    aba.add_data_validation(validacao)

    candidatos = []
    no_termo = {l["id_titulo"]: l for l in r.linhas_termo}

    for casamento in r.conciliacao.multiplos_boletos:
        if casamento.titulo.id_titulo in no_termo:
            candidatos.append((casamento, "Título com mais de um boleto na "
                                          "Grafeno — status resolvido pela prioridade"))
    for casamento in r.conciliacao.casamentos:
        if casamento.passada == 2 and casamento.titulo.id_titulo in no_termo:
            candidatos.append((casamento, "Casado pela tolerância de valor "
                                          "(par único dos dois lados)"))
    for casamento, motivo in r.selecao.barrados:
        candidatos.append((casamento, f"Barrado: {motivo}"))

    vistos = set()
    for casamento, motivo in candidatos:
        t = casamento.titulo
        marca = (t.id_titulo, motivo)
        if marca in vistos:
            continue
        vistos.add(marca)
        linha = no_termo.get(t.id_titulo)
        numero = _linha(aba, [
            t.id_titulo, t.nome, t.documento_formatado, t.vencimento,
            t.valor_nominal, linha["valor_recompra"] if linha else None,
            motivo, "", "", "", "",
        ], formatos={4: _DATA, 11: _DATA})
        validacao.add(aba.cell(row=numero, column=8))

    if aba.max_row == 1:
        _linha(aba, ["— nenhum caso pendente de decisão —"])
