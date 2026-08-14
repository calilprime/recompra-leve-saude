# -*- coding: utf-8 -*-
"""
Bateria de validações (seção 9).
================================

Roda antes de gerar qualquer arquivo. Cada item devolve ``OK``, ``AVISO`` ou
``ERRO``. Havendo ``ERRO``, a interface não habilita o botão de gerar — regra
15.5, sem exceção.

Cada validação existe por causa de um defeito real das rodadas 11ª e 12ª. O
campo ``porque`` guarda essa origem, e ela aparece na tela: quem opera precisa
saber o que a trava está protegendo, não só que ela travou.
"""

from decimal import Decimal

from core import calculo, config
from core import normalizacao as nz


class Validacao:
    """Um item da bateria."""

    __slots__ = ("chave", "titulo", "nivel", "detalhe", "porque", "dados")

    def __init__(self, chave, titulo, nivel, detalhe="", porque="", dados=None):
        self.chave = chave
        self.titulo = titulo
        self.nivel = nivel          # OK | AVISO | ERRO
        self.detalhe = detalhe
        self.porque = porque
        self.dados = dados or {}

    def como_dicionario(self):
        return {"chave": self.chave, "titulo": self.titulo, "nivel": self.nivel,
                "detalhe": self.detalhe, "porque": self.porque}

    def __repr__(self):
        return f"<{self.nivel} {self.chave}: {self.detalhe}>"


def rodar(resultado, log=None):
    """Executa a bateria inteira e devolve a lista de :class:`Validacao`."""
    log = log or (lambda *a, **k: None)
    itens = []
    itens += _bloqueantes(resultado)
    itens += _avisos(resultado)

    for item in itens:
        nivel = {"OK": "ok", "AVISO": "aviso", "ERRO": "erro"}[item.nivel]
        marca = {"OK": "✔", "AVISO": "⚠", "ERRO": "✖"}[item.nivel]
        log(f"  {marca} {item.titulo}: {item.detalhe}", nivel)

    erros = sum(1 for i in itens if i.nivel == "ERRO")
    avisos = sum(1 for i in itens if i.nivel == "AVISO")
    if erros:
        log(f"  {erros} validação(ões) em ERRO — emissão bloqueada.", "erro")
    else:
        log(f"  Nenhum erro. {avisos} aviso(s).", "ok")
    return itens


# ---------------------------------------------------------------------------
#  Bloqueiam a emissão
# ---------------------------------------------------------------------------
def _bloqueantes(r):
    itens = []
    simulacao = r.modo_simulacao

    def nivel_erro():
        return "AVISO" if simulacao else "ERRO"

    # 1. Extração única -----------------------------------------------------
    datas = r.vortx.datas_geracao
    if len(datas) == 1:
        itens.append(Validacao(
            "extracao_unica", "Extração única",
            "OK", f"Uma só DataGeracao: {nz.br(datas[0])}.",
            porque="Duas extrações misturadas na aba VORTX cobraram 5 títulos "
                   "em duplicidade na 12ª rodada (defeito 2.3)."))
    else:
        rotulo = ", ".join(nz.br(d) for d in datas) or "nenhuma"
        itens.append(Validacao(
            "extracao_unica", "Extração única", nivel_erro(),
            f"{len(datas)} datas de geração na base Vórtx: {rotulo}.",
            porque="Duas extrações misturadas na aba VORTX cobraram 5 títulos "
                   "em duplicidade na 12ª rodada (defeito 2.3). "
                   "Extraia de novo, com uma única DataGeracao.",
            dados={"datas": [nz.iso(d) for d in datas]}))

    # 2. Sem duplicidade de título ------------------------------------------
    duplicados = r.conciliacao.ids_duplicados
    itens.append(Validacao(
        "sem_duplicidade", "Sem duplicidade de título",
        "OK" if not duplicados else nivel_erro(),
        "Nenhum IdTituloVortx repetido após a deduplicação."
        if not duplicados else
        f"{len(duplicados)} IdTituloVortx repetidos: "
        f"{', '.join(duplicados[:8])}{'…' if len(duplicados) > 8 else ''}",
        porque="Título repetido vira linha repetida no Termo e cobrança dobrada.",
        dados={"ids": duplicados}))

    # 3. Sem recompra repetida ----------------------------------------------
    #  A trava já tirou esses títulos do Termo. Mesmo assim bloqueia: a fórmula
    #  da planilha não conhece o histórico e os traria de volta no recálculo.
    #  Ou o boleto é baixado na Grafeno, ou o caso vai para o relatório de
    #  exceções e alguém decide (decisão 2.2).
    repetidos = [c.titulo.id_titulo for c, motivo in r.selecao.barrados
                 if "histórico" in motivo]
    itens.append(Validacao(
        "sem_recompra_repetida", "Sem recompra repetida",
        "OK" if not repetidos else nivel_erro(),
        "Nenhum título elegível consta do histórico."
        if not repetidos else
        f"{len(repetidos)} títulos elegíveis já foram recomprados antes: "
        f"{', '.join(repetidos[:8])}{'…' if len(repetidos) > 8 else ''}",
        porque="O título 165955407 (Cíntia Farias Cordeiro) foi cobrado três "
               "vezes entre a 11ª e a 12ª porque o boleto não foi baixado na "
               "Grafeno e ele continuou 'Aberta (Vencida)'.",
        dados={"ids": repetidos}))

    # 4. Sem linha repetida no Termo ----------------------------------------
    ids = [l["id_titulo"] for l in r.linhas_termo]
    distintos = len(set(ids))
    itens.append(Validacao(
        "termo_sem_repeticao", "Sem linha repetida no Termo",
        "OK" if len(ids) == distintos else nivel_erro(),
        f"{len(ids)} linhas para {distintos} títulos distintos."
        + ("" if len(ids) == distintos else "  ← há repetição"),
        porque="O Termo da 12ª v1.1 tinha 758 linhas para 753 títulos — "
               "R$ 9.789,13 cobrados em duplicidade."))

    # 5. Defasagem zero -----------------------------------------------------
    dias = r.defasagem_dias
    detalhe = (
        f"Vórtx {nz.br(r.data_extracao_vortx)} · Grafeno "
        f"{nz.br(r.data_extracao_grafeno)} · recompra "
        f"{nz.br(r.parametros.data_recompra)}."
    )
    if getattr(r, "defasagem_problema", None):
        itens.append(Validacao(
            "defasagem_zero", "Defasagem zero", nivel_erro(),
            f"{r.defasagem_problema} {detalhe}",
            porque="Sem as duas datas de extração não há como saber se a base "
                   "está no dia. Defasagem não medida não é defasagem zero."))
    elif dias == 0:
        itens.append(Validacao(
            "defasagem_zero", "Defasagem zero", "OK",
            "Extração no dia da recompra. " + detalhe,
            porque="Cada dia útil de defasagem custa ~20 títulos e ~R$ 22 mil "
                   "em pagamento duplicado (seção 2.4)."))
    else:
        itens.append(Validacao(
            "defasagem_zero", "Defasagem zero", nivel_erro(),
            f"{calculo.frase_defasagem(dias)} {detalhe}",
            porque="Na 12ª a extração foi de sexta 07/08 e a recompra na segunda "
                   "10/08 — um dia útil. Nesse intervalo 20 sacados pagaram o "
                   "próprio boleto: R$ 25.409,69 a devolver. Extraia no dia.",
            dados=calculo.custo_defasagem(dias)))

    # 6. Convergência Python <-> Excel --------------------------------------
    itens.append(_convergencia(r, nivel_erro()))

    # 7. Soma confere -------------------------------------------------------
    soma = calculo.somar(l["valor_recompra"] for l in r.linhas_termo)
    bate = abs(soma - r.total_recompra) < Decimal("0.005")
    itens.append(Validacao(
        "soma_confere", "Soma confere",
        "OK" if bate else nivel_erro(),
        f"Total do Termo R$ {calculo.arredondar(r.total_recompra)} "
        f"= soma título a título." if bate else
        f"Total R$ {r.total_recompra} ≠ soma R$ {soma}.",
        porque="O nome do arquivo da 12ª dizia 0,59MM enquanto o Termo somava "
               "R$ 0,99 MM."))

    # 8. Janela quinzenal ---------------------------------------------------
    inicio, fim = r.parametros.janela_inicio, r.parametros.janela_fim
    valida = calculo.janela_valida(inicio, fim)
    itens.append(Validacao(
        "janela_quinzenal", "Janela quinzenal",
        "OK" if valida else nivel_erro(),
        f"{nz.br(inicio)} a {nz.br(fim)}" + ("" if valida else
        "  ← fora do padrão dia 1–15 ou 16–fim do mês"),
        porque="Decisão 1.4: janela quinzenal estrita."))

    # 9. Parâmetros preenchidos ---------------------------------------------
    faltando = []
    p = r.parametros
    for rotulo, valor in (("data de recompra", p.data_recompra),
                          ("número da rodada", p.numero_rodada),
                          ("início da janela", p.janela_inicio),
                          ("fim da janela", p.janela_fim),
                          ("juros mora", p.juros_mora),
                          ("multa", p.multa)):
        if valor in (None, ""):
            faltando.append(rotulo)
    itens.append(Validacao(
        "parametros", "Parâmetros preenchidos",
        "OK" if not faltando else nivel_erro(),
        f"Rodada {p.numero_rodada} · recompra {nz.br(p.data_recompra)} · "
        f"juros {(p.juros_mora or 0) * 100:.6f}% a.d. · multa "
        f"{(p.multa or 0) * 100:.2f}%." if not faltando else
        "Falta preencher: " + ", ".join(faltando),
        porque="Juros e multa vêm sempre da aba INFORMAÇÕES (decisão 1.5)."))

    return itens


# ---------------------------------------------------------------------------
#  Apenas avisam
# ---------------------------------------------------------------------------
def _avisos(r):
    itens = []
    c = r.conciliacao

    # Cobertura da conciliação ----------------------------------------------
    cobertura = c.cobertura * 100
    itens.append(Validacao(
        "cobertura", "Cobertura da conciliação",
        "OK" if cobertura >= 95 else "AVISO",
        f"{len(c.casados)} de {len(c.casamentos)} títulos com boleto "
        f"({cobertura:.2f}%).",
        porque="Queda de cobertura costuma indicar mudança de layout ou "
               "extração parcial de um dos lados."))

    # Divergências de valor --------------------------------------------------
    divergentes = (c.por_causa(config.CAUSA_DIVERGENCIA_VALOR)
                   + c.por_causa(config.CAUSA_TITULO_IRMAO))
    soma = calculo.somar(x.titulo.valor_nominal for x in divergentes)
    irmaos = len(c.por_causa(config.CAUSA_TITULO_IRMAO))
    itens.append(Validacao(
        "divergencia_valor", "Divergências de valor",
        "OK" if not divergentes else "AVISO",
        f"{len(divergentes)} títulos, R$ {calculo.arredondar(soma)} "
        f"({irmaos} classificados como TITULO_IRMAO)." if divergentes else
        "Nenhuma divergência de valor.",
        porque="Em 192 de 227 casos o sacado tem dois títulos no mesmo "
               "vencimento e o órfão é o de emissão mais antiga (seção 3).",
        dados={"quantidade": len(divergentes), "valor": str(soma)}))

    # Múltiplos boletos ------------------------------------------------------
    multiplos = c.multiplos_boletos
    linhas = sum(x.qtd_boletos for x in multiplos)
    pelo_numero = c.contagens.get("desempate_por_numero", 0)
    mudou_status = len(c.desempates_divergentes)
    itens.append(Validacao(
        "multiplos_boletos", "Múltiplos boletos",
        "OK" if not multiplos else "AVISO",
        f"{len(multiplos)} títulos com mais de um boleto "
        f"({linhas} linhas de boleto): {pelo_numero} resolvidos pelo Nosso_Número "
        f"e {len(multiplos) - pelo_numero} pela prioridade de status"
        + (f"; em {mudou_status} o Nosso_Número apontou status diferente do que a "
           f"prioridade escolheria." if mudou_status else ".")
        if multiplos else "Nenhuma chave com mais de um boleto.",
        porque="O relatório informa sempre as duas contagens — títulos e linhas "
               "de boleto — para a conferência entre pessoas não divergir "
               "(seção 2.6: '7 linhas' eram 5 títulos)."))

    # Resíduo de rodadas anteriores -----------------------------------------
    residuo = r.residuo
    soma_residuo = calculo.somar(x.titulo.valor_nominal for x in residuo)
    itens.append(Validacao(
        "residuo_anterior", "Resíduo de rodadas anteriores",
        "OK" if not residuo else "AVISO",
        f"{len(residuo)} títulos 'Aberta (Vencida)' vencidos antes de "
        f"{nz.br(r.parametros.janela_inicio)}, R$ {calculo.arredondar(soma_residuo)}."
        if residuo else "Nenhum título vencido sobrou de janelas anteriores.",
        porque="São candidatos a terem escapado de uma rodada passada — o "
               "passivo da 11ª foi exatamente isso.",
        dados={"quantidade": len(residuo), "valor": str(soma_residuo)}))

    # Continuidade da janela -------------------------------------------------
    anterior_inicio, anterior_fim = r.historico.janela_anterior()
    if anterior_fim is None:
        itens.append(Validacao(
            "continuidade", "Continuidade da janela", "OK",
            "Primeira rodada no histórico — nada a comparar.",
            porque="A continuidade só existe a partir da segunda rodada."))
    else:
        esperado_inicio, _ = calculo.janela_seguinte(anterior_fim)
        if r.parametros.janela_inicio == esperado_inicio:
            itens.append(Validacao(
                "continuidade", "Continuidade da janela", "OK",
                f"Emenda com a rodada anterior, que terminou em "
                f"{nz.br(anterior_fim)}.",
                porque="Lacuna deixa título para trás; sobreposição cobra duas vezes."))
        else:
            relacao = ("lacuna" if r.parametros.janela_inicio > esperado_inicio
                       else "sobreposição")
            itens.append(Validacao(
                "continuidade", "Continuidade da janela", "AVISO",
                f"Há {relacao}: a rodada anterior terminou em "
                f"{nz.br(anterior_fim)} e esta começa em "
                f"{nz.br(r.parametros.janela_inicio)} "
                f"(esperado {nz.br(esperado_inicio)}).",
                porque="Lacuna deixa título para trás; sobreposição cobra duas vezes."))

    # Boletos Grafeno sem título Vórtx ---------------------------------------
    sem_titulo = c.boletos_sem_titulo
    na_janela = [b for b in sem_titulo
                 if b.status == config.STATUS_RECOMPRAVEL
                 and b.vencimento
                 and r.parametros.janela_inicio <= b.vencimento <= r.parametros.janela_fim]
    soma_janela = calculo.somar(b.valor for b in na_janela)
    itens.append(Validacao(
        "grafeno_sem_vortx", "Boletos Grafeno sem título Vórtx", "OK",
        f"{len(sem_titulo)} boletos sem título na Vórtx; {len(na_janela)} deles "
        f"'Aberta (Vencida)' na janela, R$ {calculo.arredondar(soma_janela)}. "
        f"Informativo.",
        porque="Não existe campo que separe a carteira cedida no export da "
               "Grafeno (seção 3.2). A Vórtx é a fonte de verdade do escopo, "
               "então este número não gera ação.",
        dados={"quantidade": len(sem_titulo), "na_janela": len(na_janela)}))

    # Acentuação -------------------------------------------------------------
    perdas = [a for a in r.vortx.avisos if "acentuação" in a]
    if perdas:
        itens.append(Validacao(
            "acentuacao", "Acentuação da base Vórtx", "AVISO", perdas[0],
            porque="O export do portal corrompe acentos. Ler do banco resolve "
                   "(seção 8.1) — os nomes vão para o Termo assim como estão."))

    return itens


# ---------------------------------------------------------------------------
#  Convergência entre os dois caminhos de cálculo
# ---------------------------------------------------------------------------
def _convergencia(r, nivel_erro):
    """
    Compara o Termo do motor com o que as fórmulas corrigidas produziriam.

    São dois caminhos independentes:

    * **motor** — chave normalizada, deduplicação, prioridade de status,
      passada de tolerância e trava de histórico;
    * **planilha** — a simulação literal de ``CONT.SES`` sobre a coluna inteira
      da aba GRAFENO e da condição ``AS=1``, exatamente como as fórmulas novas
      de ``VORTX!AR2`` e ``TERMO!B5:H5`` vão calcular quando o Excel abrir.

    Qualquer diferença bloqueia, e a razão é prática: o arquivo entregue leva as
    fórmulas dentro. Quando alguém abrir, o Excel recalcula e o Termo passa a
    ser o da fórmula. Se os dois não forem idênticos na emissão, o que sai da
    casa não é o que o motor apurou. Cada diferença vem com a causa, e cada
    causa tem a sua própria validação dizendo o que corrigir.
    """
    do_motor = {l["id_titulo"] for l in r.linhas_termo}
    da_planilha = simular_formulas(r)

    so_na_planilha = da_planilha - do_motor
    so_no_motor = do_motor - da_planilha

    if not so_na_planilha and not so_no_motor:
        return Validacao(
            "convergencia", "Convergência Python ↔ Excel", "OK",
            f"Os dois caminhos convergem: {len(do_motor)} títulos.",
            porque="São dois caminhos de cálculo independentes que devem dar o "
                   "mesmo Termo; divergência bloqueia (seção 5.3).")

    causas = _causas_divergencia(r, so_na_planilha)
    partes = []
    if so_na_planilha:
        partes.append(f"{len(so_na_planilha)} títulos que a fórmula traria e o "
                      f"motor descartou ({causas})")
    if so_no_motor:
        partes.append(f"{len(so_no_motor)} títulos que só o motor encontrou "
                      f"({', '.join(sorted(so_no_motor)[:6])}) — casados pela "
                      f"passada de tolerância, que a fórmula não faz")
    return Validacao(
        "convergencia", "Convergência Python ↔ Excel", nivel_erro,
        "Os caminhos divergem: " + "; ".join(partes) + ".",
        porque="O arquivo entregue leva as fórmulas dentro e recalcula ao "
               "abrir. Emitir com divergência faria o Termo entregue diferir do "
               "apurado (seção 5.3).",
        dados={"so_na_planilha": sorted(so_na_planilha)[:50],
               "so_no_motor": sorted(so_no_motor)[:50]})


def _causas_divergencia(r, ids):
    """Diz por que o motor descartou cada título que a fórmula traria."""
    if not ids:
        return ""
    duplicados = {d.id_titulo for _m, d in r.conciliacao.duplicados_vortx}
    barrados = {}
    for casamento, motivo in r.selecao.barrados:
        barrados[casamento.titulo.id_titulo] = motivo
    contagem = {}
    for identificador in ids:
        if identificador in barrados:
            motivo = barrados[identificador]
        elif identificador in duplicados:
            motivo = "linha repetida na base Vórtx"
        else:
            motivo = "causa não identificada — investigar"
        contagem[motivo] = contagem.get(motivo, 0) + 1
    return "; ".join(f"{n}× {motivo}" for motivo, n in sorted(contagem.items()))


def simular_formulas(r):
    """
    Reproduz em Python o que a planilha corrigida calcularia.

    Sem deduplicação, sem histórico, sem passada de tolerância — a planilha não
    tem nada disso. Só a chave concatenada, o ``CONT.SES`` de coluna inteira e a
    janela.
    """
    inicio, fim = r.parametros.janela_inicio, r.parametros.janela_fim

    #  GRAFENO!AE = CONCAT(Número_Documento; DATEVALUE(Data_Vencimento); VALUE(Valor_Cobrança))
    #  CONT.SES sobre a coluna inteira, contando por status — e, na primeira
    #  metade da fórmula, também pelo Nosso_Número.
    contagem = {}
    por_ligacao = {}
    for boleto in r.grafeno.registros:
        chave = _chave_planilha(boleto.documento, boleto.vencimento, boleto.valor)
        alvo = contagem.setdefault(chave, {})
        alvo[boleto.status] = alvo.get(boleto.status, 0) + 1
        if boleto.chave_boleto:
            ligado = por_ligacao.setdefault((chave, boleto.chave_boleto), {})
            ligado[boleto.status] = ligado.get(boleto.status, 0) + 1

    #  A planilha percorre a aba VORTX inteira — sem deduplicar, sem histórico.
    escolhidos = set()
    for titulo in r.vortx.registros:
        chave = _chave_planilha(titulo.documento, titulo.vencimento,
                                titulo.valor_nominal)
        por_status = contagem.get(chave)
        if not por_status:
            continue                                        # AR = "N/A"
        #  Mesma ordem da fórmula: a ligação pelo Nosso_Número primeiro, a
        #  prioridade de status depois.
        pela_ligacao = por_ligacao.get((chave, titulo.chave_boleto)) or {}
        status = next((s for s in config.PRIORIDADE_STATUS if pela_ligacao.get(s)), None)
        if status is None:
            status = next((s for s in config.PRIORIDADE_STATUS if por_status.get(s)),
                          "N/A")
        if status != config.STATUS_RECOMPRAVEL:
            continue                                        # AS = 0
        if titulo.vencimento and inicio <= titulo.vencimento <= fim:
            escolhidos.add(titulo.id_titulo)                 # AS = 1
    return escolhidos


def _chave_planilha(documento, vencimento, valor):
    """A chave que a coluna AUX monta — documento, serial da data e valor."""
    return (f"{nz.documento(documento)}|{nz.serial_excel(vencimento)}"
            f"|{nz.valor_texto(valor)}")
