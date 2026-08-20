# -*- coding: utf-8 -*-
"""
Testes do motor de conciliação.
===============================

Roda com::

    py testes/test_conciliacao.py

Duas partes:

**Testes rápidos** — normalização, prioridade de status, janela quinzenal,
cálculo e dias úteis. Rodam em menos de um segundo, sem tocar em arquivo.

**Aceite da Fase 1** — os critérios do plano, sobre os arquivos reais das
rodadas 11ª e 12ª. Levam alguns minutos e só rodam se as amostras estiverem em
``testes/amostras/``. São eles que provam que o motor encontra o que a planilha
deixou passar:

* na 11ª, os cinco títulos que deveriam ter sido recomprados e não foram —
  Maria Izabel Basilio, Aline de Oliveira Gomes, Ana Teresa Valls Pereira,
  Irisdalva Teles de Deus e Ana Maria Silva Damasceno, R$ 10.367,84 de nominal;
* na 12ª, as sete linhas órfãs da extração antiga e os cinco títulos que o Termo
  cobrou em duplicidade.
"""

import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core import calculo, conciliacao, config, leitura, motor      # noqa: E402
from core import normalizacao as nz                                # noqa: E402
from core.fonte_grafeno import Boleto                              # noqa: E402
from core.fonte_vortx import Titulo                                # noqa: E402

AMOSTRA_11 = config.PASTA_AMOSTRAS / "Leve_Saude_11_Recompra_0,59MM_v5.1_07-2026.xlsx"
AMOSTRA_12 = config.PASTA_AMOSTRAS / "Leve_Saude_12_Recompra_0,59MM_v1.1_08-2026.xlsx"

_falhas = []


def verificar(condicao, descricao, detalhe=""):
    if condicao:
        print(f"  ok   {descricao}")
    else:
        print(f"  FALHA {descricao}" + (f"\n         {detalhe}" if detalhe else ""))
        _falhas.append(descricao)


# ---------------------------------------------------------------------------
#  Testes rápidos
# ---------------------------------------------------------------------------
def test_normalizacao():
    print("\nNormalização")
    verificar(nz.documento("349.532.217-53") == "34953221753", "CPF vira só dígitos")
    verificar(nz.documento("4.531.247-80") == "00453124780",
              "CPF sem zero à esquerda é completado")
    verificar(nz.documento("45.755.997/0001-28") == "45755997000128", "CNPJ idem")
    verificar(nz.valor("R$ 1.346,63") == Decimal("1346.63"),
              "valor em texto brasileiro")
    verificar(nz.valor(2467.92) == Decimal("2467.92"), "valor em float")
    verificar(nz.data("23/03/2026") == date(2026, 3, 23), "data dd/mm/aaaa")
    verificar(nz.data(46208) == date(2026, 7, 5), "serial do Excel")
    verificar(nz.serial_excel(date(2026, 7, 5)) == "46208", "data vira serial")
    verificar(nz.chave("349.532.217-53", date(2026, 7, 5), 2467.92)
              == "34953221753|2026-07-05|2467.92", "chave canônica")
    verificar(nz.corrigir_acentuacao("JOSÃ‰ DA SILVA") == "JOSÉ DA SILVA",
              "acentuação recuperável é consertada")
    verificar(nz.tem_perda_acentuacao("PLANOS DE SA�DE"),
              "perda irrecuperável é detectada")


def test_prioridade_status():
    print("\nPrioridade de status (decisão 1.2)")

    def boleto(status, linha):
        return Boleto(status=status, linha=linha, documento="1", valor=Decimal(1))

    #  O PROCX devolvia o primeiro da planilha. Aqui o Baixada vem antes e
    #  mesmo assim tem de perder — foi o defeito 2.2.
    escolhido = conciliacao.escolher_boleto([boleto("Baixada", 10),
                                             boleto("Aberta (Vencida)", 900)])
    verificar(escolhido.status == "Aberta (Vencida)",
              "Aberta (Vencida) vence Baixada mesmo estando depois na planilha")
    escolhido = conciliacao.escolher_boleto([boleto("Paga", 1), boleto("Aberta", 2)])
    verificar(escolhido.status == "Aberta", "Aberta vence Paga")


def test_desempate_por_numero_titulo():
    print("\nDesempate pelo Nosso_Número (Fase 3)")

    def boleto(status, linha, nosso=""):
        return Boleto(status=status, linha=linha, documento="1", valor=Decimal(1),
                      nosso_numero=nosso, chave_boleto=nz.so_digitos(nosso))

    def titulo(numero):
        return Titulo(id_titulo="1", linha=2, numero_titulo=numero,
                      chave_boleto=nz.chave_boleto(numero))

    verificar(nz.chave_boleto("382477119") == "403382477119",
              "NumeroTitulo vira o Nosso_Número esperado")
    verificar(nz.chave_boleto("") == "" and nz.chave_boleto(None) == "",
              "sem NumeroTitulo não há ligação")

    #  O caso que interessa: a prioridade escolheria o 'Aberta (Vencida)', mas o
    #  boleto daquele título é o 'Baixada'. O outro é do título irmão.
    do_titulo = boleto("Baixada", 10, "403382477119")
    do_irmao = boleto("Aberta (Vencida)", 11, "403382477222")
    escolhido = conciliacao.escolher_boleto([do_irmao, do_titulo],
                                            titulo("382477119"))
    verificar(escolhido is do_titulo,
              "o Nosso_Número vence a prioridade de status")

    #  Sem ligação, a decisão 1.2 continua valendo inteira.
    escolhido = conciliacao.escolher_boleto(
        [boleto("Baixada", 10), boleto("Aberta (Vencida)", 11)], titulo("999"))
    verificar(escolhido.status == "Aberta (Vencida)",
              "sem ligação, a prioridade de status decide")

    #  A ligação nunca amplia o conjunto: se o boleto do título não está entre
    #  os candidatos, não se inventa nada.
    escolhido = conciliacao.escolher_boleto([do_irmao], titulo("382477119"))
    verificar(escolhido is do_irmao,
              "a ligação só escolhe entre os candidatos que a chave já casou")

    #  Dois boletos com o mesmo Nosso_Número seria ambiguidade — cai na prioridade.
    a = boleto("Baixada", 10, "403382477119")
    b = boleto("Aberta (Vencida)", 11, "403382477119")
    escolhido = conciliacao.escolher_boleto([a, b], titulo("382477119"))
    verificar(escolhido is b, "ligação ambígua volta para a prioridade")


def test_deduplicacao():
    print("\nDeduplicação da Vórtx (defeito 2.3)")

    def titulo(linha, geracao):
        return Titulo(id_titulo="143950045", linha=linha, data_geracao=geracao)

    #  As sete linhas órfãs da 12ª eram da extração de 30/07 sob a de 03/08.
    mantidos, descartados = conciliacao.deduplicar([
        titulo(62801, date(2026, 7, 30)),
        titulo(2, date(2026, 8, 3)),
    ])
    verificar(len(mantidos) == 1 and mantidos[0].linha == 2,
              "fica a linha da DataGeracao mais recente")
    verificar(len(descartados) == 1 and descartados[0][1].linha == 62801,
              "a linha antiga é registrada como descartada")


def test_janela():
    print("\nJanela — vem do histórico, não do calendário (item 2)")

    #  A 14ª rodada real: janela 01/08 a 14/08, recompra em 17/08, emendando com
    #  a rodada anterior, que terminou em 31/07.
    verificar(calculo.sugerir_janela(date(2026, 8, 17), date(2026, 7, 31))
              == (date(2026, 8, 1), date(2026, 8, 14)),
              "reproduz a janela da 14ª: emenda em 01/08 e fecha 3 dias antes")
    verificar(calculo.inicio_seguinte(date(2026, 7, 31)) == date(2026, 8, 1),
              "o início é o dia seguinte ao fim da rodada anterior")
    verificar(calculo.inicio_seguinte(None) is None,
              "sem histórico não há início a sugerir — fica em branco")
    verificar(calculo.sugerir_janela(date(2026, 8, 14), None)
              == (None, date(2026, 8, 11)),
              "sem histórico só o fim é sugerido")

    #  A regra antiga de calendário produziria 16/07–31/07 para uma recompra em
    #  14/08 — a janela já consumida que zerou o Termo.
    verificar(calculo.sugerir_janela(date(2026, 8, 14), date(2026, 7, 31))[0]
              == date(2026, 8, 1),
              "para a recompra de 14/08 a janela começa em 01/08, não em 16/07")

    #  Coerência, não dia do mês: as janelas que as rodadas reais usaram passam.
    for inicio, fim, rotulo in ((date(2026, 7, 1), date(2026, 7, 15), "01–15/07, 11ª"),
                                (date(2026, 7, 1), date(2026, 7, 31), "01–31/07, 13ª"),
                                (date(2026, 8, 1), date(2026, 8, 14), "01–14/08, 14ª")):
        verificar(not calculo.problemas_da_janela(inicio, fim, date(2026, 8, 17)),
                  f"{rotulo} é janela coerente")
    verificar(calculo.problemas_da_janela(date(2026, 8, 20), date(2026, 8, 10)),
              "janela invertida é recusada")
    verificar(calculo.problemas_da_janela(date(2026, 8, 1), date(2026, 8, 20),
                                          date(2026, 8, 17)),
              "fim posterior à recompra é recusado")
    verificar(not calculo.janela_longa(date(2026, 7, 1), date(2026, 7, 31)),
              "31 dias está dentro do usual")
    verificar(calculo.janela_longa(date(2026, 1, 1), date(2026, 7, 31)),
              "sete meses de janela vira aviso")


def test_sobreposicao_de_janela():
    print("\nSobreposição de janela com rodada já emitida (item 2)")
    import tempfile
    from core.historico import Historico
    with tempfile.TemporaryDirectory() as pasta:
        hist = Historico(Path(pasta) / "h.json")
        hist.registrar(11, date(2026, 7, 24), date(2026, 7, 1), date(2026, 7, 15),
                       ["1"], 0)
        hist.registrar(13, date(2026, 8, 10), date(2026, 7, 1), date(2026, 7, 31),
                       ["2"], 0)
        hist.registrar(14, date(2026, 8, 17), date(2026, 8, 1), date(2026, 8, 14),
                       ["3"], 0)

        #  A janela da simulação que zerou o Termo: 16/07 a 31/07.
        achados = hist.sobreposicoes(date(2026, 7, 16), date(2026, 7, 31))
        verificar([r["numero"] for r in achados] == [13],
                  "a janela 16/07–31/07 é acusada de invadir a rodada 13",
                  str([r["numero"] for r in achados]))
        verificar(not hist.sobreposicoes(date(2026, 8, 15), date(2026, 8, 28)),
                  "a janela seguinte, 15/08–28/08, não invade nada")
        verificar(not hist.sobreposicoes(date(2026, 8, 1), date(2026, 8, 14),
                                         ignorar_numero=14),
                  "reemitir a 14ª não esbarra na 14ª que já está no histórico")
        verificar(hist.resumo_ultima()["numero"] == 14,
                  "a última rodada é a de janela mais recente, não a de maior número")
        verificar(not hist.numeracao_confiavel,
                  "a numeração 11, 13, 14 é acusada de ter buraco")
        verificar(hist.proximo_numero() == 15, "o próximo número é 15")


def test_reemissao_nao_se_trava():
    print("\nReemitir a mesma rodada não pode travar nos próprios títulos")
    import tempfile
    from core.historico import Historico
    with tempfile.TemporaryDirectory() as pasta:
        hist = Historico(Path(pasta) / "h.json")
        hist.registrar(11, date(2026, 7, 24), date(2026, 7, 1), date(2026, 7, 15),
                       ["165955407", "111"], 0)
        hist.registrar(14, date(2026, 8, 18), date(2026, 8, 1), date(2026, 8, 15),
                       ["222", "333"], 0)

        verificar(hist.ids_recomprados == {"165955407", "111", "222", "333"},
                  "sem número, a trava vê todos os títulos de todas as rodadas")
        #  Reemitindo a 14ª: os títulos DELA saem da trava…
        travados = hist.ids_recomprados_exceto(14)
        verificar(travados == {"165955407", "111"},
                  "reemitindo a 14ª, os títulos da própria 14ª não travam",
                  str(sorted(travados)))
        #  …mas os das outras rodadas continuam travados. É o caso da Cíntia:
        #  cobrada na 11ª, voltou na 13ª porque o boleto não foi baixado.
        verificar("165955407" in travados,
                  "o título da Cíntia, cobrado na 11ª, continua travado")
        verificar(hist.ids_recomprados_exceto(99) == hist.ids_recomprados,
                  "rodada nova não perde nada da trava")


def test_defasagem():
    print("\nDefasagem (seção 2.4)")
    #  Sexta 07/08 -> segunda 10/08: um dia útil. Foram 20 títulos pagos.
    verificar(calculo.dias_uteis(date(2026, 8, 7), date(2026, 8, 10)) == 1,
              "sexta para segunda é um dia útil")
    verificar(calculo.dias_uteis(date(2026, 8, 3), date(2026, 8, 7)) == 4,
              "03/08 a 07/08 são quatro dias úteis (saíram 74 títulos)")
    verificar(calculo.dias_uteis(date(2026, 8, 10), date(2026, 8, 10)) == 0,
              "extração no dia da recompra é defasagem zero")
    previsto = calculo.custo_defasagem(1)
    verificar(previsto["titulos_min"] <= 20 <= previsto["titulos_max"],
              "o modelo prevê os 20 títulos observados em um dia útil")


def test_calculo():
    print("\nCálculo de multa e juros")
    conta = calculo.calcular(Decimal("2966.61"), date(2026, 7, 4), date(2026, 8, 5),
                             Decimal("0.00033173"), Decimal("0.02"))
    verificar(conta.prazo == 32, "prazo em dias corridos, sem contar o vencimento")
    #  Valor conferido contra a linha 6 do Termo da 12ª: 3.057,596295726466
    verificar(abs(conta.total - Decimal("3057.596295726466")) < Decimal("0.000001"),
              "o total bate com a planilha até a sexta casa",
              f"calculado {conta.total}")


# ---------------------------------------------------------------------------
#  Aceite da Fase 1 — sobre os arquivos reais
# ---------------------------------------------------------------------------
def _rodar(arquivo, numero, data_recompra, inicio, fim, log=None):
    cfg = {
        "template": config.TEMPLATE_PADRAO,
        "numero_rodada": numero,
        "data_recompra": data_recompra,
        "janela_inicio": inicio,
        "janela_fim": fim,
        "fonte_vortx": "arquivo", "arquivo_vortx": str(arquivo),
        "fonte_grafeno": "arquivo", "arquivo_grafeno": str(arquivo),
        "data_extracao_grafeno": data_recompra,
        "modo_simulacao": True,
        "gravar_snapshot": False,
        "historico": config.RAIZ / "testes" / "historico_de_teste.json",
    }
    return motor.conciliar(cfg, log=log or (lambda *a, **k: None))


def _termo_emitido(arquivo):
    """Lê o Termo que foi de fato emitido naquele arquivo."""
    tabela = leitura.ler_aba(arquivo, config.ABA_TERMO)
    leitura.fechar_cache()
    emitidos = []
    for numero, linha in tabela.linhas:
        if numero < config.TERMO_PRIMEIRA_LINHA or len(linha) < 4:
            continue
        identificador = linha[3]
        if identificador in (None, "", "."):
            continue
        emitidos.append((str(identificador).strip(), str(linha[1] or "").strip()))
    return emitidos


def test_aceite_11a():
    print("\nAceite da Fase 1 — 11ª recompra (janela 01–15/07, recompra 24/07)")
    if not AMOSTRA_11.exists():
        print(f"  pulado: {AMOSTRA_11.name} não está em testes/amostras/")
        return

    resultado = _rodar(AMOSTRA_11, 11, date(2026, 7, 24),
                       date(2026, 7, 1), date(2026, 7, 15))
    emitidos = {i for i, _ in _termo_emitido(AMOSTRA_11)}
    calculados = {l["id_titulo"] for l in resultado.linhas_termo}
    faltantes = calculados - emitidos

    esperados = {
        "165959502": "MARIA IZABEL QUEIROZ DA SILVA BASILIO",
        "165960217": "ALINE DE OLIVEIRA GOMES",
        "165961375": "ANA TERESA VALLS PEREIRA",
        "143950451": "IRISDALVA TELES DE DEUS",
        "143950045": "ANA MARIA SILVA DAMASCENO",
    }
    verificar(faltantes == set(esperados),
              "encontra exatamente os 5 títulos que ficaram de fora do Termo",
              f"encontrados: {sorted(faltantes)}")
    verificar(not (emitidos - calculados),
              "nenhum título do Termo emitido é descartado pelo motor")

    por_id = {l["id_titulo"]: l for l in resultado.linhas_termo}
    nominal = sum((por_id[i]["valor_nominal"] for i in faltantes), Decimal(0))
    verificar(nominal == Decimal("10367.84"),
              "os 5 somam R$ 10.367,84 de valor nominal", f"soma {nominal}")

    #  Dois deles têm dois boletos (o duplo status) e três têm um só (o limite
    #  de linha da fórmula): 7 linhas de boleto para 5 títulos — a seção 2.6.
    linhas_boleto = sum(
        next(c for c in resultado.conciliacao.casamentos
             if c.titulo.id_titulo == i).qtd_boletos for i in faltantes)
    verificar(linhas_boleto == 7,
              "os 5 títulos correspondem a 7 linhas de boleto na Grafeno",
              f"contadas {linhas_boleto}")


def test_aceite_12a():
    print("\nAceite da Fase 1 — 12ª recompra v1.1 (janela 01–31/07, recompra 05/08)")
    if not AMOSTRA_12.exists():
        print(f"  pulado: {AMOSTRA_12.name} não está em testes/amostras/")
        return

    resultado = _rodar(AMOSTRA_12, 12, date(2026, 8, 5),
                       date(2026, 7, 1), date(2026, 7, 31))

    orfas = resultado.conciliacao.contagens["linhas_removidas_dedup"]
    verificar(orfas == 7,
              "identifica as 7 linhas órfãs da extração de 30/07",
              f"removidas {orfas}")
    verificar(len(resultado.conciliacao.ids_duplicados) == 7,
              "as 7 linhas pertencem a 7 IdTituloVortx repetidos")
    verificar(len(resultado.vortx.datas_geracao) == 2,
              "detecta as duas DataGeracao misturadas na aba VORTX")

    emitidos = [i for i, _ in _termo_emitido(AMOSTRA_12)]
    repetidos = {i for i in emitidos if emitidos.count(i) > 1}
    verificar(len(repetidos) == 5,
              "identifica os 5 títulos cobrados em duplicidade no Termo",
              f"repetidos: {sorted(repetidos)}")
    verificar(len(emitidos) == 758 and len(set(emitidos)) == 753,
              "o Termo emitido tem 758 linhas para 753 títulos")

    calculados = {l["id_titulo"] for l in resultado.linhas_termo}
    verificar(len(calculados) == len(resultado.linhas_termo),
              "o Termo do motor não tem linha repetida")

    #  Os cinco duplicados ficam barrados até a extração ser refeita.
    barrados = {c.titulo.id_titulo for c, _ in resultado.selecao.barrados}
    verificar(repetidos <= barrados | calculados,
              "os títulos duplicados aparecem barrados ou já corrigidos")

    validacao = next(v for v in resultado.validacoes if v.chave == "extracao_unica")
    verificar(validacao.nivel != "OK", "a validação de extração única acusa o problema")


# ---------------------------------------------------------------------------
#  Item 4 — o Termo vazio precisa levar a fórmula dentro
# ---------------------------------------------------------------------------
def test_termo_vazio_preserva_formulas():
    print("\nTermo vazio ainda leva as fórmulas FILTER (item 4)")
    import re
    import tempfile
    import zipfile
    from core import excel_saida

    if not config.TEMPLATE_PADRAO.exists():
        print("  pulado: template não está em templates/")
        return

    class Fingido:
        """O mínimo que ``excel_saida.gerar`` precisa, sem ler base nenhuma."""
        def __init__(self):
            self.template = config.TEMPLATE_PADRAO
            self.linhas_termo = []
            self.total_nominal = Decimal(0)
            self.total_recompra = Decimal(0)
            self.modo_simulacao = True
            self.parametros = calculo.Parametros(
                numero_rodada=99, data_recompra=date(2026, 8, 17),
                janela_inicio=date(2026, 8, 1), janela_fim=date(2026, 8, 14),
                juros_mora=Decimal("0.00033173"), multa=Decimal("0.02"))
            self.conciliacao = type("C", (), {"casamentos": []})()
            self.grafeno = type("G", (), {"registros": []})()

    with tempfile.TemporaryDirectory() as pasta:
        caminho = excel_saida.gerar(Fingido(), pasta)
        with zipfile.ZipFile(caminho) as z:
            alvo = next(n for n in z.namelist()
                        if re.search(r"xl/worksheets/sheet\d+\.xml$", n)
                        and b"FILTER" in z.read(n))
            xml = z.read(alvo).decode("utf-8")

    linha = config.TERMO_PRIMEIRA_LINHA
    for letra in ("B", "C", "D", "E", "F", "G", "H"):
        achado = re.search(rf'<c r="{letra}{linha}".*?</c>', xml, re.S)
        verificar(achado is not None and "FILTER" in achado.group(0),
                  f"{letra}{linha} leva a fórmula FILTER mesmo com Termo vazio",
                  achado.group(0)[:120] if achado else "célula ausente")
    verificar(f'ref="B{linha}:B{linha}"' in xml,
              "a fórmula de matriz fica ancorada na própria linha 5")
    #  Era este o falso positivo: sem valor e sem fórmula, ninguém percebia.
    verificar("<v>" not in re.search(rf'<c r="B{linha}".*?</c>', xml, re.S).group(0),
              "e sem valor em cache, porque não há título nenhum")


def test_totais_em_cache():
    """
    Os totalizadores do template guardam o número **desta** rodada.

    ``TERMO!C2/G2/H2`` e ``INFORMAÇÕES!D12/C19`` são fórmulas. A fórmula é
    preservada, mas o valor em cache vinha do template: a rodada 14, com 2.638
    títulos e R$ 3,71 MM, saía anunciando 680 títulos e R$ 894.362,31 para quem
    lê o arquivo sem abrir no Excel.
    """
    print("\nTotalizadores levam o número da própria rodada, não o do template")
    import re
    import tempfile
    from core import excel_saida

    if not config.TEMPLATE_PADRAO.exists():
        print("  pulado: template não está em templates/")
        return

    class Fingido:
        def __init__(self):
            self.template = config.TEMPLATE_PADRAO
            self.total_nominal = Decimal("3630971.40")
            self.total_recompra = Decimal("3712896.72")
            self.modo_simulacao = True
            self.parametros = calculo.Parametros(
                numero_rodada=14, data_recompra=date(2026, 8, 18),
                janela_inicio=date(2026, 8, 1), janela_fim=date(2026, 8, 15),
                juros_mora=Decimal("0.00033173"), multa=Decimal("0.02"))
            self.conciliacao = type("C", (), {"casamentos": []})()
            self.grafeno = type("G", (), {"registros": []})()
            self.linhas_termo = [
                {"nome": f"SACADO {n}", "documento_formatado": "000.000.000-00",
                 "id_titulo": str(n), "vencimento": date(2026, 8, 3),
                 "data_recompra": date(2026, 8, 18),
                 "valor_nominal": Decimal("100.00"),
                 "valor_recompra": Decimal("102.00")}
                for n in range(2638)]

    with tempfile.TemporaryDirectory() as pasta:
        caminho = excel_saida.gerar(Fingido(), pasta)
        pacote = excel_saida.Pacote(caminho)
        termo = pacote.aba(config.ABA_TERMO).corpo
        info = pacote.aba(config.ABA_INFORMACOES).corpo

    def cache(corpo, referencia):
        achado = re.search(rf'<c r="{referencia}"[^>]*>(.*?)</c>', corpo, re.S)
        if not achado:
            return None, None
        formula = re.search(r"<f\b.*?</f>", achado.group(1), re.S)
        valor = re.search(r"<v>(.*?)</v>", achado.group(1), re.S)
        return (formula.group(0) if formula else None,
                valor.group(1) if valor else None)

    for referencia, esperado, rotulo in (
            ("C2", "2638", "quantidade de títulos"),
            ("G2", "3630971.4", "total nominal"),
            ("H2", "3712896.72", "total de recompra")):
        formula, valor = cache(termo, referencia)
        verificar(formula is not None,
                  f"TERMO!{referencia} continua sendo fórmula")
        verificar(valor is not None and abs(float(valor) - float(esperado)) < 0.01,
                  f"TERMO!{referencia} tem o {rotulo} desta rodada",
                  f"cache = {valor}, esperado {esperado}")

    for referencia, esperado, rotulo in (
            (config.CEL_QTD_RECOMPRA, "2638", "quantidade"),
            (config.CEL_VALOR_TOTAL, "3712896.72", "valor total")):
        formula, valor = cache(info, referencia)
        verificar(formula is not None,
                  f"INFORMAÇÕES!{referencia} continua sendo fórmula")
        verificar(valor is not None and abs(float(valor) - float(esperado)) < 0.01,
                  f"INFORMAÇÕES!{referencia} tem o {rotulo} desta rodada",
                  f"cache = {valor}, esperado {esperado}")
    #  Era o número do template que aparecia ali.
    verificar("680" not in (cache(info, config.CEL_QTD_RECOMPRA)[1] or ""),
              "o 680 do template não sobrou em INFORMAÇÕES!D12")


# ---------------------------------------------------------------------------
#  Aceite do item 1 — a chave 403 não substitui a chave composta
# ---------------------------------------------------------------------------
def test_ligacao_403_nao_e_chave_primaria():
    print("\nA ligação '403'+NumeroTitulo só desempata; não cria par (item 1)")

    def boleto(nosso, doc, status="Aberta (Vencida)", linha=2):
        return Boleto(documento=doc, vencimento=date(2026, 8, 1),
                      valor=Decimal("100.00"), status=status, linha=linha,
                      nosso_numero=nosso, chave_boleto=nz.so_digitos(nosso))

    def titulo(numero, doc):
        return Titulo(id_titulo="1", numero_titulo=numero, documento=doc,
                      vencimento=date(2026, 8, 1), valor_nominal=Decimal("100.00"),
                      linha=2, chave_boleto=nz.chave_boleto(numero),
                      chave=nz.chave(doc, date(2026, 8, 1), Decimal("100.00")),
                      chave_parcial=nz.chave_parcial(doc, date(2026, 8, 1)))

    #  O caso medido 2.451 vezes na base da 14ª: o "403"+NumeroTitulo existe na
    #  Grafeno, com o mesmo vencimento e o mesmo valor, e é de OUTRO CPF.
    alheio = boleto("403382661122", "41999347749")
    meu = titulo("382661122", "69261881734")
    resultado = conciliacao.conciliar([meu], [alheio])
    verificar(not resultado.casados,
              "o boleto de outro CPF não é casado, mesmo com a ligação batendo")
    verificar(resultado.casamentos[0].causa is not None,
              "o título fica classificado como exceção, não emparelhado")

    #  O uso legítimo: dois boletos que a chave composta já validou, e a ligação
    #  diz qual é o do título.
    do_titulo = boleto("403382661122", "69261881734", "Aberta (Vencida)", 10)
    vizinho = boleto("403382661130", "69261881734", "Baixada", 11)
    escolhido = conciliacao.escolher_boleto([vizinho, do_titulo], meu)
    verificar(escolhido is do_titulo,
              "entre boletos da mesma chave, a ligação aponta o do título")


# ---------------------------------------------------------------------------
#  Teste de aceite pedido: rodada 14, recompra 17/08, janela 01/08 a 14/08
# ---------------------------------------------------------------------------
def _arquivo_da_14a():
    pasta = config.PASTA_ONEDRIVE / "14.Recompra_17_08_2026"
    if not pasta.exists():
        return None
    from core.retroativo import _escolher_arquivo
    return _escolher_arquivo(pasta)


def test_aceite_14a():
    """
    A 14ª v1.1 (janela 01–15/08, recompra 18/08), com o **histórico real**.

    A rodada 14 já está registrada no histórico, e isso é parte do teste: é a
    reemissão que a trava contra recompra repetida não pode impedir. Sem a
    exclusão da própria rodada, os 2.639 títulos dela voltariam como "já
    recomprados", todos seriam barrados e o Termo iria a zero.

    O alvo **não** é 2.639, é 2.638. O Termo manual trouxe um título a mais:
    OLGA ALVES RODRIGUES tem dois títulos vencendo em 11/08 e o de R$ 1.472,26
    tem boleto ``Baixada`` — a cobrança foi cancelada e reemitida como a de
    R$ 1.608,78, que é a que está vencida em aberto e entra no Termo pela linha
    do irmão. O manual cobrou os dois: R$ 1.505,13 em duplicidade.
    """
    print("\nAceite — 14ª recompra v1.1 (janela 01–15/08, recompra 18/08)")
    arquivo = _arquivo_da_14a()
    if arquivo is None:
        print("  pulado: a pasta 14.Recompra_17_08_2026 não está sincronizada")
        return
    if not config.ARQUIVO_HISTORICO.exists():
        print("  pulado: histórico não foi populado ainda "
              "(core/semente_historico.py)")
        return
    print(f"  base: {arquivo.name}")

    resultado = motor.conciliar({
        "template": config.TEMPLATE_PADRAO,
        "numero_rodada": 14,
        "data_recompra": date(2026, 8, 18),
        "janela_inicio": date(2026, 8, 1),
        "janela_fim": date(2026, 8, 15),
        "fonte_vortx": "arquivo", "arquivo_vortx": str(arquivo),
        "fonte_grafeno": "arquivo", "arquivo_grafeno": str(arquivo),
        "data_extracao_grafeno": date(2026, 8, 18),
        "modo_simulacao": True,
        "gravar_snapshot": False,
        "historico": None,          # o histórico real, com a 14ª dentro
    })

    quantidade = len(resultado.linhas_termo)
    total = resultado.total_recompra
    emitidos = {i for i, _ in _termo_emitido(arquivo)}
    calculados = {l["id_titulo"] for l in resultado.linhas_termo}

    print(f"  → motor  {quantidade} títulos · R$ {total:.2f}")
    print(f"    manual {len(emitidos)} títulos · R$ 3715609.43")

    verificar(quantidade == 2638,
              "apura 2.638 títulos — o manual menos o título duplicado",
              f"obtidos {quantidade} (diferença de {quantidade - 2638})")
    verificar(not (calculados - emitidos),
              "não traz nenhum título que o Termo manual não tenha",
              f"a mais: {sorted(calculados - emitidos)[:5]}")
    verificar(emitidos - calculados == {"165960004"},
              "a única diferença é o 165960004 (OLGA), de boleto baixado",
              f"diferença: {sorted(emitidos - calculados)}")
    #  R$ 3.715.609,43 do manual menos os R$ 1.505,13 da linha em duplicidade.
    verificar(abs(total - Decimal("3714104.31")) < Decimal("0.02"),
              "e o valor bate com o manual menos os R$ 1.505,13 cobrados a mais",
              f"obtido {total}")

    niveis = {v.chave: v.nivel for v in resultado.validacoes}
    verificar(niveis.get("sem_recompra_repetida") == "OK",
              "reemitir a 14ª não trava nos títulos da própria 14ª",
              str(niveis.get("sem_recompra_repetida")))
    verificar(not resultado.bloqueado,
              "nada bloqueia a reemissão",
              str([v.titulo for v in resultado.erros]))
    for chave in ("termo_nao_vazio", "janela_sem_sobreposicao", "janela_coerente",
                  "convergencia", "grafeno_frescor"):
        verificar(niveis.get(chave) == "OK", f"validação {chave} em OK",
                  str(niveis.get(chave)))
    verificar(resultado.grafeno_data_conteudo == date(2026, 8, 18),
              "a data da Grafeno sai do conteúdo (18/08), não do arquivo",
              str(resultado.grafeno_data_conteudo))
    #  O caso da Olga tem de aparecer, e como resolvido: a dívida dela já está
    #  no Termo pela linha do irmão, então não é aviso, é OK explicado.
    baixado = next(v for v in resultado.validacoes
                   if v.chave == "baixado_com_irmao")
    verificar("165960004" in baixado.dados.get("cobertos", []),
              "o caso da Olga é identificado como já coberto pelo irmão",
              baixado.detalhe[:120])
    print(f"    defasagem medida: {resultado.defasagem_dias} dia(s) útil(eis) — "
          f"Vórtx {nz.br(resultado.data_extracao_vortx)}, Grafeno "
          f"{nz.br(resultado.data_extracao_grafeno)}")


# ---------------------------------------------------------------------------
def main():
    print("=" * 72)
    print("  Testes do motor de recompra — FIDC Leve Saúde")
    print("=" * 72)

    test_normalizacao()
    test_prioridade_status()
    test_desempate_por_numero_titulo()
    test_ligacao_403_nao_e_chave_primaria()
    test_deduplicacao()
    test_janela()
    test_sobreposicao_de_janela()
    test_reemissao_nao_se_trava()
    test_defasagem()
    test_calculo()
    test_termo_vazio_preserva_formulas()
    test_totais_em_cache()

    if "--rapido" not in sys.argv:
        test_aceite_11a()
        test_aceite_12a()
        test_aceite_14a()
    else:
        print("\n(aceite sobre os arquivos reais pulado: --rapido)")

    print("\n" + "=" * 72)
    if _falhas:
        print(f"  {len(_falhas)} FALHA(S):")
        for f in _falhas:
            print(f"    · {f}")
        return 1
    print("  Tudo certo.")
    return 0


if __name__ == "__main__":
    sys.exit(main())