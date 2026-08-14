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
    print("\nJanela quinzenal (decisão 1.4)")
    verificar(calculo.sugerir_janela(date(2026, 8, 5))
              == (date(2026, 7, 16), date(2026, 7, 31)),
              "recompra na 1ª quinzena -> 16 ao fim do mês anterior")
    verificar(calculo.sugerir_janela(date(2026, 7, 24))
              == (date(2026, 7, 1), date(2026, 7, 15)),
              "recompra na 2ª quinzena -> 1 a 15 do mês corrente")
    verificar(calculo.janela_valida(date(2026, 7, 1), date(2026, 7, 15)),
              "01–15 é janela válida")
    verificar(not calculo.janela_valida(date(2026, 7, 1), date(2026, 7, 31)),
              "01–31 não é quinzena (foi a janela usada na 12ª)")
    verificar(calculo.janela_seguinte(date(2026, 7, 15))
              == (date(2026, 7, 16), date(2026, 7, 31)),
              "a janela seguinte emenda sem lacuna")


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
def main():
    print("=" * 72)
    print("  Testes do motor de recompra — FIDC Leve Saúde")
    print("=" * 72)

    test_normalizacao()
    test_prioridade_status()
    test_desempate_por_numero_titulo()
    test_deduplicacao()
    test_janela()
    test_defasagem()
    test_calculo()

    if "--rapido" not in sys.argv:
        test_aceite_11a()
        test_aceite_12a()
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