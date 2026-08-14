# -*- coding: utf-8 -*-
"""
Aceite da Fase 2 — a interface inteira, por HTTP, sem navegador.
================================================================

Roda com::

    py testes/test_interface.py

Sobe o mesmo servidor que o ``recompra_app.py`` sobe no duplo clique e exercita
o fluxo de ponta a ponta sobre o arquivo real da 12ª recompra: conciliar, ver o
painel, baixar o relatório de exceções e gerar a planilha.

O que ele prova, além de "as rotas respondem":

* **o fluxo é mesmo em dois tempos** — a conciliação não escreve planilha
  nenhuma, e o relatório de exceções já está disponível antes da emissão, que é
  o momento de aprovação da decisão 2.2;
* **defasagem não medida não vira defasagem zero** — a 12ª v1.1 tem duas
  ``DataGeracao``, então a data de extração da Vórtx é desconhecida e o resumo
  precisa dizer isso, não mostrar verde;
* **o download só entrega o que aquela rodada gerou** — caminho vindo da URL
  nunca é aberto;
* **o modo simulação não toca no histórico de recompras.**

Leva alguns minutos: são dois arquivos de 50 MB lidos e uma planilha de 38 MB
gravada. Só roda se a amostra estiver em ``testes/amostras/``.
"""

import json
import shutil
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import recompra_app as app                                       # noqa: E402
from core import config                                          # noqa: E402

AMOSTRA = config.PASTA_AMOSTRAS / "Leve_Saude_12_Recompra_0,59MM_v1.1_08-2026.xlsx"

_falhas = []


def verificar(condicao, descricao, detalhe=""):
    if condicao:
        print(f"  ok   {descricao}")
    else:
        print(f"  FALHA {descricao}" + (f"\n         {detalhe}" if detalhe else ""))
        _falhas.append(descricao)


# ---------------------------------------------------------------------------
#  Cliente HTTP mínimo
# ---------------------------------------------------------------------------
class Cliente:
    """O que um navegador faria, sem navegador."""

    def __init__(self, base):
        self.base = base

    def get(self, rota, timeout=1800):
        """(bytes, status, headers) — o download de .xlsx é binário."""
        with urllib.request.urlopen(self.base + rota, timeout=timeout) as r:
            return r.read(), r.status, r.headers

    def json(self, rota):
        corpo, _, _ = self.get(rota)
        return json.loads(corpo.decode("utf-8"))

    def sse(self, rota):
        """Consome um stream de eventos: (linhas_de_log, payload_final, evento)."""
        linhas, final, nome = [], None, None
        with urllib.request.urlopen(self.base + rota, timeout=3600) as r:
            evento = None
            for bruta in r:
                texto = bruta.decode("utf-8").rstrip("\n")
                if texto.startswith("event: "):
                    evento = texto[7:]
                elif texto.startswith("data: "):
                    dado = json.loads(texto[6:])
                    if evento:
                        final, nome, evento = dado, evento, None
                    else:
                        linhas.append(dado)
        return linhas, final, nome


def subir_servidor():
    porta = app.achar_porta()
    servidor = ThreadingHTTPServer(("127.0.0.1", porta), app.Handler)
    threading.Thread(target=servidor.serve_forever, daemon=True).start()
    return servidor, Cliente(f"http://127.0.0.1:{porta}")


# ---------------------------------------------------------------------------
#  Rotas de apoio — o que a tela consulta ao abrir
# ---------------------------------------------------------------------------
def test_pagina_e_apoio(cliente):
    print("\nPágina e rotas de apoio")
    pagina, codigo, _ = cliente.get("/")
    pagina = pagina.decode("utf-8")
    verificar(codigo == 200 and "Recompra FIDC" in pagina, "GET / devolve a página")
    verificar('id="btnConciliar"' in pagina and 'id="btnGerar"' in pagina,
              "a página tem os dois botões do fluxo em dois tempos")

    inicial = cliente.json("/inicial")
    verificar(inicial["template_existe"], "GET /inicial encontra o template")
    verificar(inicial.get("juros_mora") == "0.033173" and inicial.get("multa") == "2",
              "juros e multa vêm da aba INFORMAÇÕES, não do código",
              str(inicial.get("juros_mora")))

    janela = cliente.json("/janela?data=2026-08-05")
    verificar(janela == {"inicio": "2026-07-16", "fim": "2026-07-31"},
              "GET /janela aplica a regra quinzenal", str(janela))

    grafeno = cliente.json("/testar-grafeno")
    verificar(not grafeno["ok"] and "Arquivo" in grafeno["mensagem"],
              "o teste da API Grafeno responde a pendência, sem inventar endpoint")


def test_formulario_incompleto(cliente):
    print("\nErros de preenchimento são recusados antes de qualquer leitura")
    _, final, nome = cliente.sse("/conciliar?" + urllib.parse.urlencode(
        {"data_recompra": "2026-08-05", "fonte_vortx": "arquivo", "arquivo_vortx": ""}))
    verificar(nome == "erro" and "Vórtx" in final["msg"],
              "conciliar sem o arquivo da Vórtx devolve erro claro", str(final))


# ---------------------------------------------------------------------------
#  Fluxo em dois tempos, sobre a 12ª
# ---------------------------------------------------------------------------
def parametros_da_12a(pasta_saida):
    return {
        "numero_rodada": "12", "data_recompra": "2026-08-05",
        "janela_inicio": "2026-07-01", "janela_fim": "2026-07-31",
        "fonte_vortx": "arquivo", "arquivo_vortx": str(AMOSTRA), "aba_vortx": "VORTX",
        "fonte_grafeno": "arquivo", "arquivo_grafeno": str(AMOSTRA),
        "aba_grafeno": "GRAFENO", "data_extracao_grafeno": "2026-08-05",
        "pasta_saida": str(pasta_saida),
        "copiar_onedrive": "0", "atualizar_historico": "1", "modo_simulacao": "1",
    }


def test_conciliar(cliente, parametros):
    print("\n1º tempo — Conciliar a 12ª (janela 01–31/07, recompra 05/08)")
    comeco = time.time()
    linhas, r, nome = cliente.sse("/conciliar?" + urllib.parse.urlencode(parametros))
    print(f"  ({time.time() - comeco:.0f}s, {len(linhas)} linhas de log)")
    verificar(nome == "fim", "a conciliação termina com evento 'fim'", str(r)[:400])
    if nome != "fim":
        return None

    resumo = r["resumo"]
    print(f"  → {resumo['titulos_elegiveis']} títulos · nominal "
          f"R$ {resumo['valor_nominal']} · recompra R$ {resumo['valor_recompra']} "
          f"· {resumo['excecoes']} exceções")

    verificar(bool(r["token"]), "devolve um token de sessão")
    verificar(len(r["previa"]) == 12, "a prévia traz as primeiras 12 linhas")
    verificar(r["total_linhas"] == resumo["titulos_elegiveis"],
              "a prévia informa o total de linhas do Termo")
    verificar(all(c in r["previa"][0] for c in
                  ("nome", "documento", "id_titulo", "vencimento", "valor_recompra")),
              "cada linha traz sacado, documento, título, vencimento e valor")

    niveis = {v["chave"]: v["nivel"] for v in r["validacoes"]}
    verificar(len(niveis) >= 14, f"a bateria roda inteira ({len(niveis)} validações)")
    verificar(all(v["porque"] for v in r["validacoes"]),
              "toda validação explica o defeito que está protegendo")
    verificar(niveis.get("extracao_unica") == "AVISO",
              "extração única acusa as duas DataGeracao da 12ª", str(niveis))

    #  Sem a data de extração da Vórtx não há como medir a defasagem — e
    #  defasagem não medida não é defasagem zero (seção 2.4).
    verificar(not resumo["defasagem_medida"]
              and "desconhecida" in resumo["defasagem_frase"]
              and "defasagem zero" not in resumo["defasagem_frase"],
              "defasagem não medida não é anunciada como defasagem zero",
              resumo["defasagem_frase"])
    verificar(niveis.get("defasagem_zero") == "AVISO",
              "e a validação de defasagem acusa o mesmo")

    verificar(r["modo_simulacao"] and "_SIMULACAO" in r["arquivo_sugerido"],
              "o nome sugerido marca a simulação", r["arquivo_sugerido"])
    verificar(not r["bloqueado"], "em simulação nada bloqueia — o gerar libera")
    verificar(Path(r["snapshot"]).exists(),
              "o snapshot foi gravado antes do processamento (regra 15.7)",
              r["snapshot"])
    return r


def test_excecoes(cliente, token):
    print("\nRelatório de exceções, disponível já depois de conciliar (decisão 2.2)")
    excecao = cliente.json(f"/excecoes?token={token}")
    verificar(Path(excecao["caminho"]).exists(), "o arquivo de exceções foi gravado",
              str(excecao))

    corpo, codigo, cabecalhos = cliente.get(f"/baixar?token={token}&nome=excecoes")
    verificar(codigo == 200
              and "attachment" in cabecalhos.get("Content-Disposition", ""),
              "o download entrega o arquivo com nome")
    verificar(len(corpo) > 5000, f"o arquivo tem conteúdo ({len(corpo)} bytes)")

    #  A planilha da rodada ainda não existe: o download não pode inventá-la, e
    #  caminho vindo da URL nunca é aberto.
    try:
        cliente.get(f"/baixar?token={token}&nome=excel")
        verificar(False, "baixar o que ainda não foi gerado é recusado")
    except urllib.error.HTTPError as erro:
        corpo = json.loads(erro.read().decode("utf-8"))
        verificar(erro.code == 404 and "erro" in corpo,
                  "baixar o que ainda não foi gerado é recusado com 404", str(corpo))


def test_gerar(cliente, parametros, token):
    print("\n2º tempo — Gerar")
    comeco = time.time()
    _linhas, r, nome = cliente.sse("/gerar?" + urllib.parse.urlencode(
        dict(parametros, token=token)))
    print(f"  ({time.time() - comeco:.0f}s)")
    verificar(nome == "fim", "a geração termina com evento 'fim'", str(r)[:400])
    if nome != "fim":
        return

    arquivos = {a["nome"]: a for a in r["arquivos"]}
    verificar("excel" in arquivos and "excecoes" in arquivos,
              "gera a planilha da rodada e o relatório de exceções",
              str(list(arquivos)))
    for dados in arquivos.values():
        caminho = Path(dados["caminho"])
        verificar(caminho.exists(),
                  f"{dados['rotulo']}: {caminho.name}"
                  + (f" ({caminho.stat().st_size / 1024 / 1024:.1f} MB)"
                     if caminho.exists() else " — não foi gravado"))
    verificar("onedrive" not in arquivos,
              "sem a caixa marcada, nada foi copiado para o OneDrive")

    corpo, codigo, _ = cliente.get(f"/baixar?token={token}&nome=excel")
    verificar(codigo == 200 and len(corpo) > 100_000,
              f"a planilha pode ser baixada pela tela ({len(corpo)} bytes)")


def test_historico_intacto():
    caminho = config.ARQUIVO_HISTORICO
    rodadas = (json.loads(caminho.read_text(encoding="utf-8")).get("rodadas", [])
               if caminho.exists() else [])
    verificar(not any(r.get("numero") == 12 for r in rodadas),
              "modo simulação não gravou a rodada no histórico de recompras")


def test_retroativo(cliente, pasta_saida):
    """
    A rota do reprocessamento, sobre uma rodada barata.

    A rodada 1 não tem aba de Termo, então o caminho inteiro roda em segundos —
    varredura, classificação, planilha consolidada e download — sem ler as bases
    de 50 MB das rodadas reconciliáveis.
    """
    print("\nReprocessamento retroativo (rodada 1, a mais barata)")
    from core import config as cfg
    if not cfg.PASTA_ONEDRIVE.exists():
        print("  pulado: pasta das rodadas anteriores não está sincronizada")
        return

    _linhas, r, nome = cliente.sse("/retroativo?" + urllib.parse.urlencode(
        {"apenas": "1", "pasta_saida": str(pasta_saida)}))
    verificar(nome == "fim", "a rota do retroativo termina com evento 'fim'",
              str(r)[:300])
    if nome != "fim":
        return
    verificar(r["rodadas"] == 1, "o filtro 'apenas' limita as rodadas",
              str(r.get("rodadas")))
    verificar(Path(r["caminho"]).exists(),
              f"grava a planilha consolidada ({r['arquivo']})")
    verificar(r["detalhe"][0]["situacao"] == "sem estrutura"
              and "Termo" in r["detalhe"][0]["motivo"],
              "classifica a rodada sem estrutura e diz por quê",
              str(r["detalhe"][0]))
    corpo, codigo, _ = cliente.get(f"/baixar?token={r['token']}&nome=retroativo")
    verificar(codigo == 200 and len(corpo) > 3000,
              f"o consolidado pode ser baixado pela tela ({len(corpo)} bytes)")


def test_sessao(cliente, parametros):
    print("\nSessão")
    _linhas, final, nome = cliente.sse("/gerar?" + urllib.parse.urlencode(
        dict(parametros, token="inexistente")))
    verificar(nome == "erro" and "Conciliar" in final["msg"],
              "token desconhecido manda conciliar de novo", str(final))


# ---------------------------------------------------------------------------
def main():
    print("=" * 72)
    print("  Aceite da Fase 2 — interface da recompra")
    print("=" * 72)

    servidor, cliente = subir_servidor()
    print(f"servidor em {cliente.base}")

    #  As saídas vão para uma pasta temporária: o teste não polui saidas/.
    pasta_saida = Path(tempfile.mkdtemp(prefix="recompra_teste_"))
    try:
        test_pagina_e_apoio(cliente)
        test_formulario_incompleto(cliente)

        if not AMOSTRA.exists():
            print(f"\n  pulado: {AMOSTRA.name} não está em testes/amostras/")
        else:
            parametros = parametros_da_12a(pasta_saida)
            resultado = test_conciliar(cliente, parametros)
            if resultado:
                test_excecoes(cliente, resultado["token"])
                test_gerar(cliente, parametros, resultado["token"])
                test_historico_intacto()
            test_sessao(cliente, parametros)
        test_retroativo(cliente, pasta_saida)
    finally:
        servidor.shutdown()
        shutil.rmtree(pasta_saida, ignore_errors=True)

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
