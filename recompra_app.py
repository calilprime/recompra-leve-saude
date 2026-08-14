# -*- coding: utf-8 -*-
"""
Recompra FIDC Leve Saúde — Interface local (Netz Asset)
=======================================================

App web local, no mesmo padrão do ``fnet_app.py``: sobe um servidor pequeno no
seu próprio computador e abre uma página no navegador. Nada sai da máquina além
das consultas de leitura ao banco da Vórtx.

Como usar
---------
    Duplo clique em  "Abrir Recompra.bat"
    ou, no terminal:  py recompra_app.py

O navegador abre sozinho em http://127.0.0.1:<porta>. Se não abrir, copie o
endereço mostrado no terminal.

O fluxo é em dois tempos, de propósito (seção 11.2 do plano):

**1. Conciliar** — extrai, concilia, calcula e valida. Não escreve nada de
definitivo. A tela mostra o painel de validações, o resumo e a prévia do Termo.
O relatório de exceções já pode ser baixado aqui, que é o momento de aprovação
previsto na decisão 2.2.

**2. Gerar Termo e planilha** — só aparece depois, e fica desabilitado se
qualquer validação estiver em ``ERRO`` (regra 15.5).

Requisitos: Python 3.8+, ``openpyxl`` e — só para a fonte "Banco de dados" —
``psycopg2-binary``. Instale com ``pip install -r requirements.txt``. O tkinter,
usado nos botões "Procurar…", já vem com o Python padrão no Windows.
"""

import json
import os
import secrets
import sys
import threading
import urllib.parse
import webbrowser
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

try:
    from core import calculo, config, excecoes, excel_saida, fonte_grafeno
    from core import fonte_vortx, historico as historico_mod, leitura, motor
    from core import normalizacao as nz
    from core import retroativo
except ImportError as erro:                                  # pragma: no cover
    print("ERRO ao carregar os módulos do motor:", erro)
    print("Confira se a pasta 'core' está ao lado deste arquivo e rode:")
    print("    pip install -r requirements.txt")
    sys.exit(1)


# ============================================================================
#  1. Estado da sessão — o que liga o "Conciliar" ao "Gerar"
# ============================================================================
#  A conciliação devolve um Resultado grande (as duas bases inteiras). Ele fica
#  na memória, por token, para que o segundo tempo não precise reler nada — e
#  para que o que é emitido seja exatamente o que foi validado na tela, não uma
#  segunda apuração que poderia divergir.
_SESSOES = {}
_LIMITE_SESSOES = 4
_TRAVA = threading.Lock()


class Sessao:
    """Uma conciliação já feita, à espera da decisão de emitir."""

    def __init__(self, cfg, resultado):
        self.token = secrets.token_hex(8)
        self.cfg = cfg
        self.resultado = resultado
        self.criada_em = datetime.now()
        self.arquivos = {}

    def registrar_arquivo(self, nome, caminho):
        self.arquivos[nome] = Path(caminho)
        return self.arquivos[nome]


def guardar(sessao):
    _SESSOES[sessao.token] = sessao
    #  Mantém só as últimas: cada uma segura ~140 mil boletos na memória.
    for antigo in sorted(_SESSOES.values(), key=lambda s: s.criada_em)[:-_LIMITE_SESSOES]:
        _SESSOES.pop(antigo.token, None)
    return sessao


def buscar(token):
    sessao = _SESSOES.get(token or "")
    if sessao is None:
        raise RuntimeError(
            "Esta conciliação não está mais na memória (o servidor foi "
            "reiniciado?). Clique em Conciliar de novo antes de gerar."
        )
    return sessao


# ============================================================================
#  2. Log ao vivo — SSE na tela e arquivo em logs/ (seção 10.3)
# ============================================================================
class Registrador:
    """Manda cada linha para a tela e para ``logs/execucao_AAAA-MM-DD.log``."""

    def __init__(self, emitir, etapa=""):
        self.emitir = emitir
        self.caminho = config.PASTA_LOGS / f"execucao_{date.today():%Y-%m-%d}.log"
        if etapa:
            self._gravar(f"\n===== {etapa} · {datetime.now():%d/%m/%Y %H:%M:%S} =====")

    def __call__(self, mensagem, nivel="info"):
        self.emitir(mensagem, nivel)
        self._gravar(f"[{nivel:>8}] {mensagem}" if mensagem else "")

    def _gravar(self, linha):
        try:
            with open(self.caminho, "a", encoding="utf-8") as arquivo:
                arquivo.write(linha + "\n")
        except OSError:
            pass          # log em arquivo é conveniência; nunca derruba a rodada


# ============================================================================
#  3. Consultas rápidas que a tela faz ao abrir e ao mexer nos campos
# ============================================================================
def estado_inicial():
    """Preenche a tela: rodada sugerida, janela, taxas do template, fontes."""
    hoje = date.today()
    template = config.TEMPLATE_PADRAO
    dados = {
        "hoje": hoje.isoformat(),
        "template": str(template),
        "template_existe": template.exists(),
        "pasta_saida": str(config.PASTA_SAIDAS),
        "pasta_onedrive": str(config.PASTA_ONEDRIVE),
        "onedrive_existe": config.PASTA_ONEDRIVE.exists(),
        "api_grafeno": fonte_grafeno.api_disponivel(),
        "aviso_taxa": "",
        "avisos": [],
    }

    hist = historico_mod.Historico()
    dados["numero_rodada"] = hist.proximo_numero()
    dados["rodadas_no_historico"] = len(hist.rodadas)
    dados["titulos_no_historico"] = len(hist.ids_recomprados)

    inicio, fim = calculo.sugerir_janela(hoje)
    dados["janela_inicio"] = nz.iso(inicio)
    dados["janela_fim"] = nz.iso(fim)
    dados["data_recompra"] = hoje.isoformat()

    #  A janela que emenda com a rodada anterior vence a sugerida pela data:
    #  é ela que não deixa lacuna nem sobreposição (seção 8.5).
    _, fim_anterior = hist.janela_anterior()
    if fim_anterior:
        emenda_inicio, emenda_fim = calculo.janela_seguinte(fim_anterior)
        dados["janela_inicio"] = nz.iso(emenda_inicio)
        dados["janela_fim"] = nz.iso(emenda_fim)
        dados["avisos"].append(
            f"Janela pré-preenchida para emendar com a rodada "
            f"{hist.ultima.get('numero')}, que terminou em {nz.br(fim_anterior)}."
        )

    if not template.exists():
        dados["avisos"].append(
            f"Template ausente: coloque uma cópia do arquivo oficial de recompra "
            f"em {config.PASTA_TEMPLATES} com o nome {template.name}."
        )
        return dados

    try:
        parametros = calculo.ler_parametros(template)
    except Exception as erro:                                # noqa: BLE001
        dados["avisos"].append(f"Não consegui ler a aba INFORMAÇÕES: {erro}")
        return dados
    finally:
        leitura.fechar_cache()

    dados["juros_mora"] = f"{parametros.juros_mora * 100:.6f}".rstrip("0").rstrip(".")
    dados["multa"] = f"{parametros.multa * 100:.2f}".rstrip("0").rstrip(".")

    juros_antes, multa_antes = hist.taxas_anteriores()
    mudou = []
    if juros_antes is not None and juros_antes != parametros.juros_mora:
        mudou.append(f"juros de {juros_antes * 100:.6f}% para "
                     f"{parametros.juros_mora * 100:.6f}% a.d.")
    if multa_antes is not None and multa_antes != parametros.multa:
        mudou.append(f"multa de {multa_antes * 100:.2f}% para "
                     f"{parametros.multa * 100:.2f}%")
    if mudou:
        dados["aviso_taxa"] = ("Mudou em relação à rodada anterior: "
                               + "; ".join(mudou) + ". Confira a aba INFORMAÇÕES.")
    return dados


def sugerir_janela(data_recompra):
    """A janela quinzenal da data escolhida, para a tela repreencher sozinha."""
    inicio, fim = calculo.sugerir_janela(nz.data(data_recompra))
    return {"inicio": nz.iso(inicio), "fim": nz.iso(fim)}


def testar_vortx():
    """Confere o acesso ao PostgreSQL e devolve a posição mais recente."""
    dados = fonte_vortx.testar_conexao()
    return {
        "ok": True,
        "mensagem": (f"Conectado. DataGeracao mais recente: "
                     f"{nz.br(dados['data_geracao'])} · {dados['linhas']:,} linhas."
                     .replace(",", ".")),
        "data_geracao": nz.iso(dados["data_geracao"]),
        "linhas": dados["linhas"],
    }


def testar_grafeno():
    """Enquanto a conta não sai, a resposta honesta é 'ainda não' (pendência 14.2-3)."""
    if not fonte_grafeno.api_disponivel():
        return {"ok": False,
                "mensagem": "Sem token da API Grafeno no .env (pendência 14.2-3). "
                            "Use a opção Arquivo."}
    return {"ok": False,
            "mensagem": "Token encontrado, mas a extração via API ainda não foi "
                        "implementada (Fase 5). Use a opção Arquivo."}


# ============================================================================
#  4. Fluxo em dois tempos
# ============================================================================
def cfg_do_formulario(q):
    """Traduz os campos da tela para o dicionário que o motor entende."""
    def texto(nome, padrao=""):
        return (q.get(nome, [padrao])[0] or "").strip()

    def ligado(nome):
        return texto(nome, "0") == "1"

    fonte_v = texto("fonte_vortx", "arquivo")
    fonte_g = texto("fonte_grafeno", "arquivo")
    cfg = {
        "numero_rodada": texto("numero_rodada") or None,
        "data_recompra": texto("data_recompra") or None,
        "janela_inicio": texto("janela_inicio") or None,
        "janela_fim": texto("janela_fim") or None,
        "template": texto("template") or str(config.TEMPLATE_PADRAO),

        "fonte_vortx": fonte_v,
        "arquivo_vortx": texto("arquivo_vortx"),
        "aba_vortx": texto("aba_vortx", config.ABA_VORTX) or config.ABA_VORTX,

        "fonte_grafeno": fonte_g,
        "arquivo_grafeno": texto("arquivo_grafeno"),
        "aba_grafeno": texto("aba_grafeno", config.ABA_GRAFENO) or config.ABA_GRAFENO,
        "data_extracao_grafeno": texto("data_extracao_grafeno") or None,

        "pasta_saida": texto("pasta_saida") or str(config.PASTA_SAIDAS),
        "copiar_onedrive": ligado("copiar_onedrive"),
        "atualizar_historico": ligado("atualizar_historico"),
        "modo_simulacao": ligado("modo_simulacao"),
        "gravar_snapshot": True,
    }

    if fonte_v == "arquivo" and not cfg["arquivo_vortx"]:
        raise RuntimeError("Escolha o arquivo da Vórtx (ou marque a fonte "
                           "'Banco de dados').")
    if fonte_g == "arquivo" and not cfg["arquivo_grafeno"]:
        raise RuntimeError("Escolha o arquivo da Grafeno.")
    for rotulo, chave in (("Vórtx", "arquivo_vortx"), ("Grafeno", "arquivo_grafeno")):
        caminho = cfg[chave]
        if caminho and not Path(caminho).exists():
            raise RuntimeError(f"Arquivo da {rotulo} não encontrado:\n{caminho}")
    if not cfg["data_recompra"]:
        raise RuntimeError("Informe a data da recompra.")
    return cfg


def conciliar(cfg, log):
    """Primeiro tempo. Devolve a sessão guardada, sem escrever nada definitivo."""
    log("Conciliação — nada é gravado nesta etapa.", "destaque")
    if cfg["modo_simulacao"]:
        log("Modo simulação ligado: as validações que bloqueariam viram aviso, o "
            "arquivo sai com sufixo _SIMULACAO e o histórico não é atualizado.",
            "aviso")
    resultado = motor.conciliar(cfg, log=log)
    sessao = guardar(Sessao(cfg, resultado))

    log("")
    if resultado.bloqueado:
        log(f"Emissão BLOQUEADA por {len(resultado.erros)} validação(ões) em ERRO. "
            f"Corrija o que o painel aponta e concilie de novo.", "erro")
    else:
        log("Nenhum bloqueio. O botão de gerar está liberado.", "ok")
    return sessao


def gerar(sessao, cfg_saida, log):
    """Segundo tempo. Escreve a planilha, as exceções e o histórico."""
    cfg = dict(sessao.cfg)
    cfg.update(cfg_saida)
    arquivos = motor.gerar(sessao.resultado, cfg, log=log)
    for nome, caminho in arquivos.items():
        sessao.registrar_arquivo(nome, caminho)
    return sessao


def rodar_retroativo(cfg, log):
    """
    Reprocessa as rodadas anteriores do OneDrive (seção 12).

    Fora do fluxo em dois tempos de propósito: não emite Termo nenhum, não toca
    no histórico e não escreve no OneDrive. Só lê, copia para a pasta de
    trabalho e produz a planilha do passivo.
    """
    pasta = cfg.get("pasta_onedrive") or config.PASTA_ONEDRIVE
    destino = Path(cfg.get("pasta_saida") or config.PASTA_SAIDAS)
    log("Reprocessamento retroativo — leitura apenas, nada é escrito no OneDrive.",
        "destaque")
    rodadas, caminho = retroativo.rodar(
        raiz=pasta, destino=destino, apenas=cfg.get("apenas"), log=log)

    completas = [r for r in rodadas if r.situacao == retroativo.SITUACAO_COMPLETA]
    faltantes = sum(len(r.faltantes) for r in completas)
    indevidos = sum(len(r.indevidos) for r in completas)
    duplicados = sum(len(r.duplicados) for r in rodadas if r.emitidos)
    return {
        "arquivo": caminho.name,
        "caminho": str(caminho),
        "rodadas": len(rodadas),
        "reconciliadas": len(completas),
        "faltantes": faltantes,
        "indevidos": indevidos,
        "duplicados": duplicados,
        "valor_faltante": str(calculo.arredondar(
            calculo.somar(r.valor_faltante for r in completas))),
        "valor_duplicado": str(calculo.arredondar(
            calculo.somar(r.valor_duplicado for r in rodadas if r.emitidos))),
        "detalhe": [
            {"numero": r.numero, "situacao": r.situacao,
             "arquivo": r.arquivo.name if r.arquivo else "—",
             "emitidos": len(r.emitidos), "distintos": r.titulos_distintos,
             "motor": len(r.calculados), "faltantes": len(r.faltantes),
             "indevidos": len(r.indevidos), "duplicados": len(r.duplicados),
             "motivo": r.motivo}
            for r in sorted(rodadas, key=lambda x: x.numero)
        ],
    }


def gerar_excecoes(sessao, log):
    """
    O relatório de exceções, sozinho — disponível já depois de conciliar.

    É o que a decisão 2.2 pede: quem opera revisa as exceções **antes** de
    decidir emitir. Este arquivo não é o Termo e não mexe no histórico, então
    não passa pela trava da regra 15.5.
    """
    destino = Path(sessao.cfg.get("pasta_saida") or config.PASTA_SAIDAS)
    destino.mkdir(parents=True, exist_ok=True)
    caminho = excecoes.gerar(sessao.resultado, destino, log=log)
    return sessao.registrar_arquivo("excecoes", caminho)


# ---------------------------------------------------------------------------
#  O que a tela recebe depois de conciliar
# ---------------------------------------------------------------------------
def payload_resultado(sessao, limite_previa=12):
    r = sessao.resultado
    return {
        "token": sessao.token,
        "resumo": r.resumo(),
        "validacoes": [v.como_dicionario() for v in r.validacoes],
        "previa": _previa(r, limite_previa),
        "total_linhas": len(r.linhas_termo),
        "bloqueado": r.bloqueado,
        "modo_simulacao": r.modo_simulacao,
        "arquivo_sugerido": excel_saida.nome_arquivo(r),
        "snapshot": str(r.pasta_snapshot or ""),
        "pasta_saida": str(sessao.cfg.get("pasta_saida") or config.PASTA_SAIDAS),
    }


def _previa(r, limite):
    linhas = []
    for linha in r.linhas_termo[:limite]:
        linhas.append({
            "nome": linha["nome"],
            "documento": linha["documento_formatado"],
            "id_titulo": linha["id_titulo"],
            "vencimento": nz.br(linha["vencimento"]),
            "prazo": linha["prazo"],
            "valor_nominal": f"{calculo.arredondar(linha['valor_nominal']):.2f}",
            "valor_recompra": f"{calculo.arredondar(linha['valor_recompra']):.2f}",
            "boletos": linha["qtd_boletos"],
        })
    return linhas


# ============================================================================
#  5. Interface HTML
# ============================================================================
PAGINA_HTML = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Recompra FIDC · Leve Saúde — Netz Asset</title>
<style>
  :root{
    --navy:#001B5C; --navy2:#002F6C; --orange:#FF965A; --orange-d:#f57f3a;
    --blue:#116DFF; --bg:#eef3fa; --card:#ffffff; --line:#dfe6f1;
    --txt:#12233f; --muted:#5b6b8c; --ink:#001B5C;
    --ok:#0a8f68; --ok-bg:#e7f7f1; --aviso:#a86a00; --aviso-bg:#fff5e2;
    --erro:#c62828; --erro-bg:#fdecec;
    --con-bg:#04123a; --con-line:#12245e;
    --font:"IBM Plex Sans","Segoe UI",system-ui,Arial,sans-serif;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:var(--font);background:var(--bg);color:var(--txt)}
  .topbar{background:var(--navy)}
  .topbar .in{max-width:1080px;margin:0 auto;padding:16px 20px;display:flex;align-items:center;gap:16px}
  .brand{height:30px;width:auto;display:block}
  .topbar .tag{color:#aebfe0;font-size:13px;border-left:1px solid #24407e;padding-left:16px}
  .topbar .tag b{color:#fff}
  .wrap{max-width:1080px;margin:0 auto;padding:26px 20px 60px}
  .pagettl h1{font-size:20px;margin:0;font-weight:700;color:var(--ink)}
  .pagettl p{color:var(--muted);font-size:13px;margin:4px 0 0}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:22px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:20px;box-shadow:0 1px 3px rgba(0,27,92,.06)}
  .card.full{grid-column:1/-1}
  .card h2{font-size:12px;margin:0 0 16px;color:var(--navy);text-transform:uppercase;letter-spacing:.7px;font-weight:700}
  .card h3{font-size:12px;margin:18px 0 10px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;font-weight:700;border-top:1px solid var(--line);padding-top:14px}
  .card h3:first-of-type{border-top:0;margin-top:0;padding-top:0}
  label{display:block;font-size:13px;margin:0 0 6px;color:var(--txt);font-weight:600}
  .hint{color:var(--muted);font-size:12px;margin:4px 0 0;font-weight:400;line-height:1.5}
  input[type=text],input[type=number],input[type=date],select{
    width:100%;padding:11px 12px;border-radius:9px;border:1px solid var(--line);
    background:#fff;color:var(--txt);font-size:14px;outline:none;font-family:var(--font)}
  input:focus,select:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(17,109,255,.12)}
  input[readonly]{background:#f5f8fd;color:var(--muted)}
  .field{margin-bottom:16px}
  .field:last-child{margin-bottom:0}
  .dois{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .seg{display:flex;background:#eaf0fa;border:1px solid var(--line);border-radius:9px;overflow:hidden;padding:3px;gap:3px}
  .seg button{flex:1;padding:9px;background:transparent;border:0;border-radius:7px;color:var(--navy);cursor:pointer;font-size:14px;font-weight:600;font-family:var(--font)}
  .seg button.active{background:var(--navy);color:#fff}
  .filerow{display:flex;gap:8px}
  .filerow input{flex:1;min-width:0}
  .btn{padding:11px 16px;border-radius:9px;border:1px solid var(--navy);background:#fff;
       color:var(--navy);cursor:pointer;font-size:14px;font-weight:600;font-family:var(--font);white-space:nowrap}
  .btn:hover:not(:disabled){background:var(--navy);color:#fff}
  .btn:disabled{opacity:.45;cursor:not-allowed}
  .btn.small{padding:8px 12px;font-size:13px}
  .switch{display:flex;align-items:center;justify-content:space-between;gap:14px;padding:12px 0;border-top:1px solid var(--line)}
  .switch > div > div{font-weight:600;color:var(--txt);font-size:13px}
  .toggle{position:relative;width:46px;height:26px;flex:0 0 auto}
  .toggle input{display:none}
  .track{position:absolute;inset:0;background:#c9d4e8;border-radius:20px;transition:.15s}
  .thumb{position:absolute;top:3px;left:3px;width:20px;height:20px;border-radius:50%;background:#fff;transition:.15s;box-shadow:0 1px 2px rgba(0,0,0,.25)}
  .toggle input:checked + .track{background:var(--navy)}
  .toggle input:checked + .track .thumb{left:23px}
  .go{width:100%;margin-top:16px;padding:14px;border-radius:11px;border:0;font-size:15px;font-weight:700;
      background:var(--orange);color:#3a1a00;cursor:pointer;font-family:var(--font)}
  .go:hover:not(:disabled){background:var(--orange-d)}
  .go:disabled{opacity:.55;cursor:not-allowed}
  .go.secundario{background:var(--navy);color:#fff}
  .go.secundario:hover:not(:disabled){background:var(--navy2)}
  .nota{font-size:12px;padding:10px 12px;border-radius:9px;margin-top:10px;line-height:1.5}
  .nota.ok{background:var(--ok-bg);color:var(--ok)}
  .nota.aviso{background:var(--aviso-bg);color:var(--aviso)}
  .nota.erro{background:var(--erro-bg);color:var(--erro)}
  .nota:empty{display:none}
  #painel{display:none;margin-top:16px}
  .numeros{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
  .num{background:#f7faff;border:1px solid var(--line);border-radius:11px;padding:14px}
  .num .rot{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;font-weight:700}
  .num .val{font-size:20px;font-weight:700;color:var(--ink);margin-top:6px}
  .val.small{font-size:16px}
  ul.vals{list-style:none;margin:0;padding:0}
  ul.vals li{display:flex;gap:12px;padding:11px 0;border-top:1px solid var(--line);align-items:flex-start}
  ul.vals li:first-child{border-top:0}
  .selo{flex:0 0 auto;font-size:11px;font-weight:700;padding:3px 9px;border-radius:20px;letter-spacing:.4px;margin-top:2px}
  .selo.OK{background:var(--ok-bg);color:var(--ok)}
  .selo.AVISO{background:var(--aviso-bg);color:var(--aviso)}
  .selo.ERRO{background:var(--erro-bg);color:var(--erro)}
  .vtxt b{display:block;font-size:13px}
  .vtxt span{font-size:13px;color:var(--txt)}
  .vtxt em{display:block;font-size:12px;color:var(--muted);font-style:normal;margin-top:3px;line-height:1.5}
  .tabela-rolo{overflow-x:auto;margin-top:4px}
  table{border-collapse:collapse;width:100%;font-size:12.5px}
  th{text-align:left;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;font-size:11px;padding:8px 10px;border-bottom:1px solid var(--line);white-space:nowrap}
  td{padding:8px 10px;border-bottom:1px solid #f0f4fa;white-space:nowrap}
  td.n{text-align:right;font-variant-numeric:tabular-nums}
  .acoes{display:flex;gap:12px;flex-wrap:wrap;margin-top:16px}
  .acoes .go{width:auto;flex:1;min-width:220px;margin-top:0}
  .console{margin-top:22px;background:var(--con-bg);border:1px solid var(--con-line);border-radius:14px;overflow:hidden}
  .console .bar{display:flex;align-items:center;gap:8px;padding:10px 14px;border-bottom:1px solid var(--con-line);color:#aebfe0;font-size:12px}
  .dot{width:10px;height:10px;border-radius:50%;background:#3a4a7a}
  .dot.on{background:#2fd4a7;box-shadow:0 0 8px #2fd4a7}
  .dot.busy{background:var(--orange);box-shadow:0 0 8px var(--orange)}
  .dot.err{background:#ff6b6b;box-shadow:0 0 8px #ff6b6b}
  pre#log{margin:0;padding:16px;height:300px;overflow:auto;font-family:"Cascadia Code",Consolas,monospace;font-size:13px;line-height:1.55;white-space:pre-wrap;color:#dbe5f5}
  .l-info{color:#cdd8ef} .l-ok{color:#48e0aa} .l-erro{color:#ff7b7b}
  .l-aviso{color:#ffc65c} .l-destaque{color:#ffb07d;font-weight:700}
  .foot{color:var(--muted);font-size:12px;margin-top:18px;text-align:center;line-height:1.6}
  @media(max-width:880px){.grid{grid-template-columns:1fr}.numeros{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<div class="topbar"><div class="in">
  <svg class="brand" viewBox="389.66 389.65 1407.66 300.69" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Netz Asset">
    <path fill="#ff965a" d="m537.274 537.267 73.8-73.8 73.801 73.8-73.8 73.8z"/>
    <path fill="#ffffff" d="m684.87 684.87 147.61-147.6V389.66L684.87 537.27z"/>
    <path fill="#ffffff" d="m389.66 684.87 147.61-147.6V389.66L389.66 537.27z"/>
    <path fill="#ffffff" d="M1096.23 437.44c-40.75 0-71.65 33.25-75.87 79.14v-73.52h-40.28v241.66h40.28V526.89c0-37.46 23.41-56.67 57.14-56.67 40.27 0 57.14 26.23 57.14 56.67v157.83h40.28V526.89c0-51.05-33.72-89.45-78.68-89.45Z"/>
    <path fill="#ffffff" d="M1319.63 437.43c-66.51 0-110.06 51.05-110.06 126.45s43.55 126.46 110.06 126.46c53.86 0 92.26-34.19 105.38-86.18h-41.69c-10.77 31.84-34.19 51.52-63.69 51.52-38.41 0-66.98-34.19-69.78-84.31h179.84c0-85.23-42.16-133.94-110.06-133.94m-67.44 101.17c7.49-40.28 33.72-66.51 67.44-66.51s59.95 26.23 67.44 66.51z"/>
    <path fill="#ffffff" d="M1527.56 389.66h-40.28v53.39h-41.21v34.66h41.21v147.52c0 38.88 22.02 59.48 63.7 59.48h17.8v-34.66h-18.27c-14.99 0-22.95-8.44-22.95-24.83V477.7h41.22v-34.66h-41.22v-53.39Z"/>
    <path fill="#ffffff" d="M1797.32 443.21h-193.89v34.5h176.85L1603.43 618.6v66.27h193.89v-34.81h-176.84L1797.32 514z"/>
  </svg>
  <span class="tag">Ferramenta interna &middot; <b>Recompra FIDC</b></span>
</div></div>
<div class="wrap">
  <div class="pagettl">
    <h1>Recompra FIDC · Leve Saúde</h1>
    <p>Concilia o estoque da <b>Vórtx</b> com os boletos da <b>Grafeno</b>, apura o Termo de Recompra
       e gera a planilha da rodada. <b>Conciliar</b> não grava nada — a emissão é o segundo passo.</p>
  </div>

  <div class="grid">
    <!-- ---------------------------------------------------------------- -->
    <div class="card">
      <h2>Parâmetros da rodada</h2>
      <div class="dois">
        <div class="field">
          <label>Nº da recompra</label>
          <input type="number" id="numero_rodada" min="1" step="1">
        </div>
        <div class="field">
          <label>Data da recompra</label>
          <input type="date" id="data_recompra">
        </div>
      </div>
      <div class="dois">
        <div class="field">
          <label>Início da janela</label>
          <input type="date" id="janela_inicio">
        </div>
        <div class="field">
          <label>Fim da janela</label>
          <input type="date" id="janela_fim">
        </div>
      </div>
      <p class="hint">Quinzena estrita: dia 1 ao 15, ou dia 16 ao último dia do mês.
         Muda sozinha quando você troca a data da recompra — pode ajustar à mão.</p>

      <h3>Juros e multa · da aba INFORMAÇÕES</h3>
      <div class="dois">
        <div class="field">
          <label>Juros mora (% a.d.)</label>
          <input type="text" id="juros_mora" readonly>
        </div>
        <div class="field">
          <label>Multa (%)</label>
          <input type="text" id="multa" readonly>
        </div>
      </div>
      <p class="hint">Somente leitura, sempre lidos do template. Para mudar a taxa,
         edite a aba INFORMAÇÕES do arquivo — nunca o código.</p>
      <div class="nota aviso" id="avisoTaxa"></div>
    </div>

    <!-- ---------------------------------------------------------------- -->
    <div class="card">
      <h2>Fontes de dados</h2>

      <h3>Vórtx · estoque de títulos</h3>
      <div class="seg" id="segVortx">
        <button type="button" data-fonte="banco">Banco de dados</button>
        <button type="button" data-fonte="arquivo" class="active">Arquivo</button>
      </div>
      <div id="vortxBanco" style="display:none;margin-top:12px">
        <button type="button" class="btn small" id="testarVortx">Testar conexão</button>
        <div class="nota" id="notaVortx"></div>
      </div>
      <div id="vortxArquivo" style="margin-top:12px">
        <div class="filerow">
          <input type="text" id="arquivo_vortx" placeholder="C:\...\Leve_Saude_12_Recompra....xlsx">
          <button type="button" class="btn" data-alvo="arquivo_vortx">Procurar…</button>
        </div>
        <p class="hint">Aba <input type="text" id="aba_vortx" style="width:120px;display:inline-block;padding:4px 8px;font-size:12px" value="VORTX">
           — .xlsx exportado do portal ou uma rodada anterior; .csv também serve.</p>
      </div>

      <h3>Grafeno · estoque de boletos</h3>
      <div class="seg" id="segGrafeno">
        <button type="button" data-fonte="api">API</button>
        <button type="button" data-fonte="arquivo" class="active">Arquivo</button>
      </div>
      <div id="grafenoApi" style="display:none;margin-top:12px">
        <button type="button" class="btn small" id="testarGrafeno">Testar conexão</button>
        <div class="nota" id="notaGrafeno"></div>
      </div>
      <div id="grafenoArquivo" style="margin-top:12px">
        <div class="filerow">
          <input type="text" id="arquivo_grafeno" placeholder="C:\...\cobrancas.xlsx">
          <button type="button" class="btn" data-alvo="arquivo_grafeno">Procurar…</button>
        </div>
        <p class="hint">Aba <input type="text" id="aba_grafeno" style="width:120px;display:inline-block;padding:4px 8px;font-size:12px" value="GRAFENO">
           — o mesmo arquivo pode servir às duas fontes, quando for uma rodada anterior.</p>
        <div class="field" style="margin-top:12px">
          <label>Data da extração da Grafeno</label>
          <input type="date" id="data_extracao_grafeno">
          <p class="hint">O export não traz essa data em coluna nenhuma. Em branco, vale a data de
             modificação do arquivo. É por ela que a defasagem é medida.</p>
        </div>
      </div>
    </div>

    <!-- ---------------------------------------------------------------- -->
    <div class="card full">
      <h2>Saída</h2>
      <div class="dois">
        <div class="field">
          <label>Pasta de destino</label>
          <div class="filerow">
            <input type="text" id="pasta_saida">
            <button type="button" class="btn" data-alvo="pasta_saida" data-tipo="pasta">Procurar…</button>
          </div>
        </div>
        <div>
          <div class="switch" style="border-top:0">
            <div>
              <div>Copiar também para a pasta do OneDrive</div>
              <p class="hint" style="margin:2px 0 0">Nunca sobrescreve: se já existir arquivo com o mesmo nome, a cópia é recusada.</p>
            </div>
            <label class="toggle"><input type="checkbox" id="copiar_onedrive"><span class="track"><span class="thumb"></span></span></label>
          </div>
          <div class="switch">
            <div>
              <div>Atualizar histórico de recompras ao final</div>
              <p class="hint" style="margin:2px 0 0">É a trava contra recomprar o mesmo título duas vezes. Deixe ligado.</p>
            </div>
            <label class="toggle"><input type="checkbox" id="atualizar_historico" checked><span class="track"><span class="thumb"></span></span></label>
          </div>
          <div class="switch">
            <div>
              <div>Modo simulação</div>
              <p class="hint" style="margin:2px 0 0">Para conferir rodadas antigas: o que bloquearia vira aviso, o arquivo sai com sufixo _SIMULACAO e o histórico não é tocado.</p>
            </div>
            <label class="toggle"><input type="checkbox" id="modo_simulacao"><span class="track"><span class="thumb"></span></span></label>
          </div>
        </div>
      </div>
      <div class="nota aviso" id="avisosGerais"></div>
      <button class="go" id="btnConciliar">▶  Conciliar</button>
    </div>
  </div>

  <!-- ------------------------------------------------------------------ -->
  <div id="painel">
    <div class="card full" style="margin-bottom:16px">
      <h2>Resumo da rodada</h2>
      <div class="numeros">
        <div class="num"><div class="rot">Títulos elegíveis</div><div class="val" id="rTitulos">—</div></div>
        <div class="num"><div class="rot">Valor nominal</div><div class="val small" id="rNominal">—</div></div>
        <div class="num"><div class="rot">Valor de recompra</div><div class="val small" id="rRecompra">—</div></div>
        <div class="num"><div class="rot">Exceções</div><div class="val" id="rExcecoes">—</div></div>
      </div>
      <div class="nota" id="notaDefasagem"></div>
      <p class="hint" id="rDetalhe"></p>
    </div>

    <div class="card full" style="margin-bottom:16px">
      <h2>Validações</h2>
      <ul class="vals" id="listaValidacoes"></ul>
    </div>

    <div class="card full">
      <h2>Prévia do Termo</h2>
      <div class="tabela-rolo">
        <table id="tabelaPrevia">
          <thead><tr>
            <th>Sacado</th><th>CPF/CNPJ</th><th>Título</th><th>Vencimento</th>
            <th>Prazo</th><th>Valor nominal</th><th>Valor de recompra</th><th>Boletos</th>
          </tr></thead>
          <tbody></tbody>
        </table>
      </div>
      <p class="hint" id="previaRodape"></p>
      <div class="acoes">
        <button class="go" id="btnGerar">▶  Gerar Termo e planilha</button>
        <button class="go secundario" id="btnExcecoes">⬇  Baixar relatório de exceções</button>
      </div>
      <div class="nota" id="notaGerar"></div>
    </div>
  </div>

  <!-- ------------------------------------------------------------------ -->
  <div class="card full" style="margin-top:16px">
    <h2>Reprocessamento retroativo · seção 12</h2>
    <p class="hint" style="margin:0 0 14px">
      Roda o motor corrigido sobre as rodadas já emitidas, na pasta do OneDrive, e compara
      com o Termo que de fato saiu. Quantifica o passivo: títulos que deveriam ter entrado e
      não entraram, títulos que entraram sem dever e linhas cobradas em duplicidade.
      <b>Só lê</b> — copia cada arquivo para <code>snapshots/onedrive/</code> e trabalha sobre
      a cópia. Não emite Termo e não toca no histórico.
    </p>
    <div class="dois">
      <div class="field">
        <label>Pasta das rodadas anteriores</label>
        <div class="filerow">
          <input type="text" id="pasta_onedrive">
          <button type="button" class="btn" data-alvo="pasta_onedrive" data-tipo="pasta">Procurar…</button>
        </div>
      </div>
      <div class="field">
        <label>Rodadas (em branco = todas)</label>
        <input type="text" id="apenas" placeholder="ex.: 11 12">
        <p class="hint">Cada rodada reconciliada lê duas bases de dezenas de MB — rodar as
           doze leva alguns minutos.</p>
      </div>
    </div>
    <button class="go secundario" id="btnRetroativo" style="width:auto;min-width:280px">
      ▶  Reprocessar rodadas anteriores
    </button>
    <div class="nota" id="notaRetroativo"></div>
    <div class="tabela-rolo" id="rolaRetroativo" style="display:none;margin-top:14px">
      <table id="tabelaRetroativo">
        <thead><tr>
          <th>Rodada</th><th>Situação</th><th>Arquivo</th><th>Linhas no Termo</th>
          <th>Títulos</th><th>Pelo motor</th><th>Faltaram</th><th>Sem dever</th><th>Repetidos</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <div class="console">
    <div class="bar"><span class="dot" id="dot"></span><span id="status">Pronto.</span></div>
    <pre id="log"></pre>
  </div>

  <p class="foot">Netz Asset · roda 100% no seu computador · leitura apenas: nada é escrito no banco da Vórtx,
     na Grafeno ou no OneDrive sem você marcar a caixa.<br>
     Cada rodada gera snapshot em <b>snapshots/</b> e log em <b>logs/</b>.</p>
</div>

<script>
  const $ = id => document.getElementById(id);
  const logEl = $("log"), dot = $("dot"), status = $("status");
  const btnConciliar = $("btnConciliar"), btnGerar = $("btnGerar"), btnExcecoes = $("btnExcecoes");
  const btnRetroativo = $("btnRetroativo");
  let fonteVortx = "arquivo", fonteGrafeno = "arquivo";
  let token = null, ocupado = false, bloqueado = true, emitido = false;

  const moeda = v => Number(v).toLocaleString("pt-BR", {style:"currency", currency:"BRL"});

  // ---- estado inicial ------------------------------------------------------
  fetch("/inicial").then(r => r.json()).then(d => {
    $("numero_rodada").value = d.numero_rodada;
    $("data_recompra").value = d.data_recompra;
    $("janela_inicio").value = d.janela_inicio;
    $("janela_fim").value = d.janela_fim;
    $("data_extracao_grafeno").value = d.data_recompra;
    $("pasta_saida").value = d.pasta_saida;
    $("pasta_onedrive").value = d.pasta_onedrive;
    if(!d.onedrive_existe){
      $("notaRetroativo").className = "nota aviso";
      $("notaRetroativo").textContent =
        "Pasta das rodadas anteriores não encontrada — confira o caminho ou se o OneDrive está sincronizado.";
    }
    $("juros_mora").value = d.juros_mora ? d.juros_mora + " %" : "—";
    $("multa").value = d.multa ? d.multa + " %" : "—";
    if(d.aviso_taxa) $("avisoTaxa").textContent = d.aviso_taxa;
    if(d.avisos && d.avisos.length){
      const caixa = $("avisosGerais");
      caixa.textContent = d.avisos.join("  ·  ");
      caixa.className = "nota " + (d.template_existe ? "aviso" : "erro");
    }
    if(!d.api_grafeno){
      $("segGrafeno").querySelector('[data-fonte="api"]').disabled = true;
      $("segGrafeno").querySelector('[data-fonte="api"]').title =
        "Sem token da API no .env — a conta ainda não foi liberada (pendência 14.2-3).";
    }
    addLine("Pronto. Rodada " + d.numero_rodada + " sugerida a partir do histórico ("
            + d.rodadas_no_historico + " rodada(s), " + d.titulos_no_historico
            + " títulos já recomprados).", "info");
    if(!d.template_existe) addLine("⚠️  Template ausente em templates/ — a conciliação não roda sem ele.", "erro");
  });

  // ---- janela quinzenal segue a data da recompra ---------------------------
  $("data_recompra").addEventListener("change", async () => {
    const d = $("data_recompra").value;
    if(!d) return;
    const j = await (await fetch("/janela?data=" + d)).json();
    $("janela_inicio").value = j.inicio;
    $("janela_fim").value = j.fim;
  });

  // ---- seletores de fonte --------------------------------------------------
  function ligarSeg(idSeg, aoTrocar){
    $(idSeg).addEventListener("click", e => {
      const b = e.target.closest("button");
      if(!b || b.disabled) return;
      [...$(idSeg).children].forEach(x => x.classList.toggle("active", x === b));
      aoTrocar(b.dataset.fonte);
    });
  }
  ligarSeg("segVortx", f => {
    fonteVortx = f;
    $("vortxBanco").style.display   = f === "banco"   ? "block" : "none";
    $("vortxArquivo").style.display = f === "arquivo" ? "block" : "none";
  });
  ligarSeg("segGrafeno", f => {
    fonteGrafeno = f;
    $("grafenoApi").style.display      = f === "api"     ? "block" : "none";
    $("grafenoArquivo").style.display  = f === "arquivo" ? "block" : "none";
  });

  // ---- Procurar… -----------------------------------------------------------
  document.querySelectorAll("button[data-alvo]").forEach(b => {
    b.addEventListener("click", async () => {
      status.textContent = "Abrindo seletor…";
      try{
        const tipo = b.dataset.tipo || "arquivo";
        const j = await (await fetch("/procurar?tipo=" + tipo)).json();
        if(j.path) $(b.dataset.alvo).value = j.path;
        status.textContent = j.path ? "Selecionado." : "Nada escolhido.";
      }catch(err){ status.textContent = "Não foi possível abrir o seletor: " + err; }
    });
  });

  // ---- testes de conexão ---------------------------------------------------
  async function testar(rota, caixa, botao){
    botao.disabled = true;
    caixa.className = "nota"; caixa.textContent = "Testando…";
    try{
      const j = await (await fetch(rota)).json();
      caixa.className = "nota " + (j.ok ? "ok" : "aviso");
      caixa.textContent = j.mensagem;
    }catch(err){
      caixa.className = "nota erro"; caixa.textContent = "Falhou: " + err;
    }
    botao.disabled = false;
  }
  $("testarVortx").addEventListener("click", () => testar("/testar-vortx", $("notaVortx"), $("testarVortx")));
  $("testarGrafeno").addEventListener("click", () => testar("/testar-grafeno", $("notaGrafeno"), $("testarGrafeno")));

  // ---- console -------------------------------------------------------------
  function addLine(texto, nivel){
    const span = document.createElement("span");
    span.className = "l-" + (nivel || "info");
    span.textContent = texto + "\n";
    logEl.appendChild(span);
    logEl.scrollTop = logEl.scrollHeight;
  }

  function parametros(){
    return {
      numero_rodada: $("numero_rodada").value,
      data_recompra: $("data_recompra").value,
      janela_inicio: $("janela_inicio").value,
      janela_fim: $("janela_fim").value,
      fonte_vortx: fonteVortx,
      arquivo_vortx: $("arquivo_vortx").value,
      aba_vortx: $("aba_vortx").value,
      fonte_grafeno: fonteGrafeno,
      arquivo_grafeno: $("arquivo_grafeno").value,
      aba_grafeno: $("aba_grafeno").value,
      data_extracao_grafeno: $("data_extracao_grafeno").value,
      pasta_saida: $("pasta_saida").value,
      copiar_onedrive: $("copiar_onedrive").checked ? "1" : "0",
      atualizar_historico: $("atualizar_historico").checked ? "1" : "0",
      modo_simulacao: $("modo_simulacao").checked ? "1" : "0",
    };
  }

  function ocupar(mensagem){
    ocupado = true;
    btnConciliar.disabled = btnGerar.disabled = btnExcecoes.disabled = true;
    btnRetroativo.disabled = true;
    dot.className = "dot busy"; status.textContent = mensagem;
  }
  function liberar(classe, mensagem, podeGerar){
    ocupado = false;
    btnConciliar.disabled = btnRetroativo.disabled = false;
    btnGerar.disabled = !podeGerar;
    btnExcecoes.disabled = !token;
    dot.className = "dot " + classe; status.textContent = mensagem;
  }

  // ---- 1º tempo: conciliar -------------------------------------------------
  btnConciliar.addEventListener("click", () => {
    if(ocupado) return;
    logEl.innerHTML = ""; $("painel").style.display = "none";
    token = null; bloqueado = true; emitido = false;
    ocupar("Conciliando…");
    const es = new EventSource("/conciliar?" + new URLSearchParams(parametros()));
    es.onmessage = ev => { const m = JSON.parse(ev.data); addLine(m.msg, m.nivel); };
    es.addEventListener("fim", ev => {
      es.close();
      const r = JSON.parse(ev.data);
      token = r.token; bloqueado = r.bloqueado;
      mostrarPainel(r);
      liberar(bloqueado ? "err" : "on",
              bloqueado ? "Conciliado — emissão bloqueada." : "Conciliado.",
              !bloqueado);
    });
    es.addEventListener("erro", ev => {
      es.close();
      addLine("❌  " + JSON.parse(ev.data).msg, "erro");
      liberar("err", "Erro.", false);
    });
    es.onerror = () => {
      if(ocupado){ es.close(); addLine("❌  Conexão com o servidor perdida.", "erro");
                   liberar("err", "Conexão interrompida.", false); }
    };
  });

  // ---- painel --------------------------------------------------------------
  function mostrarPainel(r){
    const s = r.resumo;
    $("rTitulos").textContent   = s.titulos_elegiveis;
    $("rNominal").textContent   = moeda(s.valor_nominal);
    $("rRecompra").textContent  = moeda(s.valor_recompra);
    $("rExcecoes").textContent  = s.excecoes;
    $("rDetalhe").textContent   =
      "Rodada " + s.rodada + " · recompra em " + s.data_recompra + " · janela " + s.janela +
      " · " + s.contagens.titulos + " títulos e " + s.contagens.boletos + " boletos lidos · " +
      (s.contagens.cobertura * 100).toFixed(2) + "% de cobertura." +
      (r.modo_simulacao ? "  ⚠ Modo simulação." : "");

    //  Verde só quando a defasagem foi medida E deu zero. Sem as duas datas de
    //  extração não há defasagem zero — há defasagem desconhecida.
    const nd = $("notaDefasagem");
    nd.className = "nota " + (s.defasagem_medida && s.defasagem_dias === 0 ? "ok" : "erro");
    nd.textContent = s.defasagem_frase;

    const lista = $("listaValidacoes");
    lista.innerHTML = "";
    r.validacoes.forEach(v => {
      const li = document.createElement("li");
      li.innerHTML = '<span class="selo ' + v.nivel + '">' + v.nivel + '</span>' +
                     '<div class="vtxt"><b></b><span></span><em></em></div>';
      li.querySelector("b").textContent = v.titulo;
      li.querySelector("span:not(.selo)").textContent = v.detalhe;
      li.querySelector("em").textContent = v.nivel === "OK" ? "" : v.porque;
      lista.appendChild(li);
    });

    const corpo = $("tabelaPrevia").querySelector("tbody");
    corpo.innerHTML = "";
    r.previa.forEach(l => {
      const tr = document.createElement("tr");
      [l.nome, l.documento, l.id_titulo, l.vencimento].forEach(v => {
        const td = document.createElement("td"); td.textContent = v; tr.appendChild(td);
      });
      [l.prazo + " d", moeda(l.valor_nominal), moeda(l.valor_recompra), l.boletos].forEach(v => {
        const td = document.createElement("td"); td.className = "n"; td.textContent = v; tr.appendChild(td);
      });
      corpo.appendChild(tr);
    });
    $("previaRodape").textContent =
      r.previa.length + " de " + r.total_linhas + " linhas · arquivo que será gerado: " + r.arquivo_sugerido;

    const ng = $("notaGerar");
    if(r.bloqueado){
      ng.className = "nota erro";
      ng.textContent = "Emissão bloqueada por " + r.validacoes.filter(v => v.nivel === "ERRO").length +
        " validação(ões) em ERRO. Corrija e concilie de novo — ou ligue o modo simulação, " +
        "que emite com sufixo _SIMULACAO e sem tocar no histórico.";
    } else {
      ng.className = "nota ok";
      ng.textContent = "Sem bloqueios. Revise as exceções antes de gerar.";
    }
    $("painel").style.display = "block";
    $("painel").scrollIntoView({behavior:"smooth", block:"start"});
  }

  // ---- 2º tempo: gerar -----------------------------------------------------
  btnGerar.addEventListener("click", () => {
    if(ocupado || !token) return;
    ocupar("Gerando…");
    const p = parametros(); p.token = token;
    const es = new EventSource("/gerar?" + new URLSearchParams(p));
    es.onmessage = ev => { const m = JSON.parse(ev.data); addLine(m.msg, m.nivel); };
    es.addEventListener("fim", ev => {
      es.close();
      const r = JSON.parse(ev.data);
      addLine("", "info");
      addLine("──────── ARQUIVOS GERADOS ────────", "destaque");
      r.arquivos.forEach(a => addLine("  " + a.rotulo + ": " + a.caminho, "ok"));
      const ng = $("notaGerar");
      ng.className = "nota ok";
      ng.innerHTML = "Gerado. " + r.arquivos.map(a =>
        '<a href="/baixar?token=' + token + '&nome=' + a.nome + '">' + a.arquivo + "</a>").join(" · ") +
        ' — <a href="#" id="linkPasta">abrir a pasta</a>';
      //  A pasta abre no Explorer pelo servidor; sem o fetch, o clique levaria a
      //  página embora e o operador perderia o log da rodada.
      $("linkPasta").addEventListener("click", async ev => {
        ev.preventDefault();
        const j = await (await fetch("/abrir-pasta?token=" + token)).json();
        if(!j.ok) addLine("Não consegui abrir a pasta: " + (j.erro || ""), "aviso");
      });
      //  Emitido é emitido: reemitir a mesma conciliação regravaria o histórico
      //  e geraria um segundo arquivo com o mesmo nome. Concilie de novo.
      emitido = true;
      liberar("on", "Concluído.", false);
    });
    es.addEventListener("erro", ev => {
      es.close();
      addLine("❌  " + JSON.parse(ev.data).msg, "erro");
      liberar("err", "Erro.", !bloqueado && !emitido);
    });
    es.onerror = () => {
      if(ocupado){ es.close(); addLine("❌  Conexão com o servidor perdida.", "erro");
                   liberar("err", "Conexão interrompida.", !bloqueado && !emitido); }
    };
  });

  // ---- reprocessamento retroativo -----------------------------------------
  btnRetroativo.addEventListener("click", () => {
    if(ocupado) return;
    logEl.innerHTML = "";
    $("notaRetroativo").className = "nota";
    $("notaRetroativo").textContent = "";
    ocupar("Reprocessando rodadas anteriores…");
    const p = new URLSearchParams({
      pasta_saida: $("pasta_saida").value,
      pasta_onedrive: $("pasta_onedrive").value,
      apenas: $("apenas").value,
    });
    const es = new EventSource("/retroativo?" + p);
    es.onmessage = ev => { const m = JSON.parse(ev.data); addLine(m.msg, m.nivel); };
    es.addEventListener("fim", ev => {
      es.close();
      const r = JSON.parse(ev.data);
      const corpo = $("tabelaRetroativo").querySelector("tbody");
      corpo.innerHTML = "";
      r.detalhe.forEach(d => {
        const tr = document.createElement("tr");
        const reconciliada = d.situacao === "completa";
        [d.numero, d.situacao, d.arquivo].forEach(v => {
          const td = document.createElement("td"); td.textContent = v; tr.appendChild(td);
        });
        [d.emitidos || "—", d.distintos || "—",
         reconciliada ? d.motor : "—", reconciliada ? d.faltantes : "—",
         reconciliada ? d.indevidos : "—", d.duplicados].forEach(v => {
          const td = document.createElement("td"); td.className = "n";
          td.textContent = v; tr.appendChild(td);
        });
        corpo.appendChild(tr);
      });
      $("rolaRetroativo").style.display = "block";

      const nr = $("notaRetroativo");
      const limpo = r.faltantes === 0 && r.indevidos === 0 && r.duplicados === 0;
      nr.className = "nota " + (limpo ? "ok" : "aviso");
      nr.innerHTML =
        r.reconciliadas + " de " + r.rodadas + " rodadas reconciliadas por completo · " +
        r.faltantes + " títulos deixaram de entrar (" + moeda(r.valor_faltante) + ") · " +
        r.indevidos + " entraram sem dever · " +
        r.duplicados + " linhas repetidas (" + moeda(r.valor_duplicado) + ') — <a href="/baixar?token=' +
        r.token + '&nome=retroativo">' + r.arquivo + "</a>";
      addLine("", "info");
      addLine("Consolidado: " + r.caminho, "ok");
      liberar("on", "Reprocessamento concluído.", !bloqueado && !emitido);
    });
    es.addEventListener("erro", ev => {
      es.close();
      const e = JSON.parse(ev.data);
      addLine("❌  " + e.msg, "erro");
      $("notaRetroativo").className = "nota erro";
      $("notaRetroativo").textContent = e.msg;
      liberar("err", "Erro.", !bloqueado && !emitido);
    });
    es.onerror = () => {
      if(ocupado){ es.close(); addLine("❌  Conexão com o servidor perdida.", "erro");
                   liberar("err", "Conexão interrompida.", !bloqueado && !emitido); }
    };
  });

  // ---- relatório de exceções, já disponível depois de conciliar ------------
  btnExcecoes.addEventListener("click", async () => {
    if(ocupado || !token) return;
    ocupar("Montando o relatório de exceções…");
    try{
      const j = await (await fetch("/excecoes?token=" + token)).json();
      if(j.erro){ addLine("❌  " + j.erro, "erro"); liberar("err", "Erro.", !bloqueado && !emitido); return; }
      addLine("Relatório de exceções: " + j.caminho, "ok");
      window.location = "/baixar?token=" + token + "&nome=excecoes";
      liberar("on", "Relatório pronto.", !bloqueado && !emitido);
    }catch(err){
      addLine("❌  " + err, "erro"); liberar("err", "Erro.", !bloqueado && !emitido);
    }
  });
</script>
</body>
</html>
"""


# ============================================================================
#  6. Servidor HTTP local
# ============================================================================
def _json_texto(payload):
    return json.dumps(payload, ensure_ascii=False, default=str)


def _sse(handler, evento, payload):
    """Envia um evento SSE. ``evento=None`` -> evento 'message' padrão."""
    linha = ""
    if evento:
        linha += f"event: {evento}\n"
    linha += "data: " + _json_texto(payload) + "\n\n"
    try:
        handler.wfile.write(linha.encode("utf-8"))
        handler.wfile.flush()
    except (BrokenPipeError, ConnectionResetError):
        raise ClienteDesconectado()


class ClienteDesconectado(Exception):
    """A aba do navegador foi fechada no meio da execução."""


def escolher_arquivo():
    """Seletor de arquivo nativo (tkinter). Devolve o caminho ou ''."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        raiz = tk.Tk()
        raiz.withdraw()
        raiz.attributes("-topmost", True)
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo de dados",
            filetypes=[("Planilhas e CSV", "*.xlsx *.xlsm *.csv"),
                       ("Planilhas Excel", "*.xlsx *.xlsm"),
                       ("CSV", "*.csv"),
                       ("Todos os arquivos", "*.*")])
        raiz.destroy()
        return caminho or ""
    except Exception as erro:                                # noqa: BLE001
        print("Seletor de arquivo indisponível:", erro)
        return ""


def escolher_pasta():
    """Seletor de pasta nativo (tkinter). Devolve o caminho ou ''."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        raiz = tk.Tk()
        raiz.withdraw()
        raiz.attributes("-topmost", True)
        caminho = filedialog.askdirectory(title="Selecione a pasta de destino")
        raiz.destroy()
        return caminho or ""
    except Exception as erro:                                # noqa: BLE001
        print("Seletor de pasta indisponível:", erro)
        return ""


class Handler(BaseHTTPRequestHandler):
    #  ------------------------------------------------------------------
    def log_message(self, *args):
        pass                       # silencia o log padrão do http.server

    def _responder(self, corpo, code=200, ctype="text/html; charset=utf-8",
                   cabecalhos=()):
        dados = corpo.encode("utf-8") if isinstance(corpo, str) else corpo
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(dados)))
        for chave, valor in cabecalhos:
            self.send_header(chave, valor)
        self.end_headers()
        self.wfile.write(dados)

    def _json(self, payload, code=200):
        self._responder(_json_texto(payload), code=code,
                        ctype="application/json; charset=utf-8")

    def _abrir_sse(self):
        """
        Abre um stream de eventos.

        Sem ``Connection: keep-alive`` de propósito: esse cabeçalho faz o
        ``http.server`` segurar o socket depois do último evento, esperando uma
        segunda requisição que não vem. O navegador não sente — o ``EventSource``
        fecha sozinho no evento ``fim`` —, mas a conexão fica pendurada, e
        qualquer cliente que leia até o fim do fluxo trava. O stream termina
        quando a resposta acaba, que é o que ``close_connection`` garante.
        """
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.close_connection = True

    #  ------------------------------------------------------------------
    def do_GET(self):
        partes = urllib.parse.urlparse(self.path)
        rota = partes.path
        q = urllib.parse.parse_qs(partes.query)

        try:
            if rota in ("/", "/index.html"):
                return self._responder(PAGINA_HTML)
            if rota == "/inicial":
                return self._json(estado_inicial())
            if rota == "/janela":
                return self._json(sugerir_janela(q.get("data", [""])[0]))
            if rota == "/procurar":
                tipo = q.get("tipo", ["arquivo"])[0]
                caminho = escolher_pasta() if tipo == "pasta" else escolher_arquivo()
                return self._json({"path": caminho})
            if rota == "/testar-vortx":
                return self._json(testar_vortx())
            if rota == "/testar-grafeno":
                return self._json(testar_grafeno())
            if rota == "/conciliar":
                return self._rota_conciliar(q)
            if rota == "/gerar":
                return self._rota_gerar(q)
            if rota == "/excecoes":
                return self._rota_excecoes(q)
            if rota == "/retroativo":
                return self._rota_retroativo(q)
            if rota == "/baixar":
                return self._rota_baixar(q)
            if rota == "/abrir-pasta":
                return self._rota_abrir_pasta(q)
        except ClienteDesconectado:
            return
        except Exception as erro:                            # noqa: BLE001
            return self._json({"ok": False, "erro": str(erro)}, code=500)

        return self._responder("404", code=404, ctype="text/plain; charset=utf-8")

    #  -- rotas pesadas, com log ao vivo --------------------------------
    def _rota_conciliar(self, q):
        self._abrir_sse()
        if not _TRAVA.acquire(blocking=False):
            return _sse(self, "erro", {"msg": "Já há uma execução em andamento. "
                                              "Espere ela terminar."})
        registrar = Registrador(
            lambda msg, nivel="info": _sse(self, None, {"msg": msg, "nivel": nivel}),
            etapa="CONCILIAÇÃO")
        try:
            cfg = cfg_do_formulario(q)
            sessao = conciliar(cfg, registrar)
            _sse(self, "fim", payload_resultado(sessao))
        except ClienteDesconectado:
            pass
        except Exception as erro:                            # noqa: BLE001
            registrar(f"Falhou: {erro}", "erro")
            _sse(self, "erro", {"msg": str(erro)})
        finally:
            leitura.fechar_cache()
            _TRAVA.release()

    def _rota_gerar(self, q):
        self._abrir_sse()
        if not _TRAVA.acquire(blocking=False):
            return _sse(self, "erro", {"msg": "Já há uma execução em andamento. "
                                              "Espere ela terminar."})
        registrar = Registrador(
            lambda msg, nivel="info": _sse(self, None, {"msg": msg, "nivel": nivel}),
            etapa="GERAÇÃO")
        try:
            sessao = buscar(q.get("token", [""])[0])
            #  Do formulário só interessam as opções de saída: o resto da rodada
            #  é o que já foi conciliado e validado na tela. Revalidar os campos
            #  de fonte aqui faria a emissão falhar por um caminho de arquivo
            #  que o operador mexeu depois — sem nenhum efeito sobre o Termo.
            gerar(sessao, {
                "pasta_saida": (q.get("pasta_saida", [""])[0].strip()
                                or str(config.PASTA_SAIDAS)),
                "copiar_onedrive": q.get("copiar_onedrive", ["0"])[0] == "1",
                "atualizar_historico": q.get("atualizar_historico", ["1"])[0] == "1",
            }, registrar)
            rotulos = {"excel": "Planilha da rodada",
                       "excecoes": "Relatório de exceções",
                       "onedrive": "Cópia no OneDrive"}
            _sse(self, "fim", {"arquivos": [
                {"nome": nome, "rotulo": rotulos.get(nome, nome),
                 "arquivo": Path(caminho).name, "caminho": str(caminho)}
                for nome, caminho in sessao.arquivos.items()
            ]})
        except ClienteDesconectado:
            pass
        except Exception as erro:                            # noqa: BLE001
            registrar(f"Falhou: {erro}", "erro")
            _sse(self, "erro", {"msg": str(erro)})
        finally:
            _TRAVA.release()

    def _rota_retroativo(self, q):
        self._abrir_sse()
        if not _TRAVA.acquire(blocking=False):
            return _sse(self, "erro", {"msg": "Já há uma execução em andamento. "
                                              "Espere ela terminar."})
        registrar = Registrador(
            lambda msg, nivel="info": _sse(self, None, {"msg": msg, "nivel": nivel}),
            etapa="RETROATIVO")
        try:
            apenas = [n for n in (q.get("apenas", [""])[0] or "").replace(",", " ").split()
                      if n.isdigit()]
            resumo = rodar_retroativo({
                "pasta_saida": (q.get("pasta_saida", [""])[0].strip()
                                or str(config.PASTA_SAIDAS)),
                "pasta_onedrive": (q.get("pasta_onedrive", [""])[0].strip()
                                   or str(config.PASTA_ONEDRIVE)),
                "apenas": apenas or None,
            }, registrar)
            #  O download passa pelo mesmo caminho controlado dos demais: uma
            #  sessão própria, com o arquivo registrado nela.
            sessao = guardar(Sessao({"pasta_saida": str(Path(resumo["caminho"]).parent)},
                                    None))
            sessao.registrar_arquivo("retroativo", resumo["caminho"])
            resumo["token"] = sessao.token
            _sse(self, "fim", resumo)
        except ClienteDesconectado:
            pass
        except Exception as erro:                            # noqa: BLE001
            registrar(f"Falhou: {erro}", "erro")
            _sse(self, "erro", {"msg": str(erro)})
        finally:
            leitura.fechar_cache()
            _TRAVA.release()

    def _rota_excecoes(self, q):
        sessao = buscar(q.get("token", [""])[0])
        registrar = Registrador(lambda msg, nivel="info": None, etapa="EXCEÇÕES")
        caminho = gerar_excecoes(sessao, registrar)
        return self._json({"arquivo": caminho.name, "caminho": str(caminho),
                           "bloqueado": sessao.resultado.bloqueado})

    #  -- arquivos gerados ----------------------------------------------
    def _rota_baixar(self, q):
        """
        Entrega um arquivo **desta** sessão.

        Só serve o que a própria execução gravou e registrou em ``arquivos``.
        Caminho vindo da URL nunca é aberto — senão a página viraria um leitor
        de disco para qualquer coisa que rodasse no navegador.
        """
        sessao = buscar(q.get("token", [""])[0])
        nome = q.get("nome", [""])[0]
        caminho = sessao.arquivos.get(nome)
        if caminho is None or not Path(caminho).exists():
            return self._json({"erro": f"Arquivo '{nome}' não foi gerado nesta rodada."},
                              code=404)
        dados = Path(caminho).read_bytes()
        return self._responder(
            dados,
            ctype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            cabecalhos=[("Content-Disposition",
                         f'attachment; filename="{Path(caminho).name}"')])

    def _rota_abrir_pasta(self, q):
        sessao = buscar(q.get("token", [""])[0])
        pasta = Path(sessao.cfg.get("pasta_saida") or config.PASTA_SAIDAS)
        try:
            os.startfile(pasta)                              # noqa: S606 (Windows)
        except (AttributeError, OSError) as erro:
            return self._json({"ok": False, "erro": str(erro)})
        return self._json({"ok": True, "pasta": str(pasta)})


def achar_porta(inicio=8765, fim=8815):
    import socket
    for porta in range(inicio, fim):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", porta)) != 0:
                return porta
    return inicio


def main():
    porta = achar_porta()
    endereco = f"http://127.0.0.1:{porta}"
    servidor = ThreadingHTTPServer(("127.0.0.1", porta), Handler)
    print("=" * 64)
    print("  Recompra FIDC · Leve Saúde — Netz Asset")
    print("=" * 64)
    print(f"  Servidor rodando em: {endereco}")
    print("  (O navegador deve abrir sozinho. Se não abrir, cole o")
    print("   endereço acima no navegador.)")
    if not config.TEMPLATE_PADRAO.exists():
        print()
        print(f"  ATENÇÃO: falta o template em {config.TEMPLATE_PADRAO}")
        print("  Copie o arquivo oficial de recompra para lá antes de conciliar.")
    print("  Para encerrar: feche esta janela ou pressione Ctrl+C.")
    print("=" * 64)
    if not os.environ.get("NETZ_HUB"):
        threading.Timer(0.8, lambda: webbrowser.open(endereco)).start()
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nEncerrando…")
        servidor.shutdown()


if __name__ == "__main__":
    main()
