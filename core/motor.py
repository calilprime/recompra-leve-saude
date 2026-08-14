# -*- coding: utf-8 -*-
"""
Orquestração de uma rodada de recompra.
=======================================

Fluxo em dois tempos, como a interface (seção 11.2):

``conciliar(cfg, log)``
    extrai, grava snapshot, concilia, calcula, valida. **Não escreve nada de
    definitivo.** Devolve o :class:`Resultado`, que alimenta o painel de
    validações e a prévia do Termo.

``gerar(resultado, cfg, log)``
    só depois, e só se nenhuma validação estiver em ``ERRO``: escreve a planilha
    da rodada, o relatório de exceções e atualiza o histórico.

Essa separação é intencional — é o momento de aprovação de exceções previsto na
decisão 2.2, e impede que um Termo saia antes de alguém olhar o painel.
"""

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from core import calculo, conciliacao, config, excecoes, excel_saida
from core import fonte_grafeno, fonte_vortx, historico as historico_mod
from core import leitura
from core import normalizacao as nz
from core import snapshot, validacoes


class Resultado:
    """Tudo que a conciliação produziu, pronto para a tela e para a emissão."""

    def __init__(self):
        self.parametros = None
        self.vortx = None
        self.grafeno = None
        self.conciliacao = None
        self.selecao = None
        self.linhas_termo = []
        self.total_nominal = Decimal(0)
        self.total_recompra = Decimal(0)
        self.historico = None
        self.validacoes = []
        self.residuo = []
        self.defasagem_dias = 0
        self.defasagem_problema = None
        self.data_extracao_vortx = None
        self.data_extracao_grafeno = None
        self.pasta_snapshot = None
        self.template = None
        self.convergencia = None
        self.modo_simulacao = False
        self.arquivos = {}

    @property
    def bloqueado(self):
        return any(v.nivel == "ERRO" for v in self.validacoes)

    @property
    def erros(self):
        return [v for v in self.validacoes if v.nivel == "ERRO"]

    @property
    def avisos(self):
        return [v for v in self.validacoes if v.nivel == "AVISO"]

    def resumo(self):
        """Os números que a interface mostra no painel."""
        contagens = dict(self.conciliacao.contagens) if self.conciliacao else {}
        return {
            "rodada": self.parametros.numero_rodada if self.parametros else None,
            "data_recompra": nz.br(self.parametros.data_recompra) if self.parametros else "",
            "janela": (f"{nz.br(self.parametros.janela_inicio)} a "
                       f"{nz.br(self.parametros.janela_fim)}") if self.parametros else "",
            "titulos_elegiveis": len(self.linhas_termo),
            "valor_nominal": f"{self.total_nominal:.2f}",
            "valor_recompra": f"{self.total_recompra:.2f}",
            "excecoes": (len(self.conciliacao.nao_casados)
                         + len(self.conciliacao.multiplos_boletos)
                         + len(self.selecao.barrados)) if self.conciliacao else 0,
            "defasagem_dias": self.defasagem_dias,
            #  Defasagem não medida não é defasagem zero. Quando falta a data de
            #  extração de um dos lados, ``defasagem_dias`` fica 0 por não haver
            #  o que subtrair — e a tela não pode mostrar isso em verde como se
            #  a base estivesse no dia.
            "defasagem_medida": self.defasagem_problema is None,
            "defasagem_frase": (self.defasagem_problema
                                or calculo.frase_defasagem(self.defasagem_dias)),
            "contagens": contagens,
            "bloqueado": self.bloqueado,
            "modo_simulacao": self.modo_simulacao,
        }


# ---------------------------------------------------------------------------
#  Tempo 1 — conciliar
# ---------------------------------------------------------------------------
def conciliar(cfg, log=None):
    """Extrai, concilia, calcula e valida. Não grava nada de definitivo."""
    log = log or (lambda *a, **k: None)
    resultado = Resultado()
    resultado.modo_simulacao = bool(cfg.get("modo_simulacao"))

    # -- template e parâmetros ---------------------------------------------
    template = Path(cfg.get("template") or config.TEMPLATE_PADRAO)
    if not template.exists():
        raise FileNotFoundError(
            f"Template não encontrado: {template}\n"
            f"Coloque uma cópia do arquivo oficial de recompra em "
            f"{config.PASTA_TEMPLATES} com o nome "
            f"{config.TEMPLATE_PADRAO.name}."
        )
    resultado.template = template
    log(f"Template: {template.name}")

    parametros = calculo.ler_parametros(template, log=log)
    parametros = _aplicar_cfg(parametros, cfg, log=log)
    resultado.parametros = parametros

    hist = historico_mod.Historico(cfg.get("historico"))
    resultado.historico = hist
    ja_recomprados = hist.ids_recomprados
    if ja_recomprados:
        log(f"  Histórico: {len(hist.rodadas)} rodadas, "
            f"{len(ja_recomprados)} títulos já recomprados.")

    # -- extração ----------------------------------------------------------
    log("")
    log("── Extração ──────────────────────────────────────────", "destaque")
    try:
        resultado.vortx = _extrair_vortx(cfg, log=log)
        resultado.grafeno = _extrair_grafeno(cfg, log=log)
    finally:
        leitura.fechar_cache()

    for aviso in resultado.vortx.avisos + resultado.grafeno.avisos:
        log(f"  ⚠ {aviso}", "aviso")

    resultado.data_extracao_vortx = resultado.vortx.data_geracao
    resultado.data_extracao_grafeno = _data_extracao_grafeno(cfg, resultado, log=log)
    resultado.defasagem_dias = _defasagem(resultado, log=log)

    # -- snapshot (regra 15.7: antes de qualquer processamento) -------------
    if cfg.get("gravar_snapshot", True):
        resultado.pasta_snapshot = snapshot.gravar(
            parametros.numero_rodada, resultado.vortx, resultado.grafeno,
            parametros=parametros, log=log,
        )

    # -- conciliação --------------------------------------------------------
    log("")
    log("── Conciliação ───────────────────────────────────────", "destaque")
    resultado.conciliacao = conciliacao.conciliar(
        resultado.vortx.registros, resultado.grafeno.registros, log=log)

    resultado.selecao = conciliacao.selecionar(
        resultado.conciliacao, parametros.janela_inicio, parametros.janela_fim,
        ja_recomprados=ja_recomprados, log=log)

    resultado.residuo = conciliacao.residuo_anterior(
        resultado.conciliacao, parametros.janela_inicio, ja_recomprados)

    # -- cálculo ------------------------------------------------------------
    log("")
    log("── Cálculo ───────────────────────────────────────────", "destaque")
    _montar_termo(resultado, log=log)

    # -- validações ---------------------------------------------------------
    log("")
    log("── Validações ────────────────────────────────────────", "destaque")
    resultado.validacoes = validacoes.rodar(resultado, log=log)
    return resultado


def _aplicar_cfg(parametros, cfg, log=None):
    """Parâmetros da tela vencem os do template; juros e multa, nunca."""
    log = log or (lambda *a, **k: None)
    if cfg.get("numero_rodada"):
        parametros.numero_rodada = int(cfg["numero_rodada"])
    if cfg.get("data_recompra"):
        parametros.data_recompra = nz.data(cfg["data_recompra"])
    if cfg.get("janela_inicio"):
        parametros.janela_inicio = nz.data(cfg["janela_inicio"])
    if cfg.get("janela_fim"):
        parametros.janela_fim = nz.data(cfg["janela_fim"])

    if parametros.data_recompra and not (parametros.janela_inicio and parametros.janela_fim):
        parametros.janela_inicio, parametros.janela_fim = \
            calculo.sugerir_janela(parametros.data_recompra)
        log(f"  Janela sugerida pela regra quinzenal: "
            f"{nz.br(parametros.janela_inicio)} a {nz.br(parametros.janela_fim)}.")

    log(f"  Rodada {parametros.numero_rodada} · recompra em "
        f"{nz.br(parametros.data_recompra)} · janela "
        f"{nz.br(parametros.janela_inicio)} a {nz.br(parametros.janela_fim)}.")
    return parametros


def _extrair_vortx(cfg, log=None):
    fonte = (cfg.get("fonte_vortx") or "arquivo").lower()
    if fonte == "banco":
        return fonte_vortx.ler_banco(cfg.get("cnpj_fundo") or config.CNPJ_FUNDO, log=log)
    caminho = cfg.get("arquivo_vortx")
    if not caminho:
        raise ValueError("Informe o arquivo da Vórtx (ou escolha a fonte 'Banco').")
    return fonte_vortx.ler_arquivo(
        caminho, aba=cfg.get("aba_vortx", config.ABA_VORTX), log=log)


def _extrair_grafeno(cfg, log=None):
    fonte = (cfg.get("fonte_grafeno") or "arquivo").lower()
    if fonte == "api":
        return fonte_grafeno.ler_api(log=log)
    caminho = cfg.get("arquivo_grafeno")
    if not caminho:
        raise ValueError("Informe o arquivo da Grafeno.")
    return fonte_grafeno.ler_arquivo(
        caminho, aba=cfg.get("aba_grafeno", config.ABA_GRAFENO), log=log)


def _data_extracao_grafeno(cfg, resultado, log=None):
    """
    A data da extração da Grafeno.

    O export não traz essa data em coluna nenhuma. A interface pede ao operador;
    na falta, vale a data de modificação do arquivo, e o log diz qual foi usada.
    """
    log = log or (lambda *a, **k: None)
    informada = nz.data(cfg.get("data_extracao_grafeno"))
    if informada:
        log(f"  Extração Grafeno: {nz.br(informada)} (informada na tela).")
        return informada
    caminho = resultado.grafeno.caminho
    if caminho and Path(caminho).exists():
        data_arquivo = datetime.fromtimestamp(Path(caminho).stat().st_mtime).date()
        log(f"  Extração Grafeno: {nz.br(data_arquivo)} "
            f"(data de modificação do arquivo — confirme na tela).", "aviso")
        return data_arquivo
    return date.today()


def _defasagem(resultado, log=None):
    """
    Dias úteis entre a extração mais antiga e a data de recompra.

    Três situações que não são "defasagem zero" e não podem passar por isso:
    data de recompra em branco, data de extração desconhecida (é o que acontece
    quando a base tem mais de uma ``DataGeracao``) e extração **posterior** à
    recompra, que é impossível numa rodada real.
    """
    log = log or (lambda *a, **k: None)
    resultado.defasagem_problema = None
    recompra = resultado.parametros.data_recompra
    if not recompra:
        resultado.defasagem_problema = "Data de recompra não informada."
        return 0

    faltando = [rotulo for rotulo, valor in
                (("Vórtx", resultado.data_extracao_vortx),
                 ("Grafeno", resultado.data_extracao_grafeno)) if not valor]
    if faltando:
        resultado.defasagem_problema = (
            "Data de extração desconhecida em " + " e ".join(faltando)
            + " — sem ela não dá para medir a defasagem."
        )
        log(f"  {resultado.defasagem_problema}", "aviso")
        return 0

    posteriores = [rotulo for rotulo, valor in
                   (("Vórtx", resultado.data_extracao_vortx),
                    ("Grafeno", resultado.data_extracao_grafeno))
                   if valor > recompra]
    if posteriores:
        resultado.defasagem_problema = (
            "Extração de " + " e ".join(posteriores) + " é posterior à data de "
            "recompra — confira as datas informadas."
        )
        log(f"  {resultado.defasagem_problema}", "aviso")
        return 0

    dias = max(calculo.dias_uteis(d, recompra)
               for d in (resultado.data_extracao_vortx,
                         resultado.data_extracao_grafeno))
    log(f"  {calculo.frase_defasagem(dias)}", "ok" if dias == 0 else "aviso")
    return dias


def _montar_termo(resultado, log=None):
    """Monta as linhas do Termo e soma os totais."""
    log = log or (lambda *a, **k: None)
    parametros = resultado.parametros
    linhas = []
    total_nominal = Decimal(0)
    total_recompra = Decimal(0)

    for casamento in resultado.selecao.elegiveis:
        titulo = casamento.titulo
        conta = calculo.calcular(
            titulo.valor_nominal, titulo.vencimento, parametros.data_recompra,
            parametros.juros_mora, parametros.multa,
        )
        total_nominal += conta.valor_nominal
        total_recompra += conta.total
        linhas.append({
            "linha_vortx": titulo.linha,
            "nome": (titulo.nome or "").upper(),
            "documento": titulo.documento,
            "documento_formatado": titulo.documento_formatado,
            "id_titulo": titulo.id_titulo,
            "vencimento": titulo.vencimento,
            "data_recompra": parametros.data_recompra,
            "valor_nominal": conta.valor_nominal,
            "valor_recompra": conta.total,
            "prazo": conta.prazo,
            "multa": conta.multa,
            "juros": conta.juros,
            "status": casamento.status,
            "qtd_boletos": casamento.qtd_boletos,
            "passada": casamento.passada,
        })

    resultado.linhas_termo = linhas
    resultado.total_nominal = total_nominal
    resultado.total_recompra = total_recompra
    log(f"  Termo: {len(linhas)} títulos · nominal R$ {total_nominal:,.2f} · "
        f"recompra R$ {total_recompra:,.2f}"
        .replace(",", "@").replace(".", ",").replace("@", "."), "ok")


# ---------------------------------------------------------------------------
#  Tempo 2 — gerar
# ---------------------------------------------------------------------------
def gerar(resultado, cfg, log=None):
    """
    Escreve a planilha da rodada, o relatório de exceções e o histórico.

    Regra 15.5: nenhum Termo sai com validação em ``ERRO``.
    """
    log = log or (lambda *a, **k: None)
    if resultado.bloqueado and not resultado.modo_simulacao:
        raise RuntimeError(
            "Emissão bloqueada: "
            + "; ".join(v.titulo for v in resultado.erros)
        )

    destino = Path(cfg.get("pasta_saida") or config.PASTA_SAIDAS)
    destino.mkdir(parents=True, exist_ok=True)

    log("")
    log("── Geração ───────────────────────────────────────────", "destaque")
    caminho_excel = excel_saida.gerar(resultado, destino, log=log)
    resultado.arquivos["excel"] = caminho_excel

    caminho_excecoes = excecoes.gerar(resultado, destino, log=log)
    resultado.arquivos["excecoes"] = caminho_excecoes

    if cfg.get("atualizar_historico", True) and not resultado.modo_simulacao:
        registro = resultado.historico.registrar(
            numero=resultado.parametros.numero_rodada,
            data_recompra=resultado.parametros.data_recompra,
            janela_inicio=resultado.parametros.janela_inicio,
            janela_fim=resultado.parametros.janela_fim,
            titulos=[l["id_titulo"] for l in resultado.linhas_termo],
            valor_total=resultado.total_recompra,
            arquivo_saida=caminho_excel.name,
            juros_mora=resultado.parametros.juros_mora,
            multa=resultado.parametros.multa,
        )
        log(f"  Histórico atualizado: rodada {registro['numero']} com "
            f"{registro['qtd_titulos']} títulos.", "ok")
    elif resultado.modo_simulacao:
        log("  Modo simulação: o histórico NÃO foi atualizado.", "aviso")

    if cfg.get("copiar_onedrive"):
        copia = _copiar_para_onedrive(caminho_excel, cfg, log=log)
        if copia:
            resultado.arquivos["onedrive"] = copia

    return resultado.arquivos


def _copiar_para_onedrive(caminho, cfg, log=None):
    """
    Cópia para a pasta do OneDrive — só quando o operador marca a caixa.

    Regra 15.2: nada é sobrescrito lá sem confirmação explícita. Se já existir
    arquivo com o mesmo nome, a cópia é recusada e o operador decide.
    """
    log = log or (lambda *a, **k: None)
    import shutil
    pasta = Path(cfg.get("pasta_onedrive") or config.PASTA_ONEDRIVE)
    if not pasta.exists():
        log(f"  Pasta do OneDrive não encontrada: {pasta}", "aviso")
        return None
    alvo = pasta / caminho.name
    if alvo.exists():
        log(f"  Já existe {alvo.name} no OneDrive — cópia recusada para não "
            f"sobrescrever (regra 15.2).", "aviso")
        return None
    shutil.copy2(caminho, alvo)
    log(f"  Cópia gravada no OneDrive: {alvo}", "ok")
    return alvo
