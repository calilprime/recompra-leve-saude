# -*- coding: utf-8 -*-
"""
Motor de conciliação Vórtx <-> Grafeno.
=======================================

Três passadas, na ordem da seção 8.2 do plano:

1. **chave exata** — documento + vencimento + valor;
2. **tolerância condicional** — documento + vencimento com diferença de até
   R$ 0,05, e **só** quando houver exatamente uma linha não casada de cada lado
   para aquele par. Em 192 de 227 divergências o sacado tem dois títulos no
   mesmo vencimento; relaxar essa condição pareia o título errado (regra 15.6);
3. **classificação** — o que sobrou não casa, mas ganha uma causa provável.

O sentido é sempre Vórtx -> Grafeno: a base da administradora é a fonte de
verdade do escopo do fundo (decisão da seção 3.2). O caminho inverso vira
número informativo.

Duas coisas que este módulo faz e a planilha não fazia:

* varre a base **inteira** da Grafeno, sem limite de linha (defeito 2.1);
* quando a chave tem mais de um boleto, escolhe pela prioridade de status
  ``Aberta (Vencida) > Aberta > Baixada > Paga``, nunca pelo primeiro da
  ordem da planilha (defeito 2.2).
"""

from collections import defaultdict

from core import config
from core import normalizacao as nz


# ---------------------------------------------------------------------------
#  Resultado
# ---------------------------------------------------------------------------
class Casamento:
    """Um título da Vórtx e o que a conciliação encontrou para ele."""

    __slots__ = ("titulo", "boletos", "boleto", "passada", "causa",
                 "diferenca_valor", "desempate")

    def __init__(self, titulo):
        self.titulo = titulo
        self.boletos = []          # todos os boletos da chave
        self.boleto = None         # o escolhido — ver escolher_boleto
        self.passada = None        # 1, 2 ou None
        self.causa = None          # preenchida na passada 3
        self.diferenca_valor = None
        self.desempate = None      # "numero_titulo" | "prioridade" | None

    @property
    def casou(self):
        return self.boleto is not None

    @property
    def status(self):
        """Status resolvido. ``N/A`` quando não há boleto — como na planilha."""
        return self.boleto.status if self.boleto else "N/A"

    @property
    def qtd_boletos(self):
        return len(self.boletos)

    def __repr__(self):
        return (f"<Casamento {self.titulo.id_titulo} passada={self.passada} "
                f"status={self.status!r} boletos={self.qtd_boletos}>")


class Conciliacao:
    """O resultado completo de uma conciliação."""

    def __init__(self):
        self.casamentos = []
        self.por_id = {}
        self.duplicados_vortx = []      # (titulo_mantido, titulo_descartado)
        self.ids_duplicados = []        # IdTituloVortx repetido após dedup
        self.multiplos_boletos = []     # casamentos com mais de um boleto
        self.boletos_sem_titulo = []
        #  Casamentos em que o Nosso_Número apontou um boleto de status
        #  diferente do que a prioridade escolheria. São os únicos casos em que
        #  o motor e a fórmula da planilha podem discordar por causa do
        #  desempate — a fórmula testa a ligação primeiro, então tem de dar o
        #  mesmo; se não der, a validação de convergência mostra.
        self.desempates_divergentes = []
        #  Medição, nesta base, do que a ligação "403"+NumeroTitulo faria se
        #  fosse promovida a chave. Ver ``_medir_ligacao``.
        self.diagnostico_ligacao = {}
        self.contagens = {}

    @property
    def casados(self):
        return [c for c in self.casamentos if c.casou]

    @property
    def nao_casados(self):
        return [c for c in self.casamentos if not c.casou]

    def por_causa(self, causa):
        return [c for c in self.casamentos if c.causa == causa]

    @property
    def cobertura(self):
        """Fração de títulos da Vórtx com boleto na Grafeno."""
        if not self.casamentos:
            return 0.0
        return len(self.casados) / len(self.casamentos)


# ---------------------------------------------------------------------------
#  Deduplicação da Vórtx
# ---------------------------------------------------------------------------
def deduplicar(titulos, log=None):
    """
    Mantém uma linha por ``IdTituloVortx``, a de ``DataGeracao`` mais recente.

    É a primeira defesa contra o defeito 2.3: a aba VORTX da 12ª tinha duas
    extrações sobrepostas e sete linhas antigas sobrando no rodapé, o que fez
    cinco títulos entrarem duas vezes no Termo.
    """
    log = log or (lambda *a, **k: None)
    melhor = {}
    descartados = []

    for titulo in titulos:
        chave = titulo.id_titulo or f"__linha_{titulo.linha}"
        atual = melhor.get(chave)
        if atual is None:
            melhor[chave] = titulo
            continue
        if _mais_recente(titulo, atual):
            descartados.append((titulo, atual))     # (mantido, descartado)
            melhor[chave] = titulo
        else:
            descartados.append((atual, titulo))

    mantidos = sorted(melhor.values(), key=lambda t: t.linha)
    if descartados:
        log(f"  Deduplicação Vórtx: {len(descartados)} linhas removidas "
            f"({len(titulos)} -> {len(mantidos)}).", "aviso")
    else:
        log(f"  Deduplicação Vórtx: nenhuma linha repetida ({len(mantidos)} títulos).")
    return mantidos, descartados


def _mais_recente(a, b):
    """True se ``a`` deve prevalecer sobre ``b`` na deduplicação."""
    ga, gb = a.data_geracao, b.data_geracao
    if ga and gb and ga != gb:
        return ga > gb
    if ga and not gb:
        return True
    if gb and not ga:
        return False
    return a.linha > b.linha        # empate: fica a linha de baixo


# ---------------------------------------------------------------------------
#  Conciliação
# ---------------------------------------------------------------------------
def conciliar(titulos, boletos, log=None):
    """Roda as três passadas e devolve a :class:`Conciliacao`."""
    log = log or (lambda *a, **k: None)

    resultado = Conciliacao()
    titulos, descartados = deduplicar(titulos, log=log)
    resultado.duplicados_vortx = descartados

    #  Índices da Grafeno. Coluna inteira, sem limite de linha, sempre.
    por_chave = defaultdict(list)
    por_parcial = defaultdict(list)
    por_documento = defaultdict(list)
    por_doc_valor = defaultdict(list)
    for boleto in boletos:
        por_chave[boleto.chave].append(boleto)
        por_parcial[boleto.chave_parcial].append(boleto)
        por_documento[boleto.documento].append(boleto)
        por_doc_valor[f"{boleto.documento}|{nz.valor_texto(boleto.valor)}"].append(boleto)

    # ---- Passada 1 — chave exata -----------------------------------------
    boletos_usados = set()
    for titulo in titulos:
        casamento = Casamento(titulo)
        resultado.casamentos.append(casamento)
        resultado.por_id.setdefault(titulo.id_titulo, []).append(casamento)

        candidatos = por_chave.get(titulo.chave)
        if candidatos:
            casamento.boletos = list(candidatos)
            casamento.boleto = escolher_boleto(candidatos, titulo)
            casamento.passada = 1
            casamento.diferenca_valor = _diferenca(titulo, casamento.boleto)
            if len(candidatos) > 1:
                _registrar_desempate(resultado, casamento, candidatos)
            for b in candidatos:
                boletos_usados.add(id(b))

    casados_1 = sum(1 for c in resultado.casamentos if c.passada == 1)
    log(f"  Passada 1 (chave exata): {casados_1} de {len(titulos)} títulos casados.")

    # ---- Passada 2 — tolerância condicional -------------------------------
    #  Só vale quando sobrou exatamente uma linha de cada lado para o par
    #  (documento, vencimento). É a regra 15.6, que não se relaxa.
    pendentes_por_parcial = defaultdict(list)
    for casamento in resultado.casamentos:
        if not casamento.casou:
            pendentes_por_parcial[casamento.titulo.chave_parcial].append(casamento)

    casados_2 = 0
    bloqueados_2 = 0
    for parcial, pendentes in pendentes_por_parcial.items():
        livres = [b for b in por_parcial.get(parcial, ())
                  if id(b) not in boletos_usados]
        if len(pendentes) != 1 or len(livres) != 1:
            if pendentes and livres:
                bloqueados_2 += len(pendentes)
            continue
        casamento, boleto = pendentes[0], livres[0]
        diferenca = _diferenca(casamento.titulo, boleto)
        if diferenca is None or diferenca > config.TOLERANCIA_VALOR:
            continue
        casamento.boletos = [boleto]
        casamento.boleto = boleto
        casamento.passada = 2
        casamento.diferenca_valor = diferenca
        boletos_usados.add(id(boleto))
        casados_2 += 1

    log(f"  Passada 2 (tolerância de R$ {config.TOLERANCIA_VALOR} em par único): "
        f"{casados_2} títulos casados; {bloqueados_2} recusados por haver mais de "
        f"uma linha de um dos lados.")

    # ---- Passada 3 — classificação, sem casar ------------------------------
    #  Um título é "irmão" quando o mesmo sacado tem outro título no mesmo
    #  vencimento que casou — o padrão de 192 dos 227 casos da seção 3.
    casados_por_parcial = defaultdict(int)
    for casamento in resultado.casamentos:
        if casamento.casou:
            casados_por_parcial[casamento.titulo.chave_parcial] += 1

    for casamento in resultado.casamentos:
        if casamento.casou:
            continue
        titulo = casamento.titulo
        if not por_documento.get(titulo.documento):
            casamento.causa = config.CAUSA_SACADO_INEXISTENTE
        elif por_parcial.get(titulo.chave_parcial):
            if casados_por_parcial.get(titulo.chave_parcial):
                casamento.causa = config.CAUSA_TITULO_IRMAO
            else:
                casamento.causa = config.CAUSA_DIVERGENCIA_VALOR
            vizinhos = por_parcial[titulo.chave_parcial]
            casamento.diferenca_valor = min(
                (_diferenca(titulo, b) for b in vizinhos
                 if _diferenca(titulo, b) is not None),
                default=None,
            )
        elif por_doc_valor.get(
                f"{titulo.documento}|{nz.valor_texto(titulo.valor_nominal)}"):
            casamento.causa = config.CAUSA_DIVERGENCIA_VENCIMENTO
        else:
            casamento.causa = config.CAUSA_SEM_CORRESPONDENCIA

    # ---- Múltiplos boletos, duplicidade e o lado inverso -------------------
    resultado.multiplos_boletos = [c for c in resultado.casamentos
                                   if c.qtd_boletos > 1]
    #  Ids que apareciam mais de uma vez na base bruta. A deduplicação os
    #  resolveu aqui, mas a planilha não deduplica nada: a fórmula ainda vê as
    #  duas linhas. Por isso eles ficam barrados e a validação bloqueia.
    resultado.ids_duplicados = sorted(
        {descartado.id_titulo for _mantido, descartado in descartados
         if descartado.id_titulo}
    )
    resultado.boletos_sem_titulo = [b for b in boletos if id(b) not in boletos_usados]

    resultado.diagnostico_ligacao = _medir_ligacao(titulos, boletos, por_chave)

    resultado.contagens = {
        "titulos_lidos": len(titulos) + len(descartados),
        "titulos": len(titulos),
        "linhas_removidas_dedup": len(descartados),
        "boletos": len(boletos),
        "casados": len(resultado.casados),
        "casados_passada_1": casados_1,
        "casados_passada_2": casados_2,
        "nao_casados": len(resultado.nao_casados),
        "multiplos_boletos": len(resultado.multiplos_boletos),
        "desempate_por_numero": sum(1 for c in resultado.casamentos
                                    if c.desempate == "numero_titulo"),
        "desempate_mudou_status": len(resultado.desempates_divergentes),
        "ids_duplicados": len(resultado.ids_duplicados),
        "boletos_sem_titulo": len(resultado.boletos_sem_titulo),
        "cobertura": resultado.cobertura,
    }
    resultado.contagens.update(
        {f"ligacao_{k}": v for k, v in resultado.diagnostico_ligacao.items()})

    d = resultado.diagnostico_ligacao
    if d.get("pares"):
        log(f"  Diagnóstico da ligação \"{config.PREFIXO_NOSSO_NUMERO}\"+"
            f"NumeroTitulo, se fosse usada como chave: formaria {d['pares']} "
            f"pares, {d['outro_sacado']} deles com boleto de OUTRO sacado "
            f"({d['outro_sacado'] / d['pares'] * 100:.1f}%). "
            f"Encontraria {d['so_ela']} títulos que a chave composta não "
            f"encontra, e {d['so_ela_outro_sacado']} desses "
            f"({d['so_ela_outro_sacado'] / d['so_ela'] * 100:.1f}%) seriam do "
            f"sacado errado." if d.get("so_ela") else
            f"  Diagnóstico da ligação: {d['pares']} pares, "
            f"{d['outro_sacado']} de outro sacado.",
            "aviso" if d.get("so_ela_outro_sacado") else "info")

    log(f"  Conciliação: {len(resultado.casados)} de {len(titulos)} títulos com "
        f"boleto ({resultado.cobertura * 100:.2f}% de cobertura).",
        "ok" if resultado.cobertura > 0.9 else "aviso")
    if resultado.multiplos_boletos:
        pelo_numero = resultado.contagens["desempate_por_numero"]
        log(f"  {len(resultado.multiplos_boletos)} chaves com mais de um boleto — "
            f"{pelo_numero} resolvidas pelo Nosso_Número, "
            f"{len(resultado.multiplos_boletos) - pelo_numero} pela prioridade "
            f"de status.", "aviso")
        if resultado.desempates_divergentes:
            log(f"  Em {len(resultado.desempates_divergentes)} delas o "
                f"Nosso_Número apontou um status diferente do que a prioridade "
                f"escolheria — estão no relatório de exceções.", "aviso")
    return resultado


def _medir_ligacao(titulos, boletos, por_chave):
    """
    Mede, nesta base, o que a ligação ``"403"+NumeroTitulo`` faria como chave.

    Não muda o resultado da conciliação: é diagnóstico, e existe para que a
    decisão de **não** promover a ligação a chave primária seja reconferida a
    cada rodada, em vez de ficar valendo por uma medição de agosto de 2026.

    A pergunta que ela responde é a única que importa: dos pares que a ligação
    formaria, quantos são de outro sacado? Na base da 14ª rodada foram 2.452 de
    56.927 no total (4,3%) — e, entre os títulos que **só** a ligação encontra,
    2.451 de 2.486, ou 98,6%. A ligação acerta o que a chave composta já acerta
    e erra quase tudo que acrescenta.

    Se algum dia estes números virarem zero, a conversa muda e o plano deve ser
    revisto. Enquanto não virarem, a ligação continua servindo só de desempate.
    """
    por_nosso = {}
    for boleto in boletos:
        if boleto.chave_boleto:
            #  Nosso_Número é chave única na Grafeno (medido: 158.918 valores
            #  distintos em 158.918 linhas). Se algum dia repetir, o primeiro
            #  fica — e o diagnóstico segue valendo como ordem de grandeza.
            por_nosso.setdefault(boleto.chave_boleto, boleto)

    pares = outro_sacado = so_ela = so_ela_outro_sacado = 0
    for titulo in titulos:
        if not titulo.chave_boleto:
            continue
        boleto = por_nosso.get(titulo.chave_boleto)
        if boleto is None:
            continue
        pares += 1
        errado = boleto.documento != titulo.documento
        if errado:
            outro_sacado += 1
        if not por_chave.get(titulo.chave):
            so_ela += 1
            if errado:
                so_ela_outro_sacado += 1
    return {"pares": pares, "outro_sacado": outro_sacado, "so_ela": so_ela,
            "so_ela_outro_sacado": so_ela_outro_sacado,
            "boletos_com_ligacao": len(por_nosso)}


def escolher_boleto(candidatos, titulo=None):
    """
    Qual dos boletos da chave é o daquele título.

    Duas regras, nesta ordem:

    **1. O Nosso_Número.** O boleto que a Grafeno emitiu para o título carrega o
    ``NumeroTitulo`` dentro do ``Nosso_Número`` (``382477119`` vira
    ``403382477119``). Quando exatamente um dos candidatos bate assim, é ele —
    não por preferência, mas por identidade.

    **2. A prioridade de status** (decisão 1.2), quando a ligação não resolve:
    ``Aberta (Vencida)`` vence, porque é o que será recomprado. O ``PROCX`` da
    planilha devolvia o primeiro boleto na ordem da aba, o que deixava de fora
    títulos legitimamente vencidos com um ``Baixada`` na frente.

    A regra 1 só escolhe **entre os candidatos que a chave composta já validou**
    por documento, vencimento e valor. Nunca amplia o conjunto: sozinha, a
    ligação pareia gente diferente em 79% dos casos que a chave composta recusa
    (ver ``config.PREFIXO_NOSSO_NUMERO``).
    """
    if titulo is not None and getattr(titulo, "chave_boleto", ""):
        pelo_numero = [b for b in candidatos
                       if b.chave_boleto and b.chave_boleto == titulo.chave_boleto]
        if len(pelo_numero) == 1:
            return pelo_numero[0]
    return min(candidatos, key=lambda b: (b.prioridade, b.linha))


def _registrar_desempate(resultado, casamento, candidatos):
    """
    Anota como o empate foi resolvido — e se isso mudou o status.

    A distinção que importa não é "quem ganhou", é **se o status resolvido
    mudou**. Se o Nosso_Número apontar um boleto de status diferente do que a
    prioridade escolheria, o título pode entrar ou sair do Termo, e isso precisa
    aparecer no relatório de exceções.
    """
    por_prioridade = min(candidatos, key=lambda b: (b.prioridade, b.linha))
    casamento.desempate = ("numero_titulo"
                           if casamento.boleto is not por_prioridade
                           or (casamento.titulo.chave_boleto
                               and casamento.boleto.chave_boleto
                               == casamento.titulo.chave_boleto)
                           else "prioridade")
    if casamento.boleto.status != por_prioridade.status:
        resultado.desempates_divergentes.append((casamento, por_prioridade))


def _diferenca(titulo, boleto):
    if titulo.valor_nominal is None or boleto is None or boleto.valor is None:
        return None
    return abs(titulo.valor_nominal - boleto.valor)


# ---------------------------------------------------------------------------
#  Seleção para o Termo (seção 8.3)
# ---------------------------------------------------------------------------
class Selecao:
    """Títulos escolhidos para o Termo e os que foram barrados, com o motivo."""

    def __init__(self):
        self.elegiveis = []
        self.barrados = []        # (casamento, motivo)

    @property
    def ids(self):
        return [c.titulo.id_titulo for c in self.elegiveis]


def selecionar(conciliacao, janela_inicio, janela_fim, ja_recomprados=(), log=None):
    """
    Um título entra no Termo se, e somente se, as quatro condições valerem:

    1. status resolvido é ``Aberta (Vencida)``;
    2. o vencimento está na janela;
    3. o ``IdTituloVortx`` não consta do histórico de recompras;
    4. o ``IdTituloVortx`` não está duplicado depois da deduplicação.
    """
    log = log or (lambda *a, **k: None)
    ja_recomprados = set(ja_recomprados or ())
    duplicados = set(conciliacao.ids_duplicados)
    selecao = Selecao()

    for casamento in conciliacao.casamentos:
        titulo = casamento.titulo
        vencimento = titulo.vencimento

        if casamento.status != config.STATUS_RECOMPRAVEL:
            continue                      # não é candidato, nem é exceção
        if vencimento is None:
            selecao.barrados.append((casamento, "Sem data de vencimento"))
            continue
        if not (janela_inicio <= vencimento <= janela_fim):
            continue                      # fora da janela: não é barrado, é outra quinzena
        if titulo.id_titulo in ja_recomprados:
            selecao.barrados.append(
                (casamento, "Título já consta do histórico de recompras"))
            continue
        if titulo.id_titulo in duplicados:
            selecao.barrados.append(
                (casamento, "IdTituloVortx duplicado na base após deduplicação"))
            continue
        selecao.elegiveis.append(casamento)

    selecao.elegiveis.sort(key=lambda c: c.titulo.linha)
    log(f"  Seleção: {len(selecao.elegiveis)} títulos elegíveis na janela "
        f"{nz.br(janela_inicio)} a {nz.br(janela_fim)}; "
        f"{len(selecao.barrados)} barrados.",
        "ok" if selecao.elegiveis else "aviso")
    return selecao


# ---------------------------------------------------------------------------
#  Resíduo de rodadas anteriores (validação de aviso da seção 9)
# ---------------------------------------------------------------------------
def residuo_anterior(conciliacao, janela_inicio, ja_recomprados=()):
    """Títulos ``Aberta (Vencida)`` vencidos antes do início da janela."""
    ja_recomprados = set(ja_recomprados or ())
    return [
        c for c in conciliacao.casamentos
        if c.status == config.STATUS_RECOMPRAVEL
        and c.titulo.vencimento is not None
        and c.titulo.vencimento < janela_inicio
        and c.titulo.id_titulo not in ja_recomprados
    ]
