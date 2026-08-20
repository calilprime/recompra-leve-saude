# -*- coding: utf-8 -*-
"""
Fonte Vórtx — estoque de títulos.
=================================

Primária: PostgreSQL alimentado diariamente pela API da Vórtx (seção 7.1 do
plano), **somente SELECT**. Alternativa: arquivo exportado à mão, com o mesmo
layout de colunas.

A consulta traz só a ``DataGeracao`` mais recente. Isso elimina na origem o
defeito 2.3 — a mistura de duas extrações na mesma aba, que cobrou cinco
títulos em duplicidade na 12ª rodada.
"""

import os
from decimal import Decimal
from pathlib import Path

from core import config
from core import leitura
from core import normalizacao as nz


class Titulo:
    """Uma linha do estoque da Vórtx, já normalizada."""

    __slots__ = (
        "linha", "id_titulo", "documento", "nome", "vencimento", "valor_nominal",
        "data_geracao", "data_emissao", "numero_titulo", "campo_chave",
        "numero_boleto", "id_registradora", "situacao", "chave", "chave_parcial",
        "chave_boleto", "bruto",
    )

    def __init__(self, **campos):
        for nome in self.__slots__:
            setattr(self, nome, campos.get(nome))

    @property
    def documento_formatado(self):
        return nz.documento_formatado(self.documento)

    def __repr__(self):
        return (f"<Titulo {self.id_titulo} {self.nome!r} "
                f"venc={self.vencimento} valor={self.valor_nominal}>")


class ResultadoFonte:
    """Dados lidos + a procedência deles, que vai para o snapshot e o log."""

    __slots__ = ("registros", "tabela", "origem", "caminho", "datas_geracao",
                 "linhas_lidas", "avisos", "data_conteudo", "marcos_conteudo")

    def __init__(self, registros, tabela, origem, caminho,
                 datas_geracao=None, avisos=None, data_conteudo=None,
                 marcos_conteudo=None):
        self.registros = registros
        self.tabela = tabela
        self.origem = origem
        self.caminho = str(caminho or "")
        self.datas_geracao = sorted(datas_geracao or [])
        self.linhas_lidas = len(tabela) if tabela is not None else len(registros)
        self.avisos = avisos or []
        #  Até quando os dados vão, deduzido do **conteúdo** e não do arquivo.
        #  A data de modificação mente: o arquivo lido em 14/08 tinha conteúdo
        #  de 06 a 08/08 e a validação de defasagem reportou 1 dia útil quando a
        #  defasagem real era de seis.
        self.data_conteudo = data_conteudo
        self.marcos_conteudo = marcos_conteudo or {}

    @property
    def data_geracao(self):
        """A data de posição da extração. None quando há mais de uma."""
        return self.datas_geracao[-1] if len(self.datas_geracao) == 1 else None


# ---------------------------------------------------------------------------
#  Arquivo (.xlsx exportado do portal, ou aba VORTX de uma rodada anterior)
# ---------------------------------------------------------------------------
def ler_arquivo(caminho, aba=config.ABA_VORTX, log=None):
    """Lê o estoque de um arquivo. Aceita .xlsx (com aba) e .csv."""
    log = log or (lambda *a, **k: None)
    caminho = Path(caminho)
    log(f"  Lendo estoque Vórtx de {caminho.name}…")

    tabela = leitura.ler(
        caminho, aba=aba,
        colunas_sentinela=(config.VORTX_ID,),
        log=log,
    )
    return _montar(tabela, origem="arquivo", caminho=caminho, log=log)


def _montar(tabela, origem, caminho, log=None):
    """Transforma as linhas cruas em objetos :class:`Titulo`."""
    log = log or (lambda *a, **k: None)

    p_id = tabela.posicao(config.VORTX_ID)
    p_doc = tabela.posicao(config.VORTX_DOCUMENTO)
    p_nome = tabela.posicao(config.VORTX_NOME)
    p_venc = tabela.posicao(config.VORTX_VENCIMENTO)
    p_valor = tabela.posicao(config.VORTX_VALOR)
    p_geracao = tabela.posicao(config.VORTX_GERACAO, obrigatoria=False)
    p_emissao = tabela.posicao(config.VORTX_EMISSAO, obrigatoria=False)
    p_numero = tabela.posicao("NumeroTitulo", obrigatoria=False)
    p_campo = tabela.posicao("CampoChave", obrigatoria=False)
    p_boleto = tabela.posicao("NumeroBoleto", obrigatoria=False)
    p_reg = tabela.posicao("IdTituloRegistradora", obrigatoria=False)
    p_situacao = tabela.posicao("Situacao", obrigatoria=False)

    def pega(linha, posicao):
        if posicao is None or posicao >= len(linha):
            return None
        return linha[posicao]

    #  ``bruto`` é montado na ordem exata das colunas da aba VORTX, por nome.
    #  Assim o banco (43 colunas, outra ordem) e o arquivo produzem a mesma
    #  linha na hora de reescrever a aba — e as colunas auxiliares AP..AZ,
    #  que são fórmula, nunca entram como valor.
    posicoes_saida = [tabela.posicao(nome, obrigatoria=False)
                      for nome in config.COLUNAS_VORTX]

    registros = []
    datas = set()
    perdas_texto = 0

    for numero, linha in tabela.linhas:
        nome = nz.texto(pega(linha, p_nome))
        if nz.tem_perda_acentuacao(nome):
            perdas_texto += 1
        geracao = nz.data(pega(linha, p_geracao))
        if geracao:
            datas.add(geracao)
        documento = nz.documento(pega(linha, p_doc))
        vencimento = nz.data(pega(linha, p_venc))
        valor = nz.valor(pega(linha, p_valor))

        numero_titulo = nz.texto(pega(linha, p_numero))
        registros.append(Titulo(
            linha=numero,
            id_titulo=str(pega(linha, p_id) or "").strip(),
            documento=documento,
            nome=nome,
            vencimento=vencimento,
            valor_nominal=valor,
            data_geracao=geracao,
            data_emissao=nz.data(pega(linha, p_emissao)),
            numero_titulo=numero_titulo,
            campo_chave=nz.texto(pega(linha, p_campo)),
            numero_boleto=nz.texto(pega(linha, p_boleto)),
            id_registradora=nz.texto(pega(linha, p_reg)),
            situacao=nz.texto(pega(linha, p_situacao)),
            chave=nz.chave(documento, vencimento, valor),
            chave_parcial=nz.chave_parcial(documento, vencimento),
            chave_boleto=nz.chave_boleto(numero_titulo),
            bruto=tuple(pega(linha, p) for p in posicoes_saida),
        ))

    avisos = []
    if perdas_texto:
        avisos.append(
            f"{perdas_texto} nomes com acentuação corrompida na origem do "
            f"arquivo. Ler do banco resolve (seção 8.1 do plano)."
        )
    if len(datas) > 1:
        avisos.append(
            "Mais de uma DataGeracao na base: "
            + ", ".join(d.strftime("%d/%m/%Y") for d in sorted(datas))
        )

    log(f"    {len(registros)} títulos normalizados; "
        f"DataGeracao: {', '.join(d.strftime('%d/%m/%Y') for d in sorted(datas)) or '—'}")
    #  Do lado da Vórtx o conteúdo tem data explícita: é a própria DataGeracao.
    return ResultadoFonte(
        registros, tabela, origem, caminho, datas, avisos,
        data_conteudo=max(datas) if datas else None,
        marcos_conteudo={"DataGeracao": max(datas)} if datas else {},
    )


# ---------------------------------------------------------------------------
#  Banco de dados (Fase 3)
# ---------------------------------------------------------------------------
CONSULTA_ESTOQUE = """
SELECT *
FROM vortx.unmasked_relatorio_estoque
WHERE "DataGeracao" = (
        SELECT max("DataGeracao") FROM vortx.unmasked_relatorio_estoque
      )
"""

CONSULTA_FUNDOS = """
SELECT DISTINCT "nomeFundo", "cnpjFundo"
FROM vortx.unmasked_relatorio_estoque
ORDER BY 1
"""

CONSULTA_IDENTIFICADORES = """
SELECT
  count(*)                       AS linhas,
  count("NumeroBoleto")          AS num_boleto,
  count("IdTituloRegistradora")  AS id_registradora,
  count("CampoChave")            AS campo_chave,
  count("NumeroTitulo")          AS numero_titulo,
  count("CampoAdicional1")       AS adicional1,
  count("CampoAdicional2")       AS adicional2,
  count("CampoAdicional3")       AS adicional3,
  count("CampoAdicional4")       AS adicional4,
  count("CampoAdicional5")       AS adicional5
FROM vortx.unmasked_relatorio_estoque
WHERE "DataGeracao" = (SELECT max("DataGeracao") FROM vortx.unmasked_relatorio_estoque)
"""

CONSULTA_FUNDOS_POSICAO = """
SELECT "nomeFundo", "cnpjFundo", count(*) AS linhas
FROM vortx.unmasked_relatorio_estoque
WHERE "DataGeracao" = (SELECT max("DataGeracao") FROM vortx.unmasked_relatorio_estoque)
GROUP BY 1, 2
ORDER BY 3 DESC
"""


def credenciais(env=None):
    """Lê as credenciais do ambiente / do .env da raiz do projeto."""
    env = env or os.environ
    valores = dict(_ler_env_arquivo())
    valores.update({k: v for k, v in env.items() if k.startswith("PG")})
    faltando = [c for c in ("PGHOST", "PGDATABASE", "PGUSER", "PGPASSWORD")
                if not valores.get(c)]
    if faltando:
        raise RuntimeError(
            "Credenciais do PostgreSQL incompletas no .env: "
            + ", ".join(faltando)
        )
    return {
        "host": valores["PGHOST"],
        "port": int(valores.get("PGPORT") or 5432),
        "dbname": valores["PGDATABASE"],
        "user": valores["PGUSER"],
        "password": valores["PGPASSWORD"],
        "sslmode": valores.get("PGSSLMODE") or "require",
        "connect_timeout": int(valores.get("PGTIMEOUT") or 20),
    }


def _ler_env_arquivo():
    caminho = config.RAIZ / ".env"
    if not caminho.exists():
        return {}
    valores = {}
    for bruta in caminho.read_text(encoding="utf-8-sig").splitlines():
        bruta = bruta.strip()
        if not bruta or bruta.startswith("#") or "=" not in bruta:
            continue
        chave, _, valor = bruta.partition("=")
        valores[chave.strip()] = valor.strip().strip('"').strip("'")
    return valores


def conectar():
    """Abre a conexão somente leitura com o banco da Vórtx."""
    try:
        import psycopg2
    except ImportError as erro:                       # pragma: no cover
        raise RuntimeError(
            "psycopg2 não instalado. Rode: pip install -r requirements.txt"
        ) from erro
    conexao = psycopg2.connect(**credenciais())
    #  Cinto e suspensório: a sessão inteira fica somente leitura.
    conexao.set_session(readonly=True, autocommit=True)
    return conexao


def testar_conexao(log=None):
    """Confere acesso e devolve a DataGeracao mais recente e a contagem."""
    log = log or (lambda *a, **k: None)
    with conectar() as conexao:
        with conexao.cursor() as cursor:
            cursor.execute(
                'SELECT max("DataGeracao"), count(*) '
                "FROM vortx.unmasked_relatorio_estoque "
                'WHERE "DataGeracao" = (SELECT max("DataGeracao") '
                "FROM vortx.unmasked_relatorio_estoque)"
            )
            geracao, linhas = cursor.fetchone()
    geracao = nz.data(geracao)
    log(f"  Banco Vórtx: DataGeracao mais recente {nz.br(geracao)}, "
        f"{linhas} linhas.", "ok")
    return {"data_geracao": geracao, "linhas": linhas}


def ler_banco(cnpj_fundo=None, log=None):
    """Lê o estoque da Vórtx pela data de geração mais recente."""
    log = log or (lambda *a, **k: None)
    consulta = CONSULTA_ESTOQUE
    parametros = ()
    if cnpj_fundo:
        consulta += '  AND "cnpjFundo" = %s\n'
        parametros = (cnpj_fundo,)

    log("  Conectando ao PostgreSQL da Vórtx (somente leitura)…")
    with conectar() as conexao:
        with conexao.cursor() as cursor:
            cursor.execute(consulta, parametros)
            cabecalho = [d[0] for d in cursor.description]
            linhas = [(i, tuple(linha))
                      for i, linha in enumerate(cursor.fetchall(), start=2)]
    log(f"    {len(linhas)} linhas recebidas do banco.", "ok")

    tabela = leitura.Tabela(cabecalho, linhas, "postgresql://sor", "vortx.unmasked_relatorio_estoque")
    return _montar(tabela, origem="banco", caminho="", log=log)


def verificar_identificadores(log=None, cnpj_fundo=None):
    """
    Tarefa obrigatória da Fase 3 (seção 7.1).

    O export em Excel traz NumeroBoleto, IdTituloRegistradora e CampoAdicional1..5
    vazios. Se no banco vierem preenchidos e casarem com Nosso_Número ou
    Seu_Número da Grafeno, a chave composta deixa de ser necessária.

    A contagem é sempre por fundo: a tabela guarda mais de um, e misturar o NTZ
    IMPULSE aqui daria uma proporção de preenchimento que não é a do Leve Saúde.
    """
    log = log or (lambda *a, **k: None)
    consulta = CONSULTA_IDENTIFICADORES
    parametros = ()
    cnpj_fundo = cnpj_fundo if cnpj_fundo is not None else config.CNPJ_FUNDO
    if cnpj_fundo:
        consulta += '  AND "cnpjFundo" = %s\n'
        parametros = (cnpj_fundo,)

    with conectar() as conexao:
        with conexao.cursor() as cursor:
            cursor.execute(consulta, parametros)
            nomes = [d[0] for d in cursor.description]
            valores = cursor.fetchone()
    resultado = dict(zip(nomes, valores))
    total = resultado.get("linhas") or 0
    for campo, quantidade in resultado.items():
        if campo == "linhas":
            continue
        proporcao = (quantidade / total * 100) if total else 0
        nivel = "ok" if quantidade else "aviso"
        log(f"    {campo}: {quantidade} de {total} preenchidos ({proporcao:.1f}%)",
            nivel)
    return resultado


def listar_fundos(log=None):
    """Descobre o valor de cnpjFundo para gravar em config.py."""
    log = log or (lambda *a, **k: None)
    with conectar() as conexao:
        with conexao.cursor() as cursor:
            cursor.execute(CONSULTA_FUNDOS)
            fundos = cursor.fetchall()
    for nome, cnpj in fundos:
        log(f"    {nome} — {cnpj}")
    return fundos
