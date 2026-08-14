# -*- coding: utf-8 -*-
"""
Snapshots das extrações (seção 7.3).
====================================

Toda extração, de qualquer fonte, é gravada antes de qualquer processamento
(regra 15.7). É o que permite reconstruir uma rodada meses depois e é a base do
reprocessamento retroativo.

``snapshots/AAAA-MM-DD_rodada-NN/``
    ``vortx.parquet``   — ou ``vortx.csv.gz`` quando o pyarrow não estiver instalado
    ``grafeno.parquet``
    ``meta.json``       — origem, data de geração, contagem de linhas e hash

A pasta contém CPF e nome de pessoa física: entra no ``.gitignore``.
"""

import csv
import gzip
import hashlib
import json
from datetime import date, datetime
from pathlib import Path

from core import config


def pasta_da_rodada(numero, data_extracao=None, raiz=None):
    data_extracao = data_extracao or date.today()
    raiz = Path(raiz or config.PASTA_SNAPSHOTS)
    return raiz / f"{data_extracao.isoformat()}_rodada-{int(numero or 0):02d}"


def gravar(numero, vortx, grafeno, parametros=None, data_extracao=None,
           raiz=None, log=None):
    """
    Grava o par de extrações e o ``meta.json`` da rodada.

    ``vortx`` e ``grafeno`` são :class:`core.fonte_vortx.ResultadoFonte`.
    """
    log = log or (lambda *a, **k: None)
    pasta = pasta_da_rodada(numero, data_extracao, raiz)
    pasta.mkdir(parents=True, exist_ok=True)

    meta = {
        "rodada": numero,
        "gravado_em": datetime.now().isoformat(timespec="seconds"),
        "fundo": config.NOME_FUNDO,
        "fontes": {},
    }
    if parametros is not None:
        meta["parametros"] = {
            "data_recompra": _texto_data(getattr(parametros, "data_recompra", None)),
            "janela_inicio": _texto_data(getattr(parametros, "janela_inicio", None)),
            "janela_fim": _texto_data(getattr(parametros, "janela_fim", None)),
            "juros_mora": str(getattr(parametros, "juros_mora", "")),
            "multa": str(getattr(parametros, "multa", "")),
        }

    for nome, fonte in (("vortx", vortx), ("grafeno", grafeno)):
        if fonte is None:
            continue
        arquivo, formato = _gravar_tabela(pasta / nome, fonte.tabela)
        meta["fontes"][nome] = {
            "origem": fonte.origem,
            "caminho": fonte.caminho,
            "arquivo": arquivo.name,
            "formato": formato,
            "linhas": len(fonte.tabela),
            "colunas": len(fonte.tabela.cabecalho),
            "hash_sha256": _hash(arquivo),
            "datas_geracao": [_texto_data(d) for d in fonte.datas_geracao],
            "avisos": list(fonte.avisos),
        }
        log(f"  Snapshot {nome}: {len(fonte.tabela)} linhas em {arquivo.name}.")

    (pasta / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log(f"  Snapshot gravado em {pasta}.", "ok")
    return pasta


def _gravar_tabela(base, tabela):
    """Grava em parquet quando houver pyarrow; senão, em csv.gz."""
    try:
        import pyarrow            # noqa: F401
        import pyarrow.parquet as pq
        import pyarrow as pa
    except ImportError:
        return _gravar_csv_gz(base.with_suffix(".csv.gz"), tabela), "csv.gz"

    colunas = {}
    for posicao, nome in enumerate(tabela.cabecalho):
        rotulo = str(nome or f"coluna_{posicao + 1}")
        while rotulo in colunas:
            rotulo += "_"
        colunas[rotulo] = [
            _texto(linha[posicao]) if posicao < len(linha) else None
            for _, linha in tabela.linhas
        ]
    colunas["_linha_planilha"] = [str(numero) for numero, _ in tabela.linhas]

    caminho = base.with_suffix(".parquet")
    pq.write_table(pa.table(colunas), caminho, compression="snappy")
    return caminho, "parquet"


def _gravar_csv_gz(caminho, tabela):
    with gzip.open(caminho, "wt", encoding="utf-8", newline="") as saida:
        escritor = csv.writer(saida, delimiter=";")
        escritor.writerow(list(tabela.cabecalho) + ["_linha_planilha"])
        for numero, linha in tabela.linhas:
            escritor.writerow([_texto(v) for v in linha] + [numero])
    return caminho


def _texto(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    return str(v)


def _texto_data(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    return str(v)


def _hash(caminho):
    digest = hashlib.sha256()
    with open(caminho, "rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(1 << 20), b""):
            digest.update(bloco)
    return digest.hexdigest()


def listar(raiz=None):
    """Snapshots existentes, do mais recente para o mais antigo."""
    raiz = Path(raiz or config.PASTA_SNAPSHOTS)
    if not raiz.exists():
        return []
    pastas = []
    for pasta in sorted(raiz.iterdir(), reverse=True):
        meta = pasta / "meta.json"
        if pasta.is_dir() and meta.exists():
            try:
                pastas.append((pasta, json.loads(meta.read_text(encoding="utf-8"))))
            except json.JSONDecodeError:
                continue
    return pastas
