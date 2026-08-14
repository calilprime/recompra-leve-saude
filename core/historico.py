# -*- coding: utf-8 -*-
"""
Histórico de recompras — a trava contra recompra repetida.
==========================================================

Decisão 1.6: trava própria, em arquivo versionado, **não** em aba da planilha.
O motivo está na 11ª rodada: o título 165955407 (Cíntia Farias Cordeiro) foi
recomprado, o boleto não foi baixado na Grafeno, e ele voltou a entrar na 12ª —
duas vezes. Três cobranças no total, duas indevidas.

O arquivo é ``historico_recompras.json`` na raiz do projeto. Uma entrada por
rodada, com a lista de ``IdTituloVortx`` emitidos. Contém CPF? Não: só id de
título, valor e datas. Mesmo assim fica fora do Git junto com os snapshots.
"""

import json
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from core import config
from core import normalizacao as nz


class Historico:
    """Leitura e escrita do histórico de rodadas já emitidas."""

    def __init__(self, caminho=None):
        self.caminho = Path(caminho or config.ARQUIVO_HISTORICO)
        self.rodadas = []
        self.carregar()

    # -- leitura ------------------------------------------------------------
    def carregar(self):
        if not self.caminho.exists():
            self.rodadas = []
            return self.rodadas
        try:
            dados = json.loads(self.caminho.read_text(encoding="utf-8"))
        except json.JSONDecodeError as erro:
            raise RuntimeError(
                f"O histórico {self.caminho.name} está corrompido ({erro}). "
                f"Corrija ou renomeie o arquivo antes de emitir uma nova rodada."
            ) from erro
        self.rodadas = dados.get("rodadas", []) if isinstance(dados, dict) else dados
        return self.rodadas

    @property
    def ids_recomprados(self):
        """Todos os IdTituloVortx já emitidos, de todas as rodadas."""
        ids = set()
        for rodada in self.rodadas:
            ids.update(str(i) for i in rodada.get("titulos", []))
        return ids

    def rodada_de(self, id_titulo):
        """Em que rodada aquele título já foi recomprado."""
        id_titulo = str(id_titulo)
        for rodada in self.rodadas:
            if id_titulo in {str(i) for i in rodada.get("titulos", [])}:
                return rodada
        return None

    @property
    def ultima(self):
        if not self.rodadas:
            return None
        return max(self.rodadas, key=lambda r: (r.get("numero") or 0))

    def proximo_numero(self):
        ultima = self.ultima
        return (ultima.get("numero") or 0) + 1 if ultima else 1

    def janela_anterior(self):
        """Início e fim da janela da última rodada, para checar continuidade."""
        ultima = self.ultima
        if not ultima:
            return None, None
        return nz.data(ultima.get("janela_inicio")), nz.data(ultima.get("janela_fim"))

    def taxas_anteriores(self):
        """
        Juros e multa da última rodada emitida, ou ``(None, None)``.

        A interface compara com o que está na aba INFORMAÇÕES do template e
        avisa se a taxa mudou (seção 11.1). Rodadas gravadas antes deste campo
        existir simplesmente não têm a informação — e aí não há o que comparar.
        """
        ultima = self.ultima
        if not ultima:
            return None, None

        def taxa(valor):
            if valor in (None, ""):
                return None
            try:
                return Decimal(str(valor))
            except ArithmeticError:
                return None

        return taxa(ultima.get("juros_mora")), taxa(ultima.get("multa"))

    # -- escrita ------------------------------------------------------------
    def registrar(self, numero, data_recompra, janela_inicio, janela_fim,
                  titulos, valor_total, arquivo_saida="", observacao="",
                  juros_mora=None, multa=None):
        """
        Grava uma rodada emitida. Chamado **depois** de o Termo ser gerado.

        Regrava a rodada se o número já existir — reemitir a mesma rodada
        corrigida não pode inflar o histórico.
        """
        registro = {
            "numero": int(numero),
            "data_recompra": nz.iso(data_recompra),
            "janela_inicio": nz.iso(janela_inicio),
            "janela_fim": nz.iso(janela_fim),
            "qtd_titulos": len(titulos),
            "valor_total": str(Decimal(valor_total)) if valor_total is not None else "0",
            "juros_mora": str(juros_mora) if juros_mora is not None else "",
            "multa": str(multa) if multa is not None else "",
            "arquivo": str(arquivo_saida),
            "emitido_em": datetime.now().isoformat(timespec="seconds"),
            "observacao": observacao,
            "titulos": [str(t) for t in titulos],
        }
        self.rodadas = [r for r in self.rodadas if r.get("numero") != int(numero)]
        self.rodadas.append(registro)
        self.rodadas.sort(key=lambda r: r.get("numero") or 0)
        self.gravar()
        return registro

    def gravar(self):
        self.caminho.parent.mkdir(parents=True, exist_ok=True)
        conteudo = {
            "fundo": config.NOME_FUNDO,
            "atualizado_em": datetime.now().isoformat(timespec="seconds"),
            "rodadas": self.rodadas,
        }
        temporario = self.caminho.with_suffix(".json.tmp")
        temporario.write_text(
            json.dumps(conteudo, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temporario.replace(self.caminho)

    # -- semente ------------------------------------------------------------
    def semear_de_termo(self, numero, data_recompra, janela_inicio, janela_fim,
                        ids, valor_total=0, observacao=""):
        """
        Alimenta o histórico a partir de um Termo antigo já emitido.

        Usado no reprocessamento retroativo, para que a trava conheça as
        rodadas anteriores à automação.
        """
        return self.registrar(
            numero, data_recompra, janela_inicio, janela_fim, ids,
            valor_total, observacao=observacao or "importado do Termo emitido",
        )
