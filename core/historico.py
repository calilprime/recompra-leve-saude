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
        return self.ids_recomprados_exceto(None)

    def ids_recomprados_exceto(self, numero=None):
        """
        Os ``IdTituloVortx`` já emitidos, **menos os da própria rodada**.

        A exclusão não é conveniência, é o que torna a trava usável. A casa
        reemite a mesma rodada várias vezes até fechar — a 11ª tem v1.1, v2.0,
        v3.0, v4.0, v4.1, v5.0 e v5.1; a 14ª tem v1.0 e v1.1. Se a rodada 14 já
        está registrada e alguém a roda de novo para corrigir ou conferir, sem
        esta exclusão os seus próprios 2.639 títulos voltam como "já
        recomprados", todos são barrados, o Termo vai a zero — e zero é ERRO.
        A rodada ficaria impossível de reemitir.

        O que a trava protege continua protegido: títulos cobrados em **outra**
        rodada seguem barrados. É o caso do 165955407 (Cíntia Farias Cordeiro),
        que apareceu na 11ª e voltou na 13ª. Aquele bloqueio não depende de a
        rodada ser nova; depende de o título ser de outra.

        ``Historico.registrar`` já regrava a rodada de mesmo número em vez de
        acrescentar, então as duas pontas tratam reemissão do mesmo jeito.
        """
        ids = set()
        for rodada in self.rodadas:
            if numero is not None and rodada.get("numero") == int(numero):
                continue
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
        """
        A rodada mais recente — pela **janela**, não pelo número.

        O número não serve de ordem: os arquivos do OneDrive trazem numeração
        conflitante (a rodada de 10/08 está gravada como 13 no nome do arquivo e
        como 12 no Termo assinado). A janela e a data de recompra são
        inequívocas, e é delas que a continuidade depende.
        """
        if not self.rodadas:
            return None
        return max(self.rodadas, key=lambda r: (
            nz.iso(r.get("janela_fim")) or "",
            nz.iso(r.get("data_recompra")) or "",
            r.get("numero") or 0,
        ))

    def proximo_numero(self):
        ultima = self.ultima
        if not ultima:
            return None
        maior = max((r.get("numero") or 0) for r in self.rodadas)
        return maior + 1

    @property
    def numeracao_confiavel(self):
        """
        A numeração das rodadas forma sequência sem buraco nem repetição?

        Quando não forma, a interface **exige** o número em vez de sugerir: um
        palpite errado aqui sai no nome do arquivo e no Termo assinado.
        """
        numeros = [r.get("numero") for r in self.rodadas if r.get("numero")]
        if not numeros or len(numeros) != len(set(numeros)):
            return False
        return sorted(numeros) == list(range(min(numeros), max(numeros) + 1))

    def janela_anterior(self):
        """Início e fim da janela da última rodada, para checar continuidade."""
        ultima = self.ultima
        if not ultima:
            return None, None
        return nz.data(ultima.get("janela_inicio")), nz.data(ultima.get("janela_fim"))

    # -- sobreposição de janela ---------------------------------------------
    def sobreposicoes(self, inicio, fim, ignorar_numero=None):
        """
        Rodadas já registradas cuja janela invade o período ``inicio..fim``.

        É a trava que teria pego o erro da simulação de 14/08 **antes** de ela
        rodar: a janela 16/07–31/07 já havia sido consumida pela rodada de
        10/08, então os títulos já tinham saído da carteira e o Termo só podia
        sair vazio.

        ``ignorar_numero`` deixa de fora a própria rodada — reemitir a 14ª
        corrigida não pode esbarrar na 14ª que já está no histórico.
        """
        inicio, fim = nz.data(inicio), nz.data(fim)
        if not inicio or not fim:
            return []
        achados = []
        for rodada in self.rodadas:
            if (ignorar_numero is not None
                    and rodada.get("numero") == int(ignorar_numero)):
                continue
            r_inicio = nz.data(rodada.get("janela_inicio"))
            r_fim = nz.data(rodada.get("janela_fim"))
            if not r_inicio or not r_fim:
                continue
            if r_inicio <= fim and inicio <= r_fim:
                achados.append(rodada)
        return sorted(achados, key=lambda r: nz.iso(r.get("janela_inicio")) or "")

    def resumo_ultima(self):
        """A linha que a tela mostra ao lado dos campos de janela."""
        ultima = self.ultima
        if not ultima:
            return None
        return {
            "numero": ultima.get("numero"),
            "data_recompra": nz.br(ultima.get("data_recompra")),
            "janela": (f"{nz.br(ultima.get('janela_inicio'))} a "
                       f"{nz.br(ultima.get('janela_fim'))}"),
            "janela_fim": nz.iso(ultima.get("janela_fim")),
            "qtd_titulos": ultima.get("qtd_titulos"),
            "valor_total": str(ultima.get("valor_total") or "0"),
            "arquivo": ultima.get("arquivo") or "",
            "titulos_travados": bool(ultima.get("titulos")),
        }

    def descricao(self, rodada):
        """``rodada 13, janela 01/07/2026 a 31/07/2026, recompra 10/08/2026``."""
        return (f"rodada {rodada.get('numero') or '—'}, janela "
                f"{nz.br(rodada.get('janela_inicio'))} a "
                f"{nz.br(rodada.get('janela_fim'))}, recompra "
                f"{nz.br(rodada.get('data_recompra'))}")

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
                  juros_mora=None, multa=None, extras=None, gravar=True):
        """
        Grava uma rodada emitida. Chamado **depois** de o Termo ser gerado.

        Regrava a rodada se o número já existir — reemitir a mesma rodada
        corrigida não pode inflar o histórico.

        ``extras`` acrescenta campos de procedência (usado pela semente, que
        precisa registrar de onde tirou o número e por que não travou títulos).
        ``gravar=False`` acumula na memória sem tocar no arquivo, para quem vai
        registrar treze rodadas de uma vez.
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
        registro.update(extras or {})
        self.rodadas = [r for r in self.rodadas if r.get("numero") != int(numero)]
        self.rodadas.append(registro)
        self.rodadas.sort(key=lambda r: r.get("numero") or 0)
        if gravar:
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
