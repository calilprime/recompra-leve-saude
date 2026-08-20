# Como usar — Recompra FIDC Leve Saúde

Guia rápido de operação. Para detalhes técnicos e regras completas, veja o [README.md](README.md).

---

## 1. Abrir o programa

Dê **duplo clique em `Abrir Recompra.bat`**.

Uma janela de terminal abre por trás e o navegador abre sozinho em
`http://127.0.0.1:<porta>`. Se o navegador não abrir, copie o endereço
mostrado no terminal e cole manualmente.

**Primeira vez usando?** Antes de tudo, no terminal:

```
pip install -r requirements.txt
```

E confira se existe o arquivo `templates/Leve_Saude_Recompra_TEMPLATE.xlsx`
— sem ele a conciliação não roda.

---

## 2. Preencher a tela

> **Antes da primeira rodada:** clique em **"Popular o histórico pelas rodadas
> anteriores"**, no fim da página. Sem histórico a numeração da rodada sai
> errada (a primeira simulação saiu como `Leve_Saude_1_...` sendo a 14ª), a
> trava contra recomprar o mesmo título não existe e a checagem de janela já
> usada não funciona. Ele só lê o OneDrive; não escreve nada lá.

**Parâmetros da rodada**
- **Nº da recompra**: vem sugerido **só quando a numeração do histórico forma
  sequência**. Hoje não forma — os arquivos de agosto discordam entre si (a
  pasta `13` traz arquivo numerado 13 e Termo assinado numerado 12) — então o
  campo vem **vazio e é obrigatório**. Confira contra o Termo assinado da
  rodada anterior antes de digitar.
- **Data da recompra**: use a data de hoje, do dia da extração. Atrasar custa
  dinheiro (títulos pagos em duplicidade).
- **Início/fim da janela**: o **início** é o dia seguinte ao fim da rodada
  anterior, lido do histórico; o **fim** é a data da recompra menos 3 dias.
  **Não existe regra de dia do mês** — a antiga ("1–15 ou 16 ao fim do mês")
  foi o que zerou o Termo de 14/08, propondo uma janela já recomprada. Logo
  abaixo dos campos aparece o resumo da última rodada registrada: confira que
  a nova janela começa depois de onde aquela terminou.
- **Juros e multa**: só leitura, vêm do template. Para mudar, edite a aba
  INFORMAÇÕES da planilha — nunca aqui.

**Fontes de dados**
- **Vórtx**: escolha `Banco de dados` (recomendado, clique em "Testar conexão"
  para conferir) ou `Arquivo` (aponte o `.xlsx`/`.csv` exportado, com o botão
  "Procurar…").
- **Grafeno**: use `Arquivo` (a API ainda não está liberada). Selecione o
  arquivo de cobranças. **Exporte a Grafeno no dia da recompra**: o sistema
  deduz do próprio conteúdo até quando a base vai (`MAX(Data_Criação)` e
  `MAX(Data_Pagamento)`) e mostra na validação "Frescor da base Grafeno". Se
  você informar uma data de extração mais de um dia útil à frente do que o
  conteúdo mostra, a rodada **para** — foi assim que uma base de 08/08 passou
  por base de 14/08 e trouxe 1.225 boletos vencidos em vez de 2.326. O campo
  "Data da extração da Grafeno" serve só de conferência; quem manda é o
  conteúdo.

**Saída**
- **Pasta de destino**: onde os arquivos da rodada vão ser salvos (padrão:
  pasta `saidas/`).
- **Copiar para o OneDrive**: opcional, nunca sobrescreve arquivo existente.
- **Atualizar histórico de recompras**: deixe ligado — é a trava contra
  recomprar o mesmo título duas vezes.
- **Modo simulação**: só para conferir rodadas antigas. Não use para emitir
  de verdade (o arquivo sai com sufixo `_SIMULACAO` e o histórico não é
  gravado).

---

## 3. Rodar em dois tempos

### 3.1 · Conciliar (▶ Conciliar)

Extrai as duas fontes, concilia, calcula e valida. **Não grava nada
definitivo.** No console (parte de baixo da tela) você acompanha o passo a
passo em tempo real.

Ao terminar, aparece o **painel**:
- **Resumo da rodada**: títulos elegíveis, valor nominal, valor de recompra, exceções.
- **Validações**: lista com selo `OK` / `AVISO` / `ERRO`. Se houver qualquer
  `ERRO`, a emissão fica bloqueada até corrigir e conciliar de novo (ou usar
  modo simulação).
- **Prévia do Termo**: primeiras linhas do que vai para a planilha final.

Aqui já dá para baixar o **relatório de exceções** (botão "⬇ Baixar relatório
de exceções") — é o momento de revisão antes de decidir emitir.

### 3.2 · Gerar Termo e planilha (▶ Gerar Termo e planilha)

Só fica habilitado depois de conciliar sem erro. Grava a planilha final da
rodada, o relatório de exceções e (se marcado) atualiza o histórico.

Ao final aparecem links para baixar os arquivos e para abrir a pasta de saída
direto no Explorer.

**Importante**: depois de gerar, se quiser emitir de novo é preciso clicar em
"Conciliar" novamente — o botão "Gerar" não pode ser reusado na mesma sessão.

---

## 4. Onde ver as saídas

| O quê | Onde |
|---|---|
| Planilha da rodada (Termo) | pasta escolhida em "Pasta de destino" (padrão `saidas/`), nome tipo `Leve_Saude_<nº>_Recompra_...xlsx` |
| Relatório de exceções | mesma pasta, `Excecoes_Recompra_<nº>_<data>.xlsx` |
| Log da execução (linha a linha) | `logs/execucao_AAAA-MM-DD.log` |
| Snapshot de cada extração (cópia bruta das bases) | `snapshots/AAAA-MM-DD_rodada-NN/` |
| Histórico de recompras (trava anti-duplicidade) | `historico_recompras.json`, atualizado sozinho |

No modo simulação, os arquivos saem com sufixo `_SIMULACAO` e o histórico
não é tocado.

---

## 5. Extra: popular o histórico

Bloco **"Histórico de recompras · a trava"**, no fim da tela. Lê as rodadas já
emitidas na pasta do OneDrive e monta o `historico_recompras.json`. Rode uma
vez, antes da primeira rodada de verdade.

O que ele consegue de cada rodada:

| Rodadas | O que entra no histórico |
|---|---|
| 9 em diante | Janela, data, valor **e a lista de títulos** — a trava funciona título a título |
| 4 a 8 | Janela, data e valor. **Sem lista de títulos**: naquele layout a coluna `N°` do Termo repete o mesmo número entre sacados diferentes (1.420 linhas para 883 números na 4ª), então travar por ela barraria o título errado |
| 1 a 3 | Só a data, tirada do nome da pasta — os arquivos não têm aba de Termo |

Ele **não** sobrescreve um histórico que já tenha rodadas, a menos que você
marque a opção. O histórico é a trava contra cobrança em duplicidade.

---

## 6. Extra: reprocessamento retroativo

No fim da mesma tela há o bloco **"Reprocessamento retroativo"**: roda o
motor sobre rodadas já emitidas (pasta do OneDrive) para conferir passivo —
títulos que deveriam ter entrado e não entraram, ou que entraram sem dever.
Só lê, não emite Termo e não mexe no histórico. Gera
`saidas/Passivo_Retroativo_AAAA-MM-DD.xlsx`.

---

## Dúvidas comuns

- **"Emissão bloqueada"**: veja o painel de Validações, corrija o que estiver
  em `ERRO` e clique em Conciliar de novo.
- **"O Termo saiu com ZERO títulos"**: é `ERRO`, sempre. Confira, nesta ordem:
  (1) a janela invade uma rodada já feita? (2) a base da Grafeno é do dia?
  (3) a data da recompra está certa? Foi essa combinação que produziu o Termo
  vazio de 14/08.
- **"A tela informou extração em X, mas o conteúdo vai até Y"**: a base da
  Grafeno é de outra semana. Exporte de novo e rode outra vez.
- **"A janela invade rodada já registrada"**: aquele período já foi recomprado.
  Os títulos já saíram da carteira da Vórtx, então o Termo sairia vazio.
  Comece a janela onde a rodada anterior terminou.
- **Perdeu a sessão ("esta conciliação não está mais na memória")**: o
  servidor reiniciou. Clique em Conciliar de novo antes de Gerar.
- **Fonte "Banco de dados" falhando**: confira o `.env` (copie de
  `.env.exemplo`) e se o IP da máquina está liberado no firewall do Azure.
