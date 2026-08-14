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

**Parâmetros da rodada**
- **Nº da recompra**: já vem sugerido pelo histórico. Normalmente não precisa mexer.
- **Data da recompra**: use a data de hoje, do dia da extração. Atrasar custa
  dinheiro (títulos pagos em duplicidade).
- **Início/fim da janela**: preenche sozinho a partir da data (quinzena 1–15
  ou 16–fim do mês) e já vem ajustado para emendar com a rodada anterior.
- **Juros e multa**: só leitura, vêm do template. Para mudar, edite a aba
  INFORMAÇÕES da planilha — nunca aqui.

**Fontes de dados**
- **Vórtx**: escolha `Banco de dados` (recomendado, clique em "Testar conexão"
  para conferir) ou `Arquivo` (aponte o `.xlsx`/`.csv` exportado, com o botão
  "Procurar…").
- **Grafeno**: use `Arquivo` (a API ainda não está liberada). Selecione o
  arquivo de cobranças. Se souber a data exata da extração, preencha "Data da
  extração da Grafeno" — em branco, o sistema usa a data de modificação do
  arquivo.

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

## 5. Extra: reprocessamento retroativo

No fim da mesma tela há o bloco **"Reprocessamento retroativo"**: roda o
motor sobre rodadas já emitidas (pasta do OneDrive) para conferir passivo —
títulos que deveriam ter entrado e não entraram, ou que entraram sem dever.
Só lê, não emite Termo e não mexe no histórico. Gera
`saidas/Passivo_Retroativo_AAAA-MM-DD.xlsx`.

---

## Dúvidas comuns

- **"Emissão bloqueada"**: veja o painel de Validações, corrija o que estiver
  em `ERRO` e clique em Conciliar de novo.
- **Perdeu a sessão ("esta conciliação não está mais na memória")**: o
  servidor reiniciou. Clique em Conciliar de novo antes de Gerar.
- **Fonte "Banco de dados" falhando**: confira o `.env` (copie de
  `.env.exemplo`) e se o IP da máquina está liberado no firewall do Azure.
