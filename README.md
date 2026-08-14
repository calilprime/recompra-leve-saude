# Recompra FIDC · Leve Saúde

Automação da recompra quinzenal dos títulos vencidos da carteira cedida ao FIDC.
Concilia o estoque da **Vórtx** com os boletos da **Grafeno**, apura o Termo de
Recompra e gera a planilha da rodada no layout que a casa já usa.

A especificação completa está em `PLANO_AUTOMACAO_RECOMPRA.md`, na pasta acima.
Este arquivo é o manual de operação e o registro do que já está pronto.

---

## Como rodar

**Duplo clique em `Abrir Recompra.bat`.** O navegador abre sozinho.

Na primeira vez, instale as dependências:

```
pip install -r requirements.txt
```

E confira se `templates/Leve_Saude_Recompra_TEMPLATE.xlsx` existe. Ele é a cópia
do arquivo oficial de recompra com as duas fórmulas corrigidas — a automação
copia esse arquivo a cada rodada e preenche os dados.

---

## O fluxo, em dois tempos

A tela tem dois botões, e a ordem é intencional.

**1 · Conciliar** — extrai, concilia, calcula e valida. **Não grava nada de
definitivo.** Ao final aparecem o painel de validações, o resumo da rodada e a
prévia do Termo. O relatório de exceções já pode ser baixado aqui: é o momento
de aprovação previsto na decisão 2.2.

**2 · Gerar Termo e planilha** — só aparece depois de conciliar, e fica
desabilitado se qualquer validação estiver em `ERRO` (regra 15.5).

Entre um e outro está a conferência humana. Nenhum Termo sai sem alguém ter
olhado o painel.

### Preenchendo a tela

| Campo | O que informar |
|---|---|
| Nº da recompra | Sugerido a partir do histórico |
| Data da recompra | **Extraia no dia.** Cada dia útil de defasagem custa ~20 títulos e ~R$ 22 mil |
| Janela | Quinzena estrita (1–15 ou 16–fim do mês). Muda sozinha com a data; pré-preenchida para emendar com a rodada anterior |
| Juros e multa | Somente leitura, sempre da aba INFORMAÇÕES do template. Para mudar a taxa, edite a planilha — nunca o código |
| Fonte Vórtx | `Banco de dados` (recomendado) ou `Arquivo` |
| Fonte Grafeno | `Arquivo` enquanto a API não sai |
| Data da extração da Grafeno | O export não traz essa data. Em branco, vale a data de modificação do arquivo — é por ela que a defasagem é medida |

**Modo simulação** existe para conferir rodadas antigas: o que bloquearia vira
aviso, o arquivo sai com sufixo `_SIMULACAO` e o histórico não é tocado. Não use
para emitir de verdade.

---

## O que a automação corrige

Os quatro defeitos apurados na 11ª e na 12ª recompra, e as travas que impedem a
volta de cada um:

| Defeito | Custo apurado | Trava |
|---|---|---|
| Fórmula lia só até a linha 125.000 | 5 títulos não recomprados, R$ 10.367,84 | `CONT.SES` de coluna inteira; o motor não usa fórmula nenhuma |
| Boleto escolhido pela posição na planilha | 2 títulos não recomprados | Prioridade `Aberta (Vencida)` > `Aberta` > `Baixada` > `Paga` |
| Duas extrações misturadas na aba VORTX | 5 títulos em duplicidade, R$ 9.789,13 | Deduplicação por `IdTituloVortx` + validação de extração única |
| Extração de sexta, recompra na segunda | 20 títulos pagos duas vezes, R$ 25.409,69 | Validação de defasagem zero, com o custo estimado na tela |

Mais a trava contra recompra repetida (`historico_recompras.json`), que existe
por causa do título 165955407 — cobrado três vezes entre a 11ª e a 12ª.

---

## Arquitetura

```
Código VS CODE/
├── Abrir Recompra.bat        # duplo clique
├── recompra_app.py           # servidor local + página + orquestração
├── requirements.txt
├── .env.exemplo              # modelo; o .env real não vai para o Git
│
├── core/
│   ├── config.py             # caminhos, colunas, prioridade de status
│   ├── normalizacao.py       # documento, data, valor, chave
│   ├── leitura.py            # .xlsx e .csv, com detecção de cabeçalho
│   ├── fonte_vortx.py        # PostgreSQL (somente SELECT) + arquivo
│   ├── fonte_grafeno.py      # arquivo; API na Fase 5
│   ├── conciliacao.py        # deduplicação e as três passadas de casamento
│   ├── calculo.py            # multa, juros, janela quinzenal, dias úteis
│   ├── historico.py          # trava contra recompra repetida
│   ├── validacoes.py         # bateria; bloqueia ou avisa
│   ├── excel_saida.py        # preenche o template, sem alterar a estrutura
│   ├── excecoes.py           # relatório de exceções
│   ├── snapshot.py           # toda extração vira snapshot
│   ├── retroativo.py         # reprocessamento das rodadas anteriores
│   └── motor.py              # orquestra os dois tempos
│
├── templates/                # o arquivo oficial, com as fórmulas corrigidas
├── snapshots/                # AAAA-MM-DD_rodada-NN/ + meta.json
├── saidas/                   # planilha e relatório de cada rodada
├── logs/                     # execucao_AAAA-MM-DD.log
└── testes/
    ├── test_conciliacao.py
    └── amostras/             # arquivos reais da 11ª e da 12ª
```

### Por que o motor não confia na planilha

O Termo é calculado em Python. As fórmulas continuam no arquivo entregue, como
conferência independente — e a validação **Convergência Python ↔ Excel** compara
os dois caminhos. Se divergirem, a emissão trava: o arquivo leva as fórmulas
dentro e recalcula quando alguém abre, então o que sai da casa tem de ser
exatamente o que o motor apurou.

### A estrutura do arquivo é congelada

A planilha alimenta os BIs da Netz. Não se cria, exclui, move, renomeia ou
oculta coluna; não se cria, exclui ou renomeia aba; não se mexe em cabeçalho,
formatação ou formato de número. Por isso `excel_saida.py` grava direto no XML
de dentro do `.xlsx`, em vez de reconstruir a planilha — o que o `openpyxl` não
conhece seria perdido em silêncio.

**Se parecer necessário mudar qualquer outra coisa no arquivo, pare e pergunte.**

---

## Testes

```
py testes/test_conciliacao.py            # motor, tudo
py testes/test_conciliacao.py --rapido   # só os testes de unidade, < 1s
py testes/test_interface.py              # interface, de ponta a ponta
```

`test_interface.py` sobe o mesmo servidor do duplo clique e roda a 12ª por HTTP,
sem navegador: concilia, confere o painel, baixa o relatório de exceções e gera
a planilha. Prova, entre outras coisas, que a conciliação não escreve nada, que
o download só entrega o que aquela rodada gerou e que o modo simulação não toca
no histórico.

O aceite roda sobre os arquivos reais em `testes/amostras/` e leva alguns
minutos — são dois arquivos de 50 MB. Ele prova que o motor encontra:

* na **11ª**, os 5 títulos que deveriam ter sido recomprados e não foram
  (Maria Izabel Basilio, Aline de Oliveira Gomes, Ana Teresa Valls Pereira,
  Irisdalva Teles de Deus, Ana Maria Silva Damasceno), R$ 10.367,84 de nominal,
  correspondentes a 7 linhas de boleto na Grafeno — a reconciliação do "7 × 5";
* na **12ª**, as 7 linhas órfãs da extração de 30/07 e os 5 títulos que o Termo
  cobrou em duplicidade.

---

## Regras que não podem ser violadas

1. **Nenhuma escrita nas fontes.** O código lê do banco e da API. Nunca emite,
   cancela, baixa ou liquida boleto.
2. **OneDrive é somente leitura.** A cópia só acontece com a caixa marcada, e
   nunca sobrescreve arquivo existente.
3. **Banco de dados é somente `SELECT`.** A sessão é aberta em modo readonly.
4. **Juros e multa vêm sempre da aba INFORMAÇÕES.** Nenhum valor fixo no código.
5. **Credenciais em `.env`**, nunca no código, nunca em planilha.
6. **Nenhum Termo é emitido com validação em `ERRO`.**
7. **Tolerância de valor só quando o par é único dos dois lados.** Em 192 de 227
   casos o sacado tem dois títulos no mesmo vencimento — relaxar isso pareia o
   título errado.
8. **Toda extração gera snapshot** antes de qualquer processamento.
9. **Nada de endpoint, parâmetro ou coluna inventado.** Pesquisar ou
   inspecionar; na dúvida, registrar pendência e usar o fallback.

---

## Estado das fases

| Fase | O que é | Situação |
|---|---|---|
| 1 · Motor com arquivo | Normalização, conciliação, cálculo, validações, Excel, exceções | **Pronta** — aceite da 11ª e da 12ª passando |
| 2 · Interface | Servidor local, página, SSE, seletores, dois tempos, `.bat` | **Pronta** |
| 3 · Banco da Vórtx | Conexão, consulta pela DataGeracao mais recente, snapshot | **Pronta** — aceite verificado em 12/08/2026 (ver abaixo) |
| 3-A · Conferência pós-recompra | `pos_recompra.py` e o botão na tela | Pendente |
| 4 · Reprocessamento retroativo | `retroativo.py` sobre as rodadas do OneDrive | **Pronta** — 12 rodadas varridas em 13/08/2026 (ver abaixo) |
| 5 · API da Grafeno | Extração via API de Cobranças | Bloqueada: conta solicitada em 11/08 (pendência 14.2-3) |

### Passivo apurado no reprocessamento retroativo (13/08/2026)

Botão **Reprocessar rodadas anteriores**, na interface, ou:

```
py -c "from core import retroativo; retroativo.rodar(log=print)"
```

Saída: `saidas/Passivo_Retroativo_AAAA-MM-DD.xlsx`.

**O que cada rodada permitiu apurar.** A estrutura da planilha mudou várias
vezes ao longo de 2026, e isso limita o alcance:

| Rodadas | Abas | O que dá para apurar |
|---|---|---|
| 9 a 12 | `INFORMAÇÕES` · `TERMO` · `GRAFENO` · `VORTX` | Reconciliação completa |
| 4 a 8 | `INFORMAÇÕES` · `TERMO` · `ESTOQUE` | Só duplicidade dentro do Termo |
| 1 a 3 | `Cobranças` · `RECOMPRA` · `ESTOQUE` | Nada — não há aba de Termo |

**Títulos que deveriam ter sido recomprados e não foram — 7, R$ 13.239,45:**

| Rodada | Título | Sacado | Recompra |
|---|---|---|---|
| 9ª | 160984863 | PALOMA FERNANDEZ DE MELLO E SOUZA | R$ 1.936,05 |
| 9ª | 160983559 | CHRISTIANE COELHO DE NASCIMENTO | R$ 669,89 |
| 11ª | 165959502 | MARIA IZABEL QUEIROZ DA SILVA BASILIO | R$ 1.379,83 |
| 11ª | 165960217 | ALINE DE OLIVEIRA GOMES | R$ 1.656,54 |
| 11ª | 165961375 | ANA TERESA VALLS PEREIRA | R$ 2.018,57 |
| 11ª | 143950451 | IRISDALVA TELES DE DEUS | R$ 2.532,88 |
| 11ª | 143950045 | ANA MARIA SILVA DAMASCENO | R$ 3.045,69 |

Os cinco da 11ª somam os R$ 10.367,84 de valor nominal já conhecidos — e a
lista agora é nominal e fechada, o que **encerra a divergência de contagem
5 × 6 × 7** do adendo 17. Os **dois da 9ª são achado novo**, que ninguém tinha
apurado.

**Nenhuma linha repetida em nenhum Termo emitido.** Nem nas rodadas 4 a 8, nem
na 12ª: a v1.5, que foi a assinada, tem 680 linhas para 680 títulos. As cinco
duplicidades de R$ 9.789,13 existiram só na v1.1 e foram corrigidas antes da
assinatura.

**Títulos que entraram sem que o motor os trouxesse — 31.** Aqui a causa importa
mais que o número, e nem toda causa é defeito:

| Rodada | Qtd | Causa | Leitura |
|---|---|---|---|
| 9ª | 12 | boleto já `Paga` | **Pagamento em duplicidade — R$ 20.985,45** |
| 9ª | 9 | sem boleto na Grafeno do arquivo | A conferir |
| 10ª | 4 | vencimento fora da janela | Provável título arrastado de rodada anterior |
| 10ª | 4 | sem boleto na Grafeno do arquivo | A conferir |
| 12ª | 1 | já constava do histórico | **Cíntia Farias Cordeiro (165955407)** |
| 12ª | 1 | não existe na base VORTX do arquivo | Edição manual |

Três merecem destaque:

**Os 12 da 9ª.** Os boletos foram pagos pelos sacados entre 29/05 e 15/06 —
todos **antes** da extração de 19/06 que está no arquivo, e antes da recompra de
22/06. Ou seja: já constavam como pagos quando a base foi extraída, e mesmo
assim entraram no Termo. O fundo recebeu duas vezes: R$ 20.985,45 de recompra
contra R$ 20.601,40 pagos pelos próprios sacados. **Não é a janela cega da
seção 2.4** — ali o pagamento é posterior à extração; aqui é anterior. Precisa
de conferência com a Operações, porque o arquivo não guarda a extração da
Grafeno que gerou aquele Termo.

**A Cíntia, confirmada.** O título 165955407 aparece no Termo da 12ª barrado por
"já consta do histórico de recompras" — exatamente o caso descrito na pergunta
16.2, agora reproduzido pela automação. A trava funciona.

**A edição manual da 12ª, confirmada.** O título de MARLENE GOMES DE OLIVEIRA
JORGE aparece no Termo com o identificador `"151533082"` — **entre aspas**, e
sem correspondência na aba VORTX do próprio arquivo. É o caso que o adendo 17
levantou (680 no Termo, 679 na base), e as aspas dizem que foi digitado à mão.

---

### Aceite da Fase 3 — extração do banco × arquivo exportado

O banco guarda a posição **diária**. A comparação honesta usa a mesma data dos
dois lados: `DataGeracao` de 03/08/2026 no banco contra o subconjunto de 03/08
da aba VORTX do arquivo da 12ª.

| Verificação | Resultado |
|---|---|
| Contagem | 62.799 no banco · 62.799 no arquivo |
| Títulos só de um lado | 0 no banco · 0 no arquivo |
| Documento, vencimento e valor | 0 divergências em 62.799 |
| Acentuação | 48 nomes corrompidos no arquivo · **0 no banco** |

```
arquivo: 'JANETE Sï¿½ RIBEIRO'      banco: 'JANETE SÁ RIBEIRO'
```

Duas consequências práticas:

* **A fonte "Banco de dados" resolve a acentuação**, como o plano previa. Pelo
  arquivo, 48 nomes vão corrompidos para o Termo.
* **O histórico diário serve à Fase 4.** Dá para reconstruir a base da Vórtx de
  qualquer rodada passada pela `DataGeracao`, sem depender do arquivo do
  OneDrive. As posições disponíveis vão de 23/07/2026 em diante.

E um alerta que o aceite trouxe: a tabela guarda **mais de um fundo**. Na
posição de 12/08 são 57.750 linhas do Leve Saúde e 156 do NTZ IMPULSE
(`62.122.004/0001-08`) — sacados pessoa jurídica, valores na casa dos R$ 200
mil. O filtro por `cnpjFundo` não é opcional, e está em `config.CNPJ_FUNDO`.

### Verificação de identificadores (Fase 3) — feita em 12/08/2026

A pergunta era se algum identificador da Vórtx casa com `Nosso_Número` ou
`Seu_Número` da Grafeno, o que tornaria a chave composta desnecessária.

```
py -c "from core import fonte_vortx; fonte_vortx.verificar_identificadores(print)"
```

**Preenchimento no banco** (posição de 12/08/2026, só Leve Saúde, 57.750 linhas):

| Campo | Preenchido |
|---|---|
| `NumeroTitulo` | **100%** |
| `CampoChave` | **100%** |
| `NumeroBoleto` | 0% |
| `IdTituloRegistradora` · `Registradora` · `IdContratoRegistradora` | 0% |
| `CampoAdicional1..5` | 0% |

Os campos vazios estão vazios **na fonte**, não é limitação do relatório
exportado. E `NumeroTitulo` já vinha 100% preenchido no export também — quem
está vazio lá é `NumeroBoleto`.

**Existe uma correspondência, mas ela não substitui a chave composta.**

`CampoChave` é um sequencial interno da Vórtx (`106706.0`) e não casa com nada
da Grafeno. Já o `NumeroTitulo` casa por prefixo:

```
NumeroTitulo 382477119  →  Nosso_Número 403382477119
```

Sobre o arquivo da 12ª (as duas bases da mesma safra, 62.806 títulos):

| | Ligação direta | Chave composta |
|---|---|---|
| Cobertura | 99,84% | 96,03% |
| Títulos apontando para mais de um boleto | 0 | 21 |

O ganho aparente de 3,8 pontos **é falso**. Os 2.493 títulos que só a ligação
direta encontra foram conferidos pelo nome do sacado, contra um grupo de
controle dos títulos em que as duas chaves concordam:

| | Nome bate | Nome diferente |
|---|---|---|
| Controle — as duas chaves concordam | 99,6% | 0,4% (variação de cadastro: `MARLETE`/`MARLETTE`) |
| Os 2.493 que só a direta acha | 20,4% | **79,1% — outra pessoa** |

```
título: 'REJANE BARBOSA DE SOUZA'   doc 98230980772
boleto: 'ESTER DELFINO DA SILVA'    doc 49088467749
```

A causa é a densidade da numeração: na faixa 382655220–382818919 os dois lados
são sequenciais, com distância mais comum de 8 entre vizinhos. Quando o título
**não tem** boleto, `"403" + NumeroTitulo` cai no boleto de outra pessoa. Não é
a faixa que está quebrada — 12.874 títulos dela casam pelas duas chaves.

**Conclusão: a chave composta e a passada 2 continuam necessárias.** A pendência
14.2-9 está encerrada — não há identificador comum a pedir à Vórtx, porque o
campo que existiria (`NumeroBoleto`) está vazio na origem.

O `NumeroTitulo` tem, ainda assim, um uso legítimo, e ele **foi implementado**:
**desempatar**, e só entre os boletos que a chave composta já validou por
documento, vencimento e valor. Nos 21 casos de mais de um boleto, é ele que diz
qual é o boleto daquele título, em vez da prioridade de status.

```
título 143950045  ANA MARIA SILVA DAMASCENO  venc 04/07  R$ 2.966,61
  a chave composta acha 2 boletos:
     Baixada          nosso 403381697077
     Aberta (Vencida) nosso 403381630061
  NumeroTitulo 381630061 → 403381630061 → é o segundo, sem precisar de prioridade
```

Ordem em `conciliacao.escolher_boleto`: primeiro a ligação, depois a prioridade
de status da decisão 1.2, que continua valendo quando a ligação não resolve. A
ligação **nunca amplia** o conjunto de candidatos — é isso que a mantém segura.

Sobre a 12ª, isso não muda nenhuma linha do Termo. É redução de risco latente,
não correção de número: o que ela elimina é a chance de recomprar o título do
irmão em vez do certo.

**No arquivo, a mudança cabe inteira em `VORTX!AR`** — a mesma fórmula que o
plano já autoriza corrigir. A coluna auxiliar `AP` e a `GRAFENO!AE` continuam
como estavam, então continuam sendo duas fórmulas alteradas, não três. A
fórmula testa a ligação primeiro e a prioridade depois, na mesma ordem do
motor, e a validação de convergência confere que os dois caminhos concordam.

O relatório de exceções ganhou duas colunas na aba `MULTIPLOS_BOLETOS`:
**Resolvido por** e **Só pela prioridade seria**. Quando as duas divergem, a
linha merece conferência.
