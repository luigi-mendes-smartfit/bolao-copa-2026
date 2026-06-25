# 🏆 Bolão Copa do Mundo 2026 — Smart Fit

Dashboard ao vivo do bolão interno da Copa do Mundo 2026 entre colaboradores da Smart Fit. Página única, sem servidor, que lê os palpites e resultados direto de uma planilha Google e atualiza sozinha.

### 🔗 [Acessar Dashboard](https://luigi-mendes-smartfit.github.io/bolao-copa-2026/)

---

## ⚽ Sobre

Acompanhamento em tempo real do bolão. Os dados saem de uma planilha Google (alimentada por um Google Forms) e aparecem no site sem ninguém precisar mexer em nada durante a Copa. Tudo roda no navegador, hospedado de graça no GitHub Pages.

## ✨ Funcionalidades

- 🏅 **Ranking ao vivo** — pontuação atualizada e ordenada, com critérios de desempate. **Clique em qualquer participante** para abrir o detalhamento dele: total, posição, de onde vieram os pontos (jogos × grupos), e item a item o que cravou, acertou parcialmente ou errou.
- ⚽ **Jogos do Brasil** — placares e, ao clicar no jogo, os palpites de todo mundo agrupados por resultado, com destaque 🎯 para quem **cravou o placar exato** e marcação do grupo que acertou o resultado.
- 🌍 **Grupos** — cada grupo abre em duas abas: **Distribuição** (em quem o pessoal mais votou) e **Quem votou** (o palpite de cada pessoa e quantos pontos ele vale agora, ao vivo).
- 🏆 **Chaveamento** — o mata-mata oficial de 2026 (16-avos → final). Conforme os grupos vão sendo definidos, os confrontos se preenchem sozinhos; o que ainda não está definido mostra os candidatos no hover.
- 📊 **Estatísticas** — palpites por jogo, favoritos por grupo, **Mestres da Cravada** e curiosidades (participantes, média de gols, premiação total).
- 🎬 **Cerimônia de encerramento** — animação de pódio em tela cheia com troféu, confete dourado e revelação de 3º → 2º → 1º (ver seção [Encerramento](#-encerramento)).
- 🔄 **Auto-refresh** a cada 5 minutos, com indicador de última atualização e estado de erro quando a planilha está fora do ar.

---

## 📋 Regras de pontuação

### Jogos do Brasil (3 jogos)

| Resultado | Pontos |
|-----------|--------|
| Placar exato | **5 pts** |
| Acertou vencedor/empate (placar errado) | **3 pts** |
| Errou quem ganha | **0 pts** |

### Classificação dos grupos (12 grupos)

| Resultado | Pontos |
|-----------|--------|
| 1º e 2º na ordem correta | **5 pts** |
| Os dois certos, ordem trocada | **3 pts** |
| Só 1 acerto na posição correta | **2 pts** |
| Só 1 acerto na posição errada | **1 pt** |
| Errou ambos | **0 pts** |
| ⚠️ Mesmo país em 1º **e** 2º | **ZERA o grupo** |

> A trava do "mesmo país zera" existe porque o Google Forms não impede escolher o mesmo time nas duas perguntas do grupo — a regra evita que alguém ganhe vantagem com isso.

**Pontuação máxima: 75 pts** — 15 dos jogos (3 × 5) + 60 dos grupos (12 × 5).

### Critérios de desempate

1. Maior número de placares exatos dos jogos do Brasil
2. Quem cravou mais posições de grupo (5 pts)
3. Quem enviou o formulário primeiro (timestamp mais antigo)

### Premiação

| Colocação | Fatia |
|-----------|-------|
| 🥇 1º | 50% |
| 🥈 2º | 30% |
| 🥉 3º | 20% |

Arrecadação = nº de participantes × R$ 20,00.

---

## 🏗️ Arquitetura

O segredo do "tempo real sem servidor" é uma planilha **pública** que espelha uma planilha **privada**, e o site lê só a pública:

```
Google Forms ──> Planilha MASTER (privada)            Planilha PÚBLICA (espelho)
                 ├─ Respostas (palpites + e-mails)     ├─ Estatisticas  ─┐
                 ├─ Gabarito (resultados reais)   ==>  ├─ Gabarito       ├─> index.html (GitHub Pages)
                 ├─ Pontuação (fórmulas)               └─ Ranking       ─┘     lê via gviz/JSON
                 ├─ Ranking
                 └─ Controle Pagamento
```

- A **master** é privada (tem e-mails e dados sensíveis) e roda toda a lógica de pontuação em Apps Script.
- A **pública** é "qualquer pessoa com link: leitor" e só contém nome + palpites + resultados, trazidos por `QUERY(IMPORTRANGE(...))`. Sem e-mail, sem edição.
- O **site** consome a pública pelo endpoint **gviz** (`/gviz/tq?...`), que devolve JSON sem precisar de API key.

---

## 📂 Estrutura do repositório

```
.
├── index.html                 # o dashboard inteiro (HTML + CSS + JS, single-file)
├── apps-script/
│   ├── master.gs              # Apps Script da planilha privada (pontuação, gabarito, API)
│   └── publica.gs             # Apps Script da planilha pública (espelhos via IMPORTRANGE)
└── README.md
```

> Os arquivos `.gs` ficam aqui só como referência/backup — eles rodam dentro do editor de Apps Script de cada planilha (Extensões → Apps Script), não a partir do GitHub.

---

## 🚀 Configuração

### 1. Planilha master (privada)

1. Crie o Google Forms e vincule a uma planilha de respostas.
2. Extensões → Apps Script → cole o `apps-script/master.gs`.
3. Rode `configurarPlanilha()` (cria Gabarito, Pontuação, Ranking, Regras, Controle Pagamento, Estatisticas).
4. (Opcional, resultados automáticos) Pegue uma chave em [football-data.org](https://www.football-data.org/), cole em `SETUP_API_KEY()`, rode, e depois `criarGatilhoGabarito()` para atualizar o Gabarito a cada 2h.
5. Rode `criarGatilho()` (recalcula a cada nova resposta) e, se quiser, `protegerAbas()`.

### 2. Planilha pública (espelho)

1. Crie uma planilha nova, compartilhada como **"Qualquer pessoa com o link: Leitor"**.
2. Extensões → Apps Script → cole o `apps-script/publica.gs`.
3. Preencha `MASTER_ID` com o ID da master (o trecho da URL entre `/d/` e `/edit`).
4. Rode `configurarPublica()`.
5. Abra cada aba e clique em **"Permitir acesso"** no aviso do `IMPORTRANGE` (autorização única; o script não consegue clicar por você).
6. Confira com `diagnosticarPublica()`.

### 3. Dashboard

1. No `index.html`, preencha `SPREADSHEET_ID` com o ID da planilha **pública**.
2. Publique no GitHub Pages (Settings → Pages → branch `main`).
3. Pronto — o site atualiza sozinho a cada 5 min.

---

## 🛠️ Operação durante a Copa

- **Resultados:** se a API estiver configurada, o Gabarito da master se preenche sozinho. Sem API, basta digitar os placares e os 1º/2º de cada grupo na aba **Gabarito** da master — a pontuação recalcula em cascata.
- **Pagamento:** a aba **Controle Pagamento** marca como "Eliminado" quem não pagou até o prazo (gatilho diário).
- **Nada de editar resposta enviada:** vale o que entrou no Forms, igual para todos. O sistema só espelha o registro.

---

## 🎬 Encerramento

A cerimônia de pódio é controlada por uma flag no topo do `<script>` do `index.html`:

```js
const ENCERRAMENTO = { ativo: false };
```

Vire para `{ ativo: true }` **somente quando**: todos os jogos finalizados, ranking congelado e premiação confirmada. A partir daí:

- a cerimônia abre **sozinha na primeira visita** de cada pessoa (fluxo: abertura → 3º → 2º → 1º → pódio → estatísticas → botões **Ver Ranking Final** / **Fechar**);
- a visualização fica salva em `localStorage` (`bolao2026_ceremony_seen`), então não reaparece;
- surge o botão **"🏆 Ver cerimônia de encerramento"** no ranking para rever quando quiser;
- tem **"Pular animação"** no canto superior direito.

**Testar sem ligar a flag para todo mundo:**

```js
openCeremony()                                   // dispara na hora (console do navegador)
localStorage.removeItem('bolao2026_ceremony_seen'); location.reload()   // simula "primeira visita"
```

---

## 📑 Contrato de dados (para manutenção)

O site espera estes layouts nas abas da planilha **pública**:

- **Estatisticas** — *sem cabeçalho, dados a partir da linha 1*:
  `A`=Nome · `B`=Unidade (ignorada) · `C`=J1 Brasil · `D`=J1 Marrocos · `E`=J2 Brasil · `F`=J2 Haiti · `G`=J3 Escócia · `H`=J3 Brasil · `I/J`=Grupo A (1º/2º) … `AE/AF`=Grupo L (1º/2º).
- **Ranking** — `A`=# · `B`=Nome · `C`=Total · `D`=Placares exatos · `E`=Posições cravadas · `F`=Timestamp.
- **Gabarito** — espelho 1:1 da master; o site localiza as linhas por **rótulo** ("Brasil x Marrocos", "Grupo C"…), então não quebra se mudar de linha.

> ⚠️ No jogo 3 o **mandante é a Escócia** ("Escócia x Brasil"). No Gabarito, a coluna do 1º time é a Escócia e a do 2º é o Brasil.

---

## 🧰 Tecnologias

- HTML, CSS e JavaScript puro (single-page, self-contained, sem build)
- Google Sheets via endpoint **gviz** (JSON, sem API key)
- Google Apps Script (pontuação na master, espelhos na pública)
- [football-data.org](https://www.football-data.org/) (opcional, resultados automáticos)
- [canvas-confetti](https://github.com/catdad/canvas-confetti) (cerimônia) via CDN
- GitHub Pages (hospedagem gratuita)

---

## 🩺 Solução de problemas

| Sintoma | Causa provável | Correção |
|---------|----------------|----------|
| Falta um participante / ranking parece a menos | A aba Estatisticas não tem cabeçalho; faixa começando em `A2` pula o 1º | Ler a partir de `A1` (já corrigido no `index.html`) |
| Cravada não conta / 🎯 não aparece | O gviz devolve coluna mista (texto+número) como **string**, e `"1" === 1` é falso | Coagir placares com `toNum()` nos dois lados (já corrigido) |
| Grupos não atualizam no Gabarito | A API devolve `"Group C"`, não `"GROUP_C"` | Normalizar a letra do grupo na master (já corrigido) |
| Palpite na Turquia não pontua | Master gravava `Turquía` (com acento); respostas têm `Turquia` | Alinhar a grafia na master |
| Espelho com `#REF!` | `IMPORTRANGE` sem autorização | Abrir a aba e clicar em "Permitir acesso" |
| Site não carrega nada | Planilha não está como "qualquer pessoa com link: leitor", ou `SPREADSHEET_ID` errado | Conferir o compartilhamento e o ID (deve ser o da **pública**) |
| Confete não aparece | Sem internet (a lib vem de CDN) ou `prefers-reduced-motion` ligado | Esperado offline; o resto da cerimônia funciona |

---

## 🔒 Privacidade

A planilha pública **não** expõe e-mails nem qualquer dado além de nome, palpites e pontuação. A master, com os dados sensíveis, permanece privada.

## 📄 Licença

Projeto interno Smart Fit — uso pessoal, sem fins comerciais.

---

<div align="center">

**SMART FIT** · Bolão Copa do Mundo 2026 · Feito com ⚽ e 💛

</div>
