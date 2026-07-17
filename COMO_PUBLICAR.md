# Site Highpar — Monitoramento de CRIs (portal multi-emissão)

O portal é `index.html` + `dados.json` (o banco de dados, versionado junto).
Home = portfólio: um card por emissão com farol; clique no card (ou no menu
lateral) para abrir a lâmina completa da emissão.

## Rotina de atualização (1 comando)

Na pasta `Padrao Consolidado CRI`:

```
python atualizar_dados.py
```

Isso varre os `Consolidado - *.xlsx` da pasta (qualquer emissão no padrão,
qualquer estrutura de séries/métricas/farol) e regrava `site/dados.json`.
O identificador é formado por `Título + Nome` da aba `Config`: se a emissão já
existe, ela é atualizada; se for nova, é adicionada. Ao rodar para um arquivo
específico, as demais emissões já publicadas são preservadas.
Arquivos com `backup`, `teste` ou `template` no nome sao ignorados na varredura automatica.
Para arquivos específicos: `python atualizar_dados.py "Consolidado - CRI UNIQ.xlsx"`.

## Ver localmente

```
cd site
python -m http.server 8080
```

Abra http://localhost:8080 (o `fetch` do dados.json não funciona abrindo o
arquivo direto do disco — precisa de servidor ou hospedagem).

## Publicar no GitHub Pages (recomendado)

1. Instale Git + GitHub CLI (`winget install Git.Git GitHub.cli`) e `gh auth login`.
2. Uma vez, dentro da pasta `site`:
   ```
   git init
   git add index.html dados.json COMO_PUBLICAR.md
   git commit -m "portal"
   gh repo create highpar-cri --public --source . --push
   ```
3. No GitHub: Settings → Pages → Deploy from branch → `main`, pasta `/ (root)`.
4. **Todo mês:** rode `python atualizar_dados.py`, depois:
   ```
   cd site
   git add dados.json && git commit -m "safra do mês" && git push
   ```
   O site atualiza sozinho em ~1 minuto. O dados.json no repositório é o
   banco de dados compartilhado — todo mundo vê a mesma base.

Alternativas: Netlify (arraste a pasta `site`) ou servidor interno (copie a pasta).

## Arquivos

| Arquivo | Papel |
|---|---|
| `index.html` | Portal (estático, genérico — lê o que estiver no dados.json) |
| `dados.json` | Banco de dados (gerado pelo atualizar_dados.py, nunca edite à mão) |
| `HIGHPAR-LOGO.svg` | Logo usada no menu lateral do portal |
| `index_antigo_azul.html` | Versão anterior (backup) |

## v2 — o que mudou
- Farol (ILG, LTV, carteira do mês vs esperada) é CALCULADO PELO PORTAL — a
  planilha só fornece dados brutos e limiares (Config).
- Home com gauge de risco por emissão (pouco/médio/alto) + ILG/LTV/Carteira.
- Lâmina: fluxo do investidor com payback, razões de garantia em linha,
  recebimentos mensais com acumulado no período completo e tabela de métricas
  filtrada (botão "mostrar todas").
- Dá para arrastar um consolidado .xlsx direto na página (vale na sessão;
  para persistir: atualizar_dados.py + git push do dados.json).

## v3 — ATUALIZADO1 (jul/2026)
- O portal agora carrega **dados.js** (mesmo banco em formato script): dá para abrir o
  `index.html` com dois cliques, SEM servidor. O dados.json continua sendo gerado (fetch
  em hospedagem e integrações). **Publique os dois** (`git add index.html dados.json dados.js`).
- Novo layout (padrão Opea/Riza), Radar de Risco fora da lâmina, aba Garantias & Covenants,
  posição Highpar, exportar lâmina em PDF (botão de imprimir) e cadastro de novo CRI.
- A pasta `cadastro/` (fora do site) guarda covenants/garantias/posição por emissão —
  edite lá e rode `python atualizar_dados.py`.
