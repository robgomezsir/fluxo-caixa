# Protocolo WebView2 — Fluxo de Caixa CRUD

Comunicação bidirecional via `window.chrome.webview.postMessage` / `WebMessageReceived`.

## Request

```json
{
  "id": "uuid",
  "action": "ping|meta|list|get|create|update|upsert|delete|resumo",
  "entity": "capex|novos_neg|premissas_anot|bonus|matrix|matrix_*|rubricas",
  "payload": {}
}
```

## Response

```json
{
  "id": "uuid",
  "success": true,
  "data": {},
  "error": null
}
```

## Entidades tabulares

| Entity | Aba | Campos principais |
|--------|-----|-------------------|
| `capex` | CAPEX (2) | unidade, fluxo, pctTerral, nomeProjeto, valor (cents), origemCapex, classificacao |
| `novos_neg` | Novos Neg. | unidade, codigo, descricao, valor (cents), payback |
| `premissas_anot` | ANOTAÇÕES DE PREMISSAS | nota, unidade, criterio, responsavel |
| `bonus` | bonus_caixa_2020 | undPertencente, filialPagante, valorProvisao (cents), marco, abril, maio |

## Entidade matricial

`entity`: `matrix` com `payload.sheet`, `payload.rubrica`, `payload.mes`, `payload.tipo`, `payload.valorCents`

Ações: `list`, `get`, `upsert`, `delete` (zera valor)

## Meta

`action: "meta"` retorna entidades, abas matriciais, meses, tipos e listas fixas: `unidades`, `origensCapex`, `centrosCusto`, `classificacoesCapex`, `responsaveis`, `simNao`.

## Resumo

`action: "resumo"` retorna cards, alertas #REF! e snippet PrevxReal.
