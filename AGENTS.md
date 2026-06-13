# AGENTS.md — ProperCare / TapControl Monorepo

## Visão geral

Este repositório contém dois sistemas web para gestão de manutenção de máquinas industriais:

- **ProperCare** — sistema da empresa Proper (PCM + PCF + GAS)
- **TapControl** — sistema da empresa GrupoTAP (TCM + TCF + GAS), usado como **referência arquitetural**

Ambos compartilham a mesma arquitetura: Manager (web admin) + Field (app mobile) + Google Apps Script (backend/banco via Google Sheets).

---

## Estrutura de arquivos

| Arquivo | Papel |
|---------|-------|
| `GAS_properCare_v20.js` | Backend ProperCare — Google Apps Script (deploy como Web App) |
| `PCM_v29.html` | ProperCare Manager — painel administrativo (desktop) |
| `Proper_Field_grupotap_v13.html` | ProperCare Field — app do técnico (mobile) |
| `GAS_Code_S22_0.js` | Backend TapControl — referência |
| `TCM_S23_1.html` | TapControl Manager — referência |
| `TCF_S24_0.html` | TapControl Field — referência |
| `PROMPT_Codex_Fase1_Completo.md` | Spec de correções Fase 1 — **ler antes de qualquer mudança** |

---

## Regras absolutas

1. **Não reordenar** colunas em nenhum array `HEADERS.*` do GAS — Sheets existentes quebraria
2. **Novos campos** sempre ao **final** do array (append-only)
3. **Não modificar** estas funções: `machineKey`, `clientKey`, `pgpFlushPending`, `pgpPostJson`, `pgpBuildVisitRequest`, `pgpBuildPartsUpdateRequest`, `pgpMachineKey`, `pgpClientKey`
4. **Não alterar** a chave `proper_admin_v2` do localStorage
5. **Versionar** ao salvar: `v20→v21`, `v29→v30`, `v13→v14`
6. **Deploy obrigatório na ordem:** GAS → PCM → PCF

---

## Convenções de código

- CSS de modal: `.overlay { display:none }` / `.overlay.open { display:flex }` — **nunca usar `.overlay.show`**
- Abrir modal: `openModal('modalId')` que adiciona classe `open` — não inserir modal com classe `open` diretamente a não ser ao criar dinamicamente
- Fetch para GAS: sempre via `gsGet(action, params)` ou `syncToGS(action, payload)` — não usar fetch direto
- Toasts: `toast('mensagem', 'tipo')` no PCM / `showNotification('mensagem')` no PCF
- Arrays no Sheets: serializados como JSON string — usar `parseJsonArray_()` para ler

---

## Fluxo de dados

```
PCM / PCF  →  (fetch POST/GET)  →  GAS Web App  →  Google Sheets
```

- Autenticação do técnico: PIN → SHA-256 (SubtleCrypto no browser / Utilities no GAS) → comparado com `pin_hash` em `CAT_TECNICOS`
- OS numeradas automaticamente pelo GAS (`PGP-YYYY-NNNN`)
- Fila offline no PCF: `localStorage` key `proper_pending` — flushed por `pgpFlushPending`

---

## O que está funcionando (não tocar)

- Login por PIN com SHA-256
- Numeração automática de OS
- Cronômetro de atendimento (`_pgpOSInicio`)
- Assinaturas por canvas touch + upload Drive
- Schema `CAT_TECNICOS` e `ORDENS_SERVICO` no GAS
- Filtros de OS no PCM (status/técnico/cliente)
- Persistência de sessão do técnico em `sessionStorage`
- Pipeline de fila offline (`pgpQueueAnexo`, `pgpFlushAnexos`)

---

## Tarefa atual — Fase 1 (ver `PROMPT_Codex_Fase1_Completo.md`)

Ler o arquivo `PROMPT_Codex_Fase1_Completo.md` para a spec completa. Resumo dos blocos:

- **Bloco 1 (GAS v21):** `tipo_os` no schema, endpoints `getTecnicos`/`saveTecnico`, `seedTecnicosDefault`
- **Bloco 2 (PCM v30):** corrigir modais invisíveis (`show`→`open`), seletor de técnicos, campo `tipo_os`, filtro de máquinas por cliente, CRUD de técnicos
- **Bloco 3 (PCF v14):** 4 bug fixes de OS, desabilitar seleção de tipo durante OS, badge informativo, separação visual, propagar `tipo_os`
- **Bloco 4:** teste E2E de 12 passos — obrigatório antes de considerar Fase 1 completa
