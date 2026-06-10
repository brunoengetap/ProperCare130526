# Prompt para Codex — Correção cirúrgica da aba Estoque e KPIs de Preventivas do ProperCare

## Repositório e arquivos

Repositório: `brunoengetap/ProperCare130526` (branch `main`)

| Arquivo a ler (fonte) | Arquivo a gerar (saída) | Observação |
|---|---|---|
| `PCM_v26.html` | `PCM_v27.html` | Salvar no mesmo diretório |
| `GAS_properCare_v18.js` | `GAS_properCare_v19.js` | Salvar no mesmo diretório |
| `Proper_Field_grupotap_v12.html` | — | **Não tocar** |

Ao finalizar, confirmar que os dois arquivos de saída foram gravados e listar as funções alteradas.

---

## Objetivo

Corrigir a aba **Estoque** e os cards/KPIs da aba **Controle de Preventivas** sem quebrar funcionalidades já aprovadas. Cada correção deve ser cirúrgica: alterar somente os trechos identificados abaixo, sem refatoração ampla.

---

## Problemas confirmados por auditoria do código

### Bug 1 — CRÍTICO: `getMachineWorstStatus` retorna `level:'faltando'`, mas counts e filtros procuram por `null` ou `'sem_dados'` — incompatibilidade de strings causa o colapso dos KPIs

**Raiz do problema** (confirmar com grep antes de editar):

```
grep -n "'faltando'\|'sem_dados'\|sem-dados" PCM_v26.html
```

Resultado esperado do grep:
- Linha ~2000: `return {level:'faltando', ...}` — `getMachineWorstStatus` retorna `'faltando'`
- Linha ~2795: `sem_dados: rows.filter(r=>!r.ws).length` — conta somente `null`, nunca `'faltando'`
- Linha ~2794: `ok: rows.filter(r=>!r.ws||r.ws.level==='ok').length` — trata `null` como OK
- Linha ~2786: `if (!ws) return _prevFilter === 'sem_dados'` — filtro captura `null`, não `'faltando'`

**Consequência direta dos prints enviados:**
- Card "Em dia" mostra 4, mas ao clicar não exibe nenhuma máquina — porque `counts.ok` inclui máquinas com `ws=null` (sem dados), mas o filtro `'ok'` busca `ws.level==='ok'`, e essas máquinas têm `ws=null` ou `ws.level==='faltando'`
- Card "Sem dados" mostra 4, mas é `no-action` e nunca responde — porque `counts.sem_dados` conta `!ws` (null), mas as máquinas reais estão com `level:'faltando'` e o card não tem handler de clique

**Correção obrigatória:**

1. Em `getMachineWorstStatus()`, trocar o retorno de `'faltando'` por `'sem_dados'`:
   ```js
   // ANTES:
   if(countSemDados>0) return {level:'faltando', count:countSemDados, label:`...`};
   return null;
   
   // DEPOIS:
   if(countSemDados>0) return {level:'sem_dados', count:countSemDados, label:`${countSemDados} peça${countSemDados>1?'s':''} sem dados`};
   return {level:'ok', count:0, label:'OK'};
   ```
   Atenção: `return null` também deve ser eliminado — `getMachineWorstStatus` deve sempre retornar objeto com `level` explícito (`'vencido'`, `'alerta'`, `'sem_dados'` ou `'ok'`). Máquina sem nenhuma peça avaliável retorna `'sem_dados'`.

2. Em `renderPreventivas()`, corrigir `counts` para usar `level` explícito:
   ```js
   // ANTES:
   ok:        rows.filter(r=>!r.ws||r.ws.level==='ok').length,
   sem_dados: rows.filter(r=>!r.ws).length,
   
   // DEPOIS:
   ok:        rows.filter(r=>r.ws?.level==='ok').length,
   sem_dados: rows.filter(r=>r.ws?.level==='sem_dados').length,
   ```

3. Em `renderPreventivas()`, corrigir o filtro que usa `!ws`:
   ```js
   // ANTES:
   if (!ws) return _prevFilter === 'sem_dados';
   return ws.level === _prevFilter;
   
   // DEPOIS:
   return ws?.level === _prevFilter;
   ```

4. Em `renderPreventivas()`, corrigir a ordenação — `'faltando'` não existe mais:
   ```js
   // ANTES:
   const order={vencido:0,alerta:1,faltando:2,ok:3,null:4};
   
   // DEPOIS:
   const order={vencido:0,alerta:1,sem_dados:2,ok:3};
   // e na sort: order[a.ws?.level ?? 'sem_dados']
   ```

5. Em `renderPreventivas()`, o card superior "Sem dados" está com `no-action`. Tornar clicável quando `counts.sem_dados > 0`:
   ```js
   // ANTES:
   <div class="stat-card no-action">
     <div class="stat-val" ...>${counts.sem_dados}</div>
     <div class="stat-lbl">Sem dados</div>
   
   // DEPOIS:
   <div class="stat-card ${counts.sem_dados ? '' : 'no-action'}"
        ${counts.sem_dados ? `style="cursor:pointer" onclick="setPrevFilter('sem_dados')" title="Ver máquinas sem dados de peças"` : ''}>
     <div class="stat-val" ...>${counts.sem_dados}</div>
     <div class="stat-lbl">${counts.sem_dados ? 'Sem dados →' : 'Sem dados'}</div>
   ```

6. Verificar e corrigir todas as outras ocorrências de `ws?.level==='faltando'` ou `ws.level==='faltando'` no arquivo — substituir por `'sem_dados'`:
   ```
   grep -n "faltando" PCM_v26.html
   ```
   Esperado nas linhas ~3359 e ~5521 (badges de status na tabela de máquinas e na aba Clientes). Trocar `'faltando'` por `'sem_dados'` nessas comparações também.

7. Em `renderDashboard()`, corrigir o mesmo padrão:
   ```
   grep -n "!ws.*nSemDados\|nSemDados.*!ws\|!ws.*nOk\|nOk.*!ws" PCM_v26.html
   ```
   Se encontrado, refatorar para usar `ws.level` explícito nas contagens de `nVencidas`, `nAlertas`, `nOk`, `nSemDados`.

---

### Bug 2: `calcularEstoque()` inclui peças sem referência real e cria grupo `—`

**Localização (confirmar com grep):**
```
grep -n "ref.*||.*'—'\|const ref = part" PCM_v26.html
```
Esperado em ~linha 5140: `const ref = part.ref || '—';`

**Correção:**
```js
// ANTES (dentro do forEach de parts em calcularEstoque):
const ref = part.ref || '—';
// ... a peça sempre entra no resultados independente de ref

// DEPOIS: pular peça se não tiver referência real
const ref = (part.ref || '').trim();
if (!ref || ref === '—' || ref === '-') return; // sem referência = sem ação de compra
```

Após essa correção, o agrupamento por `normalizeRef(r.ref)` nunca receberá chave vazia.

---

### Bug 3: `calcularEstoque()` não resolve equivalentes/similares do catálogo

**Situação atual:** a função agrupa por `ref` da peça da máquina, mas não inclui refs alternativas do catálogo (`similarities`) na linha de estoque.

**Correção em `renderEstoqueResultado()`** (ou equivalente de renderização):
- Para cada grupo de resultado, buscar no catálogo (`db.catalog` / `db.parts`) a entrada que corresponde à ref principal
- Extrair o array `similarities` dessa entrada de catálogo
- Renderizar coluna ou sublinha "Equivalentes/Similares" com os códigos, separados por ` · `
- Se `similarities` for vazio ou ausente, não renderizar essa sublinha

Não incluir referência vazia (`''`) nem `'—'` no array de similares renderizado.

---

### Bug 4: Peças sem `lastChange` real sendo tratadas como vencidas nos KPIs e no Estoque

**Padrão problemático (confirmar):**
```
grep -n "lastChange || 0\|lastChange||0" PCM_v26.html
```
Esperado em ~linha 4791: `calcPartStatus(p.lastChange||0, p.interval, eq.hourTotal||0)`

**Em `_getMachineWorst(eq)` (modal badge):**
```js
// ANTES:
const {status}=calcPartStatus(p.lastChange||0, p.interval, eq.hourTotal||0);

// DEPOIS:
if (p.lastChange === undefined || p.lastChange === null || p.lastChange === '') continue;
const {status}=calcPartStatus(p.lastChange, p.interval, eq.hourTotal||0);
```

**Em `calcPartStatus()` e `getPartStatus()`:**
Confirmar que campos `lastChange` vazios retornam `'sem_dados'`, não `'vencido'`. Verificar:
```
grep -n "function calcPartStatus\|function getPartStatus" PCM_v26.html
```
Se `calcPartStatus(0, ...)` retorna vencido (porque `lastChange=0` e `hourTotal > interval`), isso é correto — `0` é horímetro real. Se `calcPartStatus('', ...)` retorna vencido, isso é o bug. A lógica existente em `getPartStatus()` (linha ~1978) já parece proteger contra `last === null`, mas `_getMachineWorst` o contorna com `|| 0`.

**Em `calcularEstoque()`:**
O guard `hasVal(part.lastChange)` já existe na função. Verificar se `hasVal(0)` retorna `true` (zero é válido):
```
grep -n "function hasVal" PCM_v26.html
```
Se `hasVal` usa `!!v` ou truthy check, corrigir para:
```js
function hasVal(v) { return v !== undefined && v !== null && v !== ''; }
```

---

### Bug 5 — GAS: `getMachinePartsById()` e `updateMachineParts()` convertem células vazias em `0` ou `2000`

**Confirmar com grep:**
```
grep -n "parseInt.*|| 0\|parseInt.*|| 2000" GAS_properCare_v18.js
```
Esperado nas linhas ~588, ~705–720, ~802–803.

**Adicionar helpers antes das funções afetadas:**
```js
function intOrBlank(v) {
  if (v === undefined || v === null || v === '') return '';
  const s = String(v).trim();
  if (s === '') return '';
  const n = Number(s.replace(',', '.'));
  return Number.isFinite(n) ? Math.trunc(n) : '';
}
// Atenção: intOrBlank(0) e intOrBlank('0') devem retornar 0 (número), não ''
```

**Em `getMachinePartsById()` (~linha 802):**
```js
// ANTES:
lastChange: parseInt(data[i][idxLch]) || 0,
interval:   parseInt(data[i][idxInt]) || 2000,

// DEPOIS:
lastChange: intOrBlank(data[i][idxLch]),
interval:   intOrBlank(data[i][idxInt]),
```

**Em `updateMachineParts()` (~linhas 705–720):**
```js
// ANTES:
parseInt(ps.lastChange) || 0,
parseInt(ps.interval)   || 2000,

// DEPOIS (update de linha existente):
// Se ps.lastChange for '' ou undefined, preservar valor existente da célula — NÃO gravar 0
// Se ps.lastChange for 0 (explícito), gravar 0
// Implementar: const lcVal = (ps.lastChange !== '' && ps.lastChange !== undefined && ps.lastChange !== null) ? intOrBlank(ps.lastChange) : existingCellValue;
```
Para leitura do valor existente, usar a linha corrente da sheet antes de sobrescrever.

**Em `saveVisit()` e `savePreventiva()` (~linhas 588, 972, 1001):**
```js
// Campos de peça (Últ.Troca, Intervalo_H):
parseInt(ps.lastChange) || 0  →  intOrBlank(ps.lastChange)
parseInt(ps.interval)   || 0  →  intOrBlank(ps.interval)

// Campos de horímetro de máquina (hourTotal, hpw) podem manter parseInt(...) || 0
// pois horímetro zero é raro mas válido como fallback
```

---

### Bug 6 — GAS: `getVisitsNormalized()` perde referência quando `Ref_Nova` está vazia

**Confirmar:**
```
grep -n "piRef.*indexOf\|Ref_Nova\|refNova\|refBase" GAS_properCare_v18.js
```
Esperado em ~linha 837: `const piRef = plH.indexOf('Ref_Nova');`

**Correção:**
```js
// ANTES (linha ~855):
// ref vem somente de Ref_Nova

// DEPOIS:
const piRefNova = plH.indexOf('Ref_Nova');
const piRefBase = plH.indexOf('Ref.');
// ...
ref: (piRefNova >= 0 ? cellStr(plData[r][piRefNova]) : '') || (piRefBase >= 0 ? cellStr(plData[r][piRefBase]) : ''),
```

---

## Verificações obrigatórias (grep pós-edição)

Executar após cada arquivo editado e reportar resultado:

```bash
# 1. Confirmar que 'faltando' foi eliminado do PCM
grep -c "faltando" PCM_v27.html
# Esperado: 0

# 2. Confirmar que getMachineWorstStatus nunca retorna null
grep -n "return null" PCM_v27.html | grep -A2 -B2 "getMachineWorstStatus"
# Esperado: nenhuma ocorrência dentro da função

# 3. Confirmar eliminação de lastChange||0 em _getMachineWorst
grep -n "lastChange||0\|lastChange || 0" PCM_v27.html
# Esperado: 0 ocorrências

# 4. Confirmar que calcularEstoque não inclui ref vazia
grep -n "ref.*||.*'—'\|part\.ref || '—'" PCM_v27.html
# Esperado: 0 ocorrências

# 5. Confirmar helpers intOrBlank no GAS
grep -n "intOrBlank" GAS_properCare_v19.js | head -10
# Esperado: definição da função + chamadas nos locais corrigidos

# 6. Confirmar que parseInt || 0 foi removido de lastChange/interval no GAS
grep -n "parseInt(ps.lastChange)\|parseInt(ps.interval)\|parseInt(data\[i\]\[idxLch\])" GAS_properCare_v19.js
# Esperado: 0 ocorrências

# 7. Confirmar fallback de ref no getVisitsNormalized
grep -n "Ref_Nova\|Ref\.\|piRefNova\|piRefBase" GAS_properCare_v19.js | grep -v "\/\/"
# Esperado: fallback para Ref. presente
```

---

## Testes manuais esperados

1. **KPI "Em dia" — consistência:** número do card deve bater com linhas na tabela ao clicar
2. **KPI "Sem dados" — clicável:** card deve responder ao clique e mostrar exatamente as máquinas com `level:'sem_dados'`
3. **Nenhuma máquina em "Em dia" e "Sem dados" simultaneamente:** os dois conjuntos devem ser disjuntos
4. **Estoque sem grupo `—`:** tabela e PDF não devem conter linha com referência vazia ou `—`
5. **Zero preservado:** lançar `lastChange=0` explicitamente → salvar → sincronizar → recarregar → valor deve permanecer `0`, não `''` nem `2000`
6. **Máquina sem referência não aparece no Estoque:** máquina com peça vencida mas sem ref deve aparecer em Preventivas, mas não no Estoque

---

## Restrições absolutas

- Não alterar `machineKey()`, `clientKey()`, `proper_admin_v2`, cabeçalhos de planilha, CRUD de clientes/máquinas/catálogo, nenhuma função de PDF/foto do PCF
- Não alterar assinaturas de `pgpFlushPending`, `pgpPostJson`, `pgpBuildVisitRequest`, `pgpBuildPartsUpdateRequest`
- Não remover as 9 peças padrão da UI — apenas impedir que sejam tratadas como dado real quando vazias
- `0` é valor válido — nunca tratar `0` como vazio em horímetro, lastChange ou valorMostrado
- Deploy order após entrega: **GAS primeiro**, depois PCM

---

## Entrega

1. Arquivos `PCM_v27.html` e `GAS_properCare_v19.js` gravados no mesmo diretório dos originais
2. Resumo com:
   - Lista de funções alteradas (nome + arquivo)
   - Resultado dos greps de validação
   - Confirmação dos testes manuais realizados
