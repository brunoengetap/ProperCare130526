# CORREÇÃO CIRÚRGICA — PCM — Dropdown de clientes duplicado por variação de grafia

Arquivo a editar: **`PCM_v13.html`** (único arquivo com alterações nesta tarefa).
Não alterar `GAS_properCare_v10.js` nem `pcf_v11.html`.

---

## Causa raiz diagnosticada

O dropdown "Filtrar por cliente" em **Máquinas**, e a lista em **Plano de Preventivas**, são construídos com um objeto `byClient` usando o campo bruto `eq.client` como chave:

```js
// CÓDIGO ATUAL — gera duplicatas
const byClient = {};
db.machines.forEach(eq => {
  const k = eq.client || 'Sem cliente';   // ← nome bruto, sem normalização
  if (!byClient[k]) byClient[k] = [];
  byClient[k].push(eq);
});
```

Como máquinas diferentes têm `client` gravado com grafias distintas (`"PROPER locação"`, `"Proper locação"`, `"Proper Locação"`), o `byClient` gera **uma entrada separada por variante**, e o dropdown exibe todas elas como opções distintas.

A função `clientKey()` já existe e já resolve isso — só não é usada na construção do `byClient`.

---

## Correção

Há **três ocorrências** do mesmo padrão a corrigir em `PCM_v13.html`. Localize cada uma pelo contexto abaixo e aplique a substituição.

### Função auxiliar — adicionar uma vez, antes de `renderDashboard`

Adicionar esta função auxiliar (se ainda não existir):

```js
/**
 * Agrupa db.machines por clientKey, retornando um mapa:
 *   { clientKey → { label: string, machines: [] } }
 * label = nome canônico preferindo db.clients > primeira máquina encontrada
 */
function buildMachinesByClientKey() {
  const map = {};
  db.machines.forEach(eq => {
    const raw = eq.client || '';
    if (!raw) return;
    const k = clientKey(raw);
    if (!k) return;
    if (!map[k]) {
      // Preferir o nome canônico de db.clients se existir
      const registered = (db.clients || []).find(c => clientKey(c.nome) === k);
      map[k] = { label: registered ? registered.nome : raw, machines: [] };
    }
    map[k].machines.push(eq);
  });
  return map;
}
```

---

### Ocorrência 1 — `renderMachines()`

**Localizar** o trecho:
```js
const byClient={};
db.machines.forEach(eq=>{const k=eq.client||'Sem cliente';if(!byClient[k])byClient[k]=[];byClient[k].push(eq)});
```

**Substituir por:**
```js
// Agrupar por clientKey para evitar duplicatas por variação de grafia
const _byClientKey = buildMachinesByClientKey();
const byClient = {}; // mapa label → array (para compatibilidade com o restante da função)
Object.values(_byClientKey).forEach(({ label, machines }) => {
  byClient[label] = machines;
});
```

Também localizar a linha que gera as `<option>` do dropdown:
```js
${Object.keys(byClient).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'})).map(c=>`<option value="${c.replace(/"/g,'&quot;')}" ${_machFilter===c?'selected':''}>${c} (${byClient[c].length})</option>`).join('')}
```

**Substituir por** (a comparação `selected` passa a usar `clientKey` para tolerar `_machFilter` com grafia diferente do label):
```js
${Object.keys(byClient).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'})).map(c=>`<option value="${c.replace(/"/g,'&quot;')}" ${clientKey(_machFilter)===clientKey(c)?'selected':''}>${c} (${byClient[c].length})</option>`).join('')}
```

---

### Ocorrência 2 — `renderPreventivas()`

**Localizar** o trecho idêntico dentro de `renderPreventivas` (linha ~2312):
```js
const byClient={};
db.machines.forEach(eq=>{const k=eq.client||'Sem cliente';if(!byClient[k])byClient[k]=[];byClient[k].push(eq)});
```

**Substituir pela mesma lógica:**
```js
const _byClientKey = buildMachinesByClientKey();
const byClient = {};
Object.values(_byClientKey).forEach(({ label, machines }) => {
  byClient[label] = machines;
});
```

---

### Ocorrência 3 — `renderDashboard()`

**Localizar** o mesmo padrão dentro de `renderDashboard` (linha ~2312):
```js
const byClient={};
db.machines.forEach(eq=>{const k=eq.client||'Sem cliente';if(!byClient[k])byClient[k]=[];byClient[k].push(eq)});
```

**Substituir:**
```js
const _byClientKey = buildMachinesByClientKey();
const byClient = {};
Object.values(_byClientKey).forEach(({ label, machines }) => {
  byClient[label] = machines;
});
```

> `renderDashboard` usa `Object.keys(byClient).length` para contar clientes distintos — com a correção, esse número passa a ser correto.

---

## Restrições

- Não alterar `clientKey`, `normalizeTextKey`, `isSuspectClientTerm` — estão corretas.
- Não alterar `machinesByClient` em `renderClients` — já usa `clientKey` como chave (correto).
- Não alterar `buildCanonicalClientOptions` — já está correto.
- Não alterar lógica de filtro `matchClient` na linha `const matchClient=!_machFilter||clientKey(eq.client)===clientKey(_machFilter)` — já usa `clientKey` corretamente.
- `buildMachinesByClientKey` deve ser adicionada **uma única vez**.

---

## Checklist de entrega

- [ ] `buildMachinesByClientKey()` adicionada uma vez
- [ ] `renderMachines`: `byClient` usa `buildMachinesByClientKey`
- [ ] `renderMachines`: `<option selected>` compara com `clientKey`
- [ ] `renderPreventivas`: `byClient` usa `buildMachinesByClientKey`
- [ ] `renderDashboard`: `byClient` usa `buildMachinesByClientKey`
- [ ] Dropdown não exibe mais "PROPER locação", "Proper locação" e "Proper Locação" como entradas separadas
- [ ] Contagem de clientes distintos no Dashboard está correta
