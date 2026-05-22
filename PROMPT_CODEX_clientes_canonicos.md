# CORREÇÃO CIRÚRGICA — PCM/GAS — Identidade canônica de clientes, merge e busca sem duplicidades

> **Instrução inicial ao Codex:** Este arquivo se chama `PROMPT_CODEX_clientes_canonicos.md` e está no repositório `brunoengetap/ProperCare130526`. Leia-o por completo antes de qualquer edição. Os arquivos a modificar (`PCM_v13.html`, `GAS_properCare_v10.js`) também estão no mesmo repositório. O `pcf_v11.html` é referência somente.

---

## Arquivos principais

- `PCM_v13.html` — ProperCare Manager (desktop)
- `GAS_properCare_v10.js` — backend Google Apps Script
- `pcf_v11.html` — verificação de compatibilidade apenas; **não alterar** salvo necessidade indispensável

## Schema real — leia antes de codar

### `db.clients[]` — campos reais no PCM
```
id, nome, cnpj, cidade, telefone, email,
antecedencia_alerta_dias, drive_url, observacoes,
tipoRegistro, createdAt, updatedAt
```
> Não existem `filial`, `contato`, `contact`, `phone`, `city`, `branch` em `db.clients`.
> Esses campos existem em `db.machines[]`, não em clientes.

### `db.machines[]` — campos reais no PCM
```
id, client, clientId, branch, cnpj, contact, phone, email,
city, brand, model, type, serial, year, tag, location,
power, hourTotal, hpw, obs, driveLink, updatedAt
```

### Aba `CLIENTES` no Google Sheets (HEADERS.CLIENTES)
```
ID, Nome, CNPJ, Cidade, Telefone, Email, Observações,
Antecedencia_Alerta_Dias, Drive_URL, Ativo, Tipo_Registro,
Created_At, Deleted_At, Deleted_By, Atualizado,
Contato, Filial, Prox_Visita, Ult_Visita
```

### Aba `MAQUINAS` no Google Sheets (HEADERS.MAQUINAS)
```
ID, Cliente, Filial, Marca, Modelo, Série, Ano, TAG, Localização,
Hor.Total, h/Semana, Observações, Ativo, Tipo_Registro,
Created_At, Deleted_At, Deleted_By, Atualizado
```

### Aba `VISITAS` no Google Sheets (HEADERS.VISITAS)
Coluna de cliente: **`Cliente`** (índice via `headers.indexOf('Cliente')`)

---

## Contexto do bug real

Existe um cliente errado chamado `3844` (que é na verdade o número de série de uma máquina). Deve ser mesclado com `Fundição Claudiense`. Problemas:

1. Na aba **Clientes**, ao editar `3844`, digitar `Fundição Claudiense` e clicar em **"Editar este"** no alerta de similaridade, nada acontece de forma visível.
2. Causa: o botão chama `openModalClient('${c.id}');closeModal('modalClient')` — abre e fecha o modal imediatamente. **Atenção:** `closeModal` foi sobrescrita no PCM (linha ~1300) e chama `closeModalSafe`, que tem lógica de guarda de estado. A solução é substituir o botão por uma nova função que **nunca chame `closeModal` nem `closeModalSafe`**.
3. Mesmo renomeando `3844`, as máquinas vinculadas a ele não são transferidas para `Fundição Claudiense`.
4. Em **Plano de Preventivas > editar máquina**, digitar `fund...` mostra opções duplicadas por diferença de maiúsculas/acentos.

---

# 1. PCM — Corrigir botão "Editar este" / "Mesclar com este"

Em `checkClientDuplicates(val)`, substituir o `onclick` do botão `Editar este` por uma nova função:

```js
let _pendingClientMergeSourceId = null;
let _pendingClientMergeSourceName = '';

function selectDuplicateClientForMerge(targetId) {
  const target = db.clients.find(c => c.id === targetId);
  if (!target) { toast('Cliente de destino não encontrado', 'error'); return; }

  const source = editClientId ? db.clients.find(c => c.id === editClientId) : null;
  _pendingClientMergeSourceId = (source && source.id !== target.id) ? source.id : null;
  _pendingClientMergeSourceName = source ? source.nome : '';

  // Carregar o cliente canônico no modal sem fechar/reabrir
  editClientId = target.id;
  document.getElementById('clNome').value       = target.nome || '';
  document.getElementById('clCnpj').value       = target.cnpj || '';
  document.getElementById('clCidade').value     = target.cidade || '';
  document.getElementById('clTelefone').value   = target.telefone || '';
  document.getElementById('clEmail').value      = target.email || '';
  document.getElementById('clAlerta').value     = target.antecedencia_alerta_dias || 30;
  document.getElementById('clDrive').value      = target.drive_url || '';
  document.getElementById('clObs').value        = target.observacoes || '';

  // Preencher campos vazios do target com dados do source, se source existir
  if (source) {
    if (!target.cnpj       && source.cnpj)       document.getElementById('clCnpj').value     = source.cnpj;
    if (!target.cidade     && source.cidade)     document.getElementById('clCidade').value   = source.cidade;
    if (!target.telefone   && source.telefone)   document.getElementById('clTelefone').value = source.telefone;
    if (!target.email      && source.email)      document.getElementById('clEmail').value    = source.email;
    if (!target.drive_url  && source.drive_url)  document.getElementById('clDrive').value    = source.drive_url;
    if (!target.observacoes && source.observacoes) document.getElementById('clObs').value   = source.observacoes;
  }

  document.getElementById('dupClientAlert').classList.remove('show');
  document.getElementById('modalClientTitle').textContent = 'Mesclar com cliente';
  toast('Carregado: ' + target.nome + '. Salve para confirmar a mescla.', 'warning');
  console.log('[PCM client merge] selectDuplicateClientForMerge — source:', _pendingClientMergeSourceId, '→ target:', targetId);
}
```

No HTML do `dupClientList`, trocar:
```html
<!-- ANTES -->
<button onclick="openModalClient('${c.id}');closeModal('modalClient')">Editar este</button>

<!-- DEPOIS -->
<button onclick="selectDuplicateClientForMerge('${c.id}')">Mesclar com este</button>
```

Regras:
- Nunca chamar `closeModal` ou `closeModalSafe` dentro de `selectDuplicateClientForMerge`.
- O modal permanece aberto, carregado com o cliente canônico.
- O rótulo do botão muda para **"Mesclar com este"** em modo edição.

---

# 2. PCM — Corrigir `saveClient()` em modo edição

## 2.1 Merge seguro ao substituir o objeto

Ao editar cliente, não substituir o objeto inteiro — fazer merge preservando campos existentes. Os campos reais de `db.clients` a preservar são:

```js
const old = db.clients[idx] || {};
const updated = {
  ...old,
  ...data,
  id: old.id || data.id,
  // preservar campos que o modal não edita
  tipoRegistro: old.tipoRegistro || data.tipoRegistro || 'PRODUCAO',
  createdAt:    old.createdAt    || data.createdAt    || new Date().toISOString(),
  // não sobrescrever campo preenchido com string vazia
  cnpj:         data.cnpj       || old.cnpj       || '',
  cidade:       data.cidade     || old.cidade     || '',
  telefone:     data.telefone   || old.telefone   || '',
  email:        data.email      || old.email      || '',
  drive_url:    data.drive_url  || old.drive_url  || '',
  observacoes:  data.observacoes|| old.observacoes|| '',
};
db.clients[idx] = updated;
```

> **Não tentar preservar `filial`, `contato`, `lastVisit`, `nextVisit`** — esses campos não existem em `db.clients`. Pertencem a `db.machines`.

## 2.2 Detectar duplicidade também durante edição

Antes de salvar, verificar se já existe outro cliente com mesma chave canônica ou mesmo CNPJ:

```js
const onlyDigits = v => String(v || '').replace(/\D/g, '');
const targetDup = db.clients.find(c =>
  c.id !== editClientId && (
    clientKey(c.nome) === clientKey(data.nome) ||
    (onlyDigits(c.cnpj) && onlyDigits(data.cnpj) &&
     onlyDigits(c.cnpj) === onlyDigits(data.cnpj))
  )
);
```

Se `targetDup` existir: executar merge com o canônico (seções 2.3 e 2.4) em vez de salvar como objeto separado.

## 2.3 Reatribuir máquinas ao cliente canônico

Sempre que um cliente for renomeado ou mesclado, atualizar `db.machines`:

```js
function reassignMachinesToClient(sourceId, sourceName, canonical) {
  let count = 0;
  db.machines.forEach(m => {
    const matchById   = m.clientId === sourceId;
    const matchByKey  = clientKey(m.client || '') === clientKey(sourceName || '');
    const matchExact  = m.client === sourceName;
    // Para sourceName suspeito (numérico), usar apenas match exato
    const isSuspect   = isSuspectClientTerm(sourceName);
    const shouldUpdate = isSuspect ? matchExact : (matchById || matchByKey || matchExact);

    if (shouldUpdate) {
      m.client   = canonical.nome;
      m.clientId = canonical.id;
      // Preencher dados cadastrais vazios da máquina com dados do cliente canônico
      if (!m.cnpj     && canonical.cnpj)     m.cnpj     = canonical.cnpj;
      if (!m.city     && canonical.cidade)   m.city     = canonical.cidade;
      if (!m.phone    && canonical.telefone) m.phone    = canonical.telefone;
      if (!m.email    && canonical.email)    m.email    = canonical.email;
      if (!m.driveLink && canonical.drive_url) m.driveLink = canonical.drive_url;
      count++;
    }
  });
  console.log('[PCM client merge] reassignMachinesToClient:', count, 'máquinas atualizadas → ', canonical.nome);
  return count;
}
```

## 2.4 Remover/inativar cliente origem após merge

```js
// Remover source do array local
db.clients = db.clients.filter(c => c.id !== _pendingClientMergeSourceId);
_pendingClientMergeSourceId = null;
_pendingClientMergeSourceName = '';
```

Após merge, chamar (apenas as que existem no PCM v13):
```js
save();
renderClients();
renderMachines();
renderPreventivas();
renderSidebar();
```

> **Verificação:** `renderPreventivas` **existe** no PCM v13 (linha ~2316). Todas as quatro funções acima existem — usar exatamente esses nomes.

## 2.5 Sincronização pós-merge

```js
// Sincronizar cliente canônico
syncToGS('saveClient', { client: canonical });

// Sincronizar cada máquina alterada
machinasAlteradas.forEach(m => syncToGS('saveMachine', { machine: m }));

// Chamar endpoint de merge no GAS (implementado na seção 5)
syncToGS('mergeClients', {
  sourceId: _pendingClientMergeSourceId,
  sourceName: _pendingClientMergeSourceName,
  targetId: canonical.id,
  targetName: canonical.nome,
  client: canonical,
  deletedBy: 'admin'
});
```

Tratar falha de sync sem quebrar estado local (try/catch com toast de aviso).

---

# 3. PCM — Corrigir `suggestClient(val)` para deduplicar por `clientKey`

Substituir a implementação atual por:

```js
function buildCanonicalClientOptions() {
  const map = new Map();

  const put = (name, source, obj = null) => {
    if (!name || isSuspectClientTerm(name)) return;
    const k = clientKey(name);
    if (!k) return;
    const score = source === 'client' ? 3 : source === 'machine' ? 2 : 1;
    const existing = map.get(k);
    if (!existing || score > existing.score) {
      map.set(k, { key: k, name, source, obj, score });
    }
  };

  (db.clients  || []).forEach(c => put(c.nome,     'client',  c));
  (db.machines || []).forEach(m => put(m.client,   'machine', null));
  (PGP_CLIENTS || []).forEach(n => put(n,          'static',  null));

  return [...map.values()]
    .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'base' }));
}

function suggestClient(val) {
  const dl = document.getElementById('clientSuggestions');
  if (!val || val.length < 2) { dl.innerHTML = ''; return; }

  const vk = clientKey(val);
  const options = buildCanonicalClientOptions()
    .filter(o => o.key.includes(vk) || o.name.toLowerCase().includes(val.toLowerCase()))
    .slice(0, 10);

  dl.innerHTML = options.map(o => `<option value="${o.name}">`).join('');

  // Auto-preenchimento com dados do cliente canônico cadastrado
  if (!editMachineId) {
    const hit = options.find(o => o.obj && clientKey(o.name) === clientKey(val));
    if (hit && hit.obj) {
      // Mapeamento correto: IDs do formulário → campos reais de db.clients
      // db.clients usa: cnpj, cidade, telefone, email, drive_url
      // db.machines usa: branch, cnpj, contact, phone, email, city, driveLink
      const fieldMap = [
        { elId: 'mcCnpj',      val: hit.obj.cnpj      || '' },
        { elId: 'mcPhone',     val: hit.obj.telefone   || '' },  // db.clients.telefone → mcPhone
        { elId: 'mcEmail',     val: hit.obj.email      || '' },
        { elId: 'mcCity',      val: hit.obj.cidade     || '' },  // db.clients.cidade → mcCity
        { elId: 'mcDriveLink', val: hit.obj.drive_url  || '' },  // db.clients.drive_url → mcDriveLink
        // mcBranch e mcContact não têm equivalente direto em db.clients — não preencher
      ];
      fieldMap.forEach(({ elId, val: fVal }) => {
        const el = document.getElementById(elId);
        if (el && !el.value.trim() && fVal) el.value = fVal;
      });
    }
  }
}
```

> **Atenção crítica:** `db.clients` **não tem** `filial`, `contato`, `contact`, `phone`, `city`, `branch`. O mapeamento correto é `telefone → mcPhone`, `cidade → mcCity`, `drive_url → mcDriveLink`. Não tentar mapear campos inexistentes.

---

# 4. PCM — Padronizar uso de `clientKey`

## 4.1 Bug de sort na ordenação por máquinas (confirmado no código)

Em `renderClients()`, no bloco `else if(_clientSort==='maquinas')`:

```js
// ANTES (bug confirmado):
va = (machinesByClient[(a.nome||'').toLowerCase().trim()] || []).length;
vb = (machinesByClient[(b.nome||'').toLowerCase().trim()] || []).length;

// DEPOIS (correto — machinesByClient é indexado por clientKey):
va = (machinesByClient[clientKey(a.nome || '')] || []).length;
vb = (machinesByClient[clientKey(b.nome || '')] || []).length;
```

## 4.2 Outros pontos a padronizar com `clientKey`

Substituir comparações por `toLowerCase().trim()` ou `normalizeEntityName()` quando a intenção for **identidade de cliente** em:

- `saveMachine()` — normalização do nome do cliente antes de salvar
- `checkMachineDuplicateCandidate()` — comparação de cliente
- `openModalMachine()` — lookup do cliente cadastrado para autopreenchimento
- qualquer filtro `_machFilter` comparando nomes de cliente

---

# 5. GAS — Criar endpoint `mergeClients`

## 5.1 Registrar no `doPost`

```js
case 'mergeClients':
  result = mergeClients(body);
  break;
```

## 5.2 Implementar `mergeClients(body)`

```js
function mergeClients(body) {
  const { sourceId, sourceName, targetId, targetName, client, deletedBy } = body || {};
  const now = new Date().toISOString();
  Logger.log('[GAS mergeClients] source: ' + sourceName + ' (' + sourceId + ') → target: ' + targetName + ' (' + targetId + ')');

  const clientSheet = getOrCreateSheet('CLIENTES', HEADERS.CLIENTES);
  const clientData  = clientSheet.getDataRange().getValues();
  const cHeaders    = clientData[0];
  const idxCId      = cHeaders.indexOf('ID');
  const idxCNome    = cHeaders.indexOf('Nome');
  const idxAtivo    = cHeaders.indexOf('Ativo');
  const idxDelAt    = cHeaders.indexOf('Deleted_At');
  const idxDelBy    = cHeaders.indexOf('Deleted_By');

  const ck = v => clientKey(String(v || ''));
  const onlyDigits = v => String(v || '').replace(/\D/g, '');

  // Localizar target e source
  let targetRow = -1, sourceRow = -1;
  for (let i = 1; i < clientData.length; i++) {
    const rowId   = String(clientData[i][idxCId]   || '').trim();
    const rowNome = String(clientData[i][idxCNome] || '').trim();
    if (rowId === String(targetId || '').trim() || ck(rowNome) === ck(targetName)) {
      if (targetRow < 0) targetRow = i;
    }
    if (rowId === String(sourceId || '').trim() || rowNome === sourceName) {
      if (sourceRow < 0) sourceRow = i;
    }
  }

  // Mesclar campos vazios do target com dados do source/payload
  if (targetRow >= 0 && client) {
    const existingObj = {};
    cHeaders.forEach((h, i) => { existingObj[h] = clientData[targetRow][i]; });
    const mergeVal = (nv, ev) => (nv && String(nv).trim()) ? String(nv).trim() : String(ev || '');
    const updatedRow = cHeaders.map(h => {
      switch(h) {
        case 'Nome':   return mergeVal(client.nome,        existingObj['Nome']);
        case 'CNPJ':   return mergeVal(client.cnpj,        existingObj['CNPJ']);
        case 'Cidade': return mergeVal(client.cidade,      existingObj['Cidade']);
        case 'Telefone': return mergeVal(client.telefone,  existingObj['Telefone']);
        case 'Email':  return mergeVal(client.email,       existingObj['Email']);
        case 'Drive_URL': return mergeVal(client.drive_url, existingObj['Drive_URL']);
        case 'Observações': return mergeVal(client.observacoes, existingObj['Observações']);
        case 'Atualizado': return now;
        default: return existingObj[h] !== undefined ? existingObj[h] : '';
      }
    });
    clientSheet.getRange(targetRow + 1, 1, 1, cHeaders.length).setValues([updatedRow]);
  }

  // Inativar source (se source e target forem linhas diferentes)
  let sourceDeactivated = false;
  if (sourceRow >= 0 && sourceRow !== targetRow) {
    if (idxAtivo >= 0)  clientSheet.getRange(sourceRow + 1, idxAtivo  + 1).setValue('NÃO');
    if (idxDelAt >= 0)  clientSheet.getRange(sourceRow + 1, idxDelAt  + 1).setValue(now);
    if (idxDelBy >= 0)  clientSheet.getRange(sourceRow + 1, idxDelBy  + 1).setValue(deletedBy || 'merge');
    sourceDeactivated = true;
  }

  // Atualizar MAQUINAS
  const machSheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const machData  = machSheet.getDataRange().getValues();
  const mHeaders  = machData[0];
  const idxMCli   = mHeaders.indexOf('Cliente');
  const isSuspect = isSuspectClientTerm(sourceName);
  let machinesUpdated = 0;

  for (let i = 1; i < machData.length; i++) {
    const rowCli = String(machData[i][idxMCli] || '').trim();
    const matches = isSuspect
      ? (rowCli === sourceName)                          // sourceName numérico: match exato apenas
      : (ck(rowCli) === ck(sourceName) || rowCli === sourceName);
    if (matches) {
      machSheet.getRange(i + 1, idxMCli + 1).setValue(targetName);
      machinesUpdated++;
    }
  }

  // Atualizar VISITAS — coluna 'Cliente' (não atualizar MACHINE_PARTS, não tem cliente)
  const visitSheet = getOrCreateSheet('VISITAS', HEADERS.VISITAS);
  const visitData  = visitSheet.getDataRange().getValues();
  const vHeaders   = visitData[0];
  const idxVCli    = vHeaders.indexOf('Cliente');
  let visitsUpdated = 0;

  if (idxVCli >= 0) {
    for (let i = 1; i < visitData.length; i++) {
      const rowCli = String(visitData[i][idxVCli] || '').trim();
      const matches = isSuspect
        ? (rowCli === sourceName)
        : (ck(rowCli) === ck(sourceName) || rowCli === sourceName);
      if (matches) {
        visitSheet.getRange(i + 1, idxVCli + 1).setValue(targetName);
        visitsUpdated++;
      }
    }
  }

  Logger.log('[GAS mergeClients] machinesUpdated: ' + machinesUpdated + ', visitsUpdated: ' + visitsUpdated + ', sourceDeactivated: ' + sourceDeactivated);

  return {
    status: 'ok',
    merged: true,
    targetId,   targetName,
    sourceId,   sourceName,
    machinesUpdated,
    visitsUpdated,
    sourceDeactivated
  };
}
```

> **Não mexer em `MACHINE_PARTS`** — essa aba não tem coluna de cliente.

---

# 6. GAS — Melhorar `saveClient` para usar `clientKey` global

A função `saveClient(c)` tem uma declaração local:
```js
const _normEnt = v => String(v || '').trim().toLowerCase()...
```

**Remover completamente essa declaração local** e substituir todas as suas chamadas por `clientKey()` (função global já existente no GAS desde a linha 59).

Regras de match (em ordem de prioridade):
1. Match por `ID` exato
2. Match por CNPJ normalizado (`onlyDigits`), quando CNPJ existir em ambos
3. Match por `clientKey(nome)`

Nunca inserir novo cliente se já existir mesma chave canônica ou mesmo CNPJ. Não sobrescrever campo preenchido com string vazia (lógica `mergeVal` já existente — manter).

---

# 7. GAS — Melhorar `getDuplicates()`

Atualizar para detectar duplicidades de clientes:

```js
// Dentro de getDuplicates(), adicionar seção:
const clientsData  = getSheetData('CLIENTES');
const clientByKey  = {};
const clientByCnpj = {};

clientsData.forEach(c => {
  const nome = String(c['Nome'] || '').trim();
  const cnpj = String(c['CNPJ'] || '').replace(/\D/g, '');
  const ativo = String(c['Ativo'] || 'SIM').toUpperCase();
  if (ativo === 'NÃO') return;

  const k = clientKey(nome);
  if (k) {
    if (!clientByKey[k])  clientByKey[k] = [];
    clientByKey[k].push(nome);
  }
  if (cnpj.length >= 11) {
    if (!clientByCnpj[cnpj]) clientByCnpj[cnpj] = [];
    clientByCnpj[cnpj].push(nome);
  }
});

const dupClientsByKey  = Object.entries(clientByKey).filter(([,v]) => v.length > 1);
const dupClientsByCnpj = Object.entries(clientByCnpj).filter(([,v]) => v.length > 1);
```

Incluir `dupClientsByKey` e `dupClientsByCnpj` no retorno de `getDuplicates()`. Não criar UI nova agora.

---

# 8. PCF — Verificação de compatibilidade

Verificar apenas que as mudanças no GAS são compatíveis com estas funções do PCF (`pcf_v11.html`):
- `getClientsForField`
- `getMachinesByClient`
- `searchMachine`
- `saveVisit`
- `saveMachine`
- `saveClient`

**Não alterar o PCF** salvo necessidade indispensável.

---

# 9. Caso de teste obrigatório

## Dados iniciais

**Clientes:** `3844`, `Fundição Claudiense`

**Máquina:** Cliente `3844`, Série `3844`, Marca `Techto`, Modelo `SDI50HP`

## Teste A — Aba Clientes

1. Editar cliente `3844` → digitar `Fundição Claudiense` → alerta de similaridade aparece.
2. Clicar **"Mesclar com este"** → modal permanece aberto, carregado com `Fundição Claudiense`.
3. Salvar.

**Resultado esperado:**
- `3844` não aparece mais como cliente ativo.
- Existe apenas uma `Fundição Claudiense`.
- A máquina série `3844` fica vinculada a `Fundição Claudiense`.
- Contador de máquinas correto. "Cadastro incompleto" não aponta `3844`.

## Teste B — Plano de Preventivas

1. Abrir Plano de Preventivas → máquina série `3844` → editar dados.
2. Digitar `fund` no campo Razão Social.

**Resultado esperado:**
- Apenas uma opção `Fundição Claudiense` (sem duplicata por maiúscula/acento).
- Ao selecionar, campos cadastrais preenchem quando existirem.

## Teste C — Sync

1. Sincronizar → recarregar com Ctrl+Shift+R.

**Resultado esperado:**
- `3844` não volta como cliente.
- A máquina mantém cliente `Fundição Claudiense`.
- PCF ao buscar série `3844` retorna máquina com cliente correto.

---

# 10. Restrições absolutas

- Não reescrever a arquitetura.
- Não mudar nomes de chaves do localStorage (`proper_admin_v2`, `pgp_pending_sync`).
- Não apagar máquinas, visitas ou peças durante merge.
- Não alterar lógica de catálogo, preventivas ou peças fora do escopo acima.
- Não criar duplicata para resolver duplicata.
- `closeModal` / `closeModalSafe` **nunca** devem ser chamadas dentro de `selectDuplicateClientForMerge`.
- Logs com prefixos `[PCM client merge]` e `[GAS mergeClients]`.
- Antes de qualquer merge destrutivo local, criar snapshot via mecanismo já existente (`_captureFormSnapshot` ou similar).

---

# 11. Entrega

Ao finalizar, entregar:

1. Resumo das alterações por arquivo (funções alteradas/criadas).
2. Checklist dos testes A, B e C acima.
3. Aviso de qualquer limitação encontrada.
4. Versão dos arquivos entregues: `PCM_v13.1.html` e `GAS_properCare_v10.1.js` (ou conforme convenção do repositório).
