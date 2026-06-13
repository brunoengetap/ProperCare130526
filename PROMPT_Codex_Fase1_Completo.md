# PROMPT CODEX — ProperCare Fase 1 · Correção Completa
**Arquivos:** `GAS_properCare_v20.js` → v21 · `PCM_v29.html` → v30 · `Proper_Field_grupotap_v13.html` → v14  
**Deploy obrigatório:** GAS → PCM → PCF (nessa ordem)

---

## CONTEXTO

A Fase 1 do sistema de Ordens de Serviço ProperCare foi implementada mas tem bugs críticos e gaps arquiteturais que impedem o fluxo completo. Este prompt corrige **todos** os pontos pendentes para completar a Fase 1. O GAS/PCM/PCF foram **todos** auditados — nenhum arquivo deve ser considerado correto a priori.

---

## REGRAS ABSOLUTAS (não violar em nenhum ponto)

- **NÃO reordenar colunas** existentes em nenhum array de headers do GAS
- **NÃO modificar** as funções: `machineKey()`, `clientKey()`, `pgpFlushPending`, `pgpPostJson`, `pgpBuildVisitRequest`, `pgpBuildPartsUpdateRequest`, `pgpMachineKey()`, `pgpClientKey()`
- **NÃO alterar** a chave `proper_admin_v2` do localStorage
- **Adicionar campos** ao final dos arrays HEADERS (append-only, nunca inserir no meio)
- **Versionar** os arquivos: v20→v21, v29→v30, v13→v14

---

## BLOCO 1 — GAS v20 → v21

### 1.1 — Adicionar `tipo_os` ao schema `ORDENS_SERVICO`

Localizar o array `HEADERS.ORDENS_SERVICO` (atualmente termina em `'pdf_url'`). Adicionar `'tipo_os'` ao **final**:

```js
// ANTES
ORDENS_SERVICO: [
  'id_os','numero_os','id_cliente','cliente','descricao',
  'data_abertura','data_prevista','status',
  'tecnicos_vinculados','maquinas_vinculadas','id_visita_resultado',
  'drive_folder_id','drive_folder_url',
  'inicio_atendimento','fim_atendimento','pdf_url'
],

// DEPOIS
ORDENS_SERVICO: [
  'id_os','numero_os','id_cliente','cliente','descricao',
  'data_abertura','data_prevista','status',
  'tecnicos_vinculados','maquinas_vinculadas','id_visita_resultado',
  'drive_folder_id','drive_folder_url',
  'inicio_atendimento','fim_atendimento','pdf_url','tipo_os'
],
```

### 1.2 — Adicionar `tipo_os` no payload de criação de OS (`saveOS`)

Na função `saveOS(body)`, no bloco de criação de nova OS, incluir `tipo_os` no objeto de valores mapeados ao header. Valor padrão: `'Preventiva'` se não informado.

No bloco de **atualização** de OS existente, adicionar `'tipo_os'` ao array `allowed`:

```js
const allowed = ['status','maquinas_vinculadas','tecnicos_vinculados','descricao',
  'data_prevista','id_visita_resultado','inicio_atendimento','fim_atendimento',
  'pdf_url','drive_folder_id','drive_folder_url','tipo_os']; // ← adicionar tipo_os
```

### 1.3 — Adicionar endpoints `getTecnicos` e `saveTecnico`

No `switch(action)` do `doPost` / `doGet`, adicionar os cases:

```js
case 'getTecnicos':
  result = getTecnicos();
  break;

case 'saveTecnico':
  result = saveTecnico(body);
  break;
```

Implementar as funções:

```js
function getTecnicos() {
  ensureSheetHeaders('CAT_TECNICOS', HEADERS.CAT_TECNICOS);
  const rows = getSheetData('CAT_TECNICOS');
  return { status: 'ok', tecnicos: rows.filter(function(t){ return t.ativo !== 'false' && t.ativo !== false; }) };
}

function saveTecnico(body) {
  // body: { id, nome, pin, ativo }  — pin em texto plano, será armazenado como sha256 hex
  // Se id existente → atualiza; senão → cria
  ensureSheetHeaders('CAT_TECNICOS', HEADERS.CAT_TECNICOS);
  const sheet = getOrCreateSheet('CAT_TECNICOS', HEADERS.CAT_TECNICOS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const pinIdx = headers.indexOf('pin_hash');
  const nomeIdx = headers.indexOf('nome');
  const ativoIdx = headers.indexOf('ativo');
  const perfilIdx = headers.indexOf('perfil');

  // Hash SHA-256 via Utilities (GAS nativo)
  let pinHash = body.pin_hash || '';
  if (body.pin && !body.pin_hash) {
    const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(body.pin));
    pinHash = raw.map(function(b){ return (b < 0 ? b + 256 : b).toString(16).padStart(2,'0'); }).join('');
  }

  // Buscar linha existente pelo id
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(body.id || '')) {
      if (body.nome !== undefined) sheet.getRange(i+1, nomeIdx+1).setValue(body.nome);
      if (pinHash) sheet.getRange(i+1, pinIdx+1).setValue(pinHash);
      if (body.ativo !== undefined) sheet.getRange(i+1, ativoIdx+1).setValue(body.ativo);
      return { status: 'ok', action: 'updated' };
    }
  }

  // Novo técnico
  const newId = body.id || ('TEC-' + String(Date.now()).slice(-5));
  const row = headers.map(function(h) {
    if (h === 'id') return newId;
    if (h === 'nome') return body.nome || '';
    if (h === 'pin_hash') return pinHash;
    if (h === 'ativo') return body.ativo !== undefined ? body.ativo : true;
    if (h === 'perfil') return body.perfil || 'tecnico';
    return '';
  });
  sheet.appendRow(row);
  return { status: 'ok', action: 'created', id: newId };
}
```

### 1.4 — Adicionar `seedTecnicosDefault()`

Criar função utilitária para seed inicial (executar uma única vez manualmente via editor GAS):

```js
function seedTecnicosDefault() {
  // PIN padrão: 1234 → hash SHA-256
  saveTecnico({ id: 'TEC-001', nome: 'Técnico Padrão', pin: '1234', ativo: true, perfil: 'tecnico' });
  Logger.log('Seed concluído: TEC-001 / PIN 1234');
}
```

### Gate GAS

Após deploy, verificar no editor GAS:
1. Executar `seedTecnicosDefault()` → confirmar linha em `CAT_TECNICOS`
2. Chamar `?action=getTecnicos` → retorna JSON com array `tecnicos`
3. Chamar `?action=getOS` → retorna array (pode ser vazio)
4. Verificar aba `ORDENS_SERVICO` no Sheets → confirmar que a coluna `tipo_os` existe no cabeçalho

---

## BLOCO 2 — PCM v29 → v30

### 2.1 — CRÍTICO: Corrigir modal invisível de OS

**Bug:** os modais de OS são inseridos com `class="overlay show"`, mas o CSS só exibe overlays com classe `open`.

**Localizar e corrigir AMBAS as ocorrências:**

```js
// openOSForm — ANTES
const html=`<div class="overlay show" id="modalOS">...`;

// openOSForm — DEPOIS
const html=`<div class="overlay open" id="modalOS">...`;
```

```js
// openOSDetail — ANTES
const html=`<div class="overlay show" id="modalOSDetail">...`;

// openOSDetail — DEPOIS
const html=`<div class="overlay open" id="modalOSDetail">...`;
```

**Gate obrigatório:**
```bash
grep -n "overlay show" PCM_v30.html
# Deve retornar 0 ocorrências relacionadas a OS modals
```

### 2.2 — Substituir campo técnico por seletor dinâmico

**Remover** o campo de texto livre de técnicos do `openOSForm`. **Substituir** por um `<select multiple>` carregado a partir do endpoint `getTecnicos`:

A função `openOSForm` deve:
1. Fazer `gsGet('getTecnicos')` antes de montar o HTML
2. Construir opções `<option value="${t.id}">${t.nome}</option>` para cada técnico ativo
3. Exibir `<select id="osTecnicosSelect" multiple size="3">...</select>` com label "Técnicos"
4. Se `getTecnicos` falhar ou retornar vazio, exibir aviso `"Nenhum técnico cadastrado. Cadastre em Configurações > Técnicos."` e bloquear o botão Salvar

Em `saveOSFromPCM()`, capturar os técnicos do select:
```js
tecnicos_vinculados: [...document.querySelectorAll('#osTecnicosSelect option:checked')].map(o=>o.value)
```

E validar:
```js
if(!payload.tecnicos_vinculados.length) return toast('Selecione ao menos um técnico','warning');
```

### 2.3 — Adicionar campo `tipo_os` no formulário de criação

Adicionar `<select id="osTipo">` no `openOSForm`, com as opções:
- `Preventiva` (selecionado por padrão)
- `Corretiva`
- `Visita Técnica`
- `Diagnóstico`
- `Genérica`

Label: "Tipo de OS"

Em `saveOSFromPCM()`, incluir no payload:
```js
tipo_os: document.getElementById('osTipo')?.value || 'Preventiva'
```

### 2.4 — Filtrar máquinas por cliente no formulário de OS

Na função `openOSForm`:

1. Manter o campo `id="osCliente"` como input com datalist
2. Adicionar `onchange="osFilterMachines()"` ao input do cliente
3. Criar função `osFilterMachines()`:

```js
function osFilterMachines(){
  const clienteNorm = (document.getElementById('osCliente')?.value||'').trim().toLowerCase();
  document.querySelectorAll('.osMachineChk').forEach(function(chk){
    const clienteLabel = chk.dataset.client || '';
    const match = !clienteNorm || clienteLabel.toLowerCase().includes(clienteNorm);
    chk.closest('label').style.display = match ? 'block' : 'none';
  });
}
```

4. No HTML das checkboxes de máquinas, adicionar `data-client="${osEscape(m.client||'')}"` em cada `<input>`:
```html
<input type="checkbox" class="osMachineChk" value="${osEscape(m.id)}" data-client="${osEscape(m.client||'')}">
```

### 2.5 — Salvar `id_cliente` canonicamente

No objeto `payload` de `saveOSFromPCM()`, adicionar:
```js
id_cliente: (db.clients||[]).find(c=>(c.nome||'').toLowerCase().trim() === payload.cliente.toLowerCase().trim())?.id || ''
```

### 2.6 — Adicionar aba de gestão de Técnicos no PCM

Criar seção "Técnicos" acessível via menu ou botão de configuração. A seção deve:

1. Carregar e listar técnicos via `getTecnicos`
2. Exibir tabela: ID · Nome · Status (ativo/inativo)
3. Botão "Novo Técnico" → modal com campos: Nome, PIN (texto, obscurecido), Ativo (checkbox)
4. Ao salvar, chamar `syncToGS('saveTecnico', { nome, pin, ativo: true })`
5. Botão de toggle ativo/inativo por linha

**Campos do modal Novo Técnico:**
```html
<input id="tecNome" placeholder="Nome completo">
<input id="tecPin" type="password" placeholder="PIN (ex: 1234)" maxlength="8">
<label><input id="tecAtivo" type="checkbox" checked> Ativo</label>
```

**Payload ao salvar:**
```js
{ nome: tecNome.value.trim(), pin: tecPin.value, ativo: document.getElementById('tecAtivo').checked }
```

### Gate PCM — Modal de OS

1. Abrir PCM → ir em Ordens de Serviço
2. Clicar `+ Nova OS`
3. Modal deve abrir visualmente
4. Console: `getComputedStyle(document.getElementById('modalOS')).display` deve ser `flex`
5. Seletor de técnicos deve estar populado com ao menos TEC-001
6. Selecionar técnico, tipo, cliente, máquinas → Salvar
7. `loadOSList()` deve exibir a OS criada
8. Clicar "Abrir" → modal de detalhe deve abrir

---

## BLOCO 3 — PCF v13 → v14

### 3.1 — Bug #1: Capturar `visit_id` da resposta do GAS

Na função que envia a preventiva (dentro do bloco `try` após `const d=JSON.parse(txt)`):

**ANTES:**
```js
// (captura feita ANTES do fetch, usa visit_id local)
if(_pgpOSAtual) _pgpOSUltimaVisita = payload.visit_id || _pgpOSUltimaVisita;
```

**DEPOIS:** manter a captura local como fallback, e sobrescrever com a resposta do GAS se disponível:
```js
// Após d.status === 'ok':
if(_pgpOSAtual){
  _pgpOSUltimaVisita = d.visit_id || payload.visit_id || _pgpOSUltimaVisita;
}
```

A captura local antes do fetch (linha ~4905) deve ser removida ou mantida apenas como fallback de offline. A fonte de verdade é `d.visit_id` retornado pelo GAS.

### 3.2 — Bug #2: Não resetar tipo de visita após preventiva dentro de OS

Na linha que contém:
```js
if(!_pgpOSAtual) pgpSetTipoVisita('inspecao');
```

Confirmar que a condição `if(!_pgpOSAtual)` está presente. Se estiver ausente ou comentada, restaurar. Se estiver presente, ela está correta e protege o fluxo de OS.

Adicionalmente, após `d.status === 'ok'` dentro de OS, exibir badge informativo se necessário mas **não chamar `pgpSetTipoVisita`** em nenhuma hipótese quando `_pgpOSAtual !== null`.

### 3.3 — Bug #3: Garantir que `_pgpOSUltimaVisita` seja propagado corretamente ao concluir OS

Em `pgpConcluirOS()`, confirmar que o `saveOS` usa `_pgpOSUltimaVisita`:
```js
await pgpPostJson('saveOS',{
  id_os: _pgpOSAtual.id_os,
  status: 'concluida',
  fim_atendimento: fim,
  id_visita_resultado: _pgpOSUltimaVisita || ''
});
```

Se `_pgpOSUltimaVisita` estiver vazio ao concluir, exibir aviso mas **não bloquear** a conclusão:
```js
if(!_pgpOSUltimaVisita) showNotification('⚠ Nenhuma preventiva registrada nesta OS — concluindo sem vínculo de visita.');
```

### 3.4 — Bug #4: `maquinas_vinculadas` não deve ser sobrescrito na conclusão

Confirmar que `pgpConcluirOS()` **não** inclui `maquinas_vinculadas` no payload do `saveOS` final. O campo já foi salvo em `pgpAplicarMaquinasOS()` e não deve ser relido de checkboxes que podem não existir mais no DOM.

Se houver qualquer linha em `pgpConcluirOS` que leia `.pgpOSMachineChk`, **remover**.

### 3.5 — Desabilitar seleção manual de tipo dentro de OS ativa

Criar função `pgpSetOSModeUI(ativo)` que desabilita/habilita os botões de tipo:

```js
function pgpSetOSModeUI(ativo) {
  const btns = ['btnTipoInspecao','btnTipoPreventiva'];
  btns.forEach(function(id){
    const el = document.getElementById(id);
    if(!el) return;
    el.disabled = ativo;
    el.style.opacity = ativo ? '0.4' : '';
    el.style.pointerEvents = ativo ? 'none' : '';
    el.title = ativo ? 'Tipo definido pela OS' : '';
  });
  // Badge informativo
  const badge = document.getElementById('pgpOSBadge');
  if(badge) badge.style.display = ativo ? 'inline-block' : 'none';
}
```

Chamar `pgpSetOSModeUI(true)` dentro de `pgpAbrirOS()`, após `pgpSetTipoVisita('preventiva')`.
Chamar `pgpSetOSModeUI(false)` dentro de `pgpConcluirOS()` (após saveOS) e em qualquer função de cancelamento de OS.

### 3.6 — Badge informativo de OS ativa

Após o grupo de botões de tipo de visita, adicionar um elemento inicialmente oculto:

```html
<div id="pgpOSBadge" style="display:none;background:var(--blue);color:white;
  border-radius:8px;padding:4px 10px;font-size:12px;font-weight:700;margin:6px 0">
  🔧 Modo OS — tipo definido pelo PCM
</div>
```

Ao abrir uma OS, atualizar o texto do badge com o número e tipo:
```js
const badge = document.getElementById('pgpOSBadge');
if(badge){
  badge.textContent = `🔧 ${_pgpOSAtual.tipo_os || 'Preventiva'} · ${_pgpOSAtual.numero_os}`;
  badge.style.display = 'inline-block';
}
```

### 3.7 — Separar visualmente "Tipo de Visita" de "Abrir OS"

**Atual:** os três botões Inspeção / Preventiva / Ordens de Serviço estão no mesmo grupo `.pgp-tipo-btn`.

**Novo layout:**

```html
<!-- SEÇÃO 1: visitas livres (sem OS) -->
<div class="pgp-tipo-grupo" id="pgpGrupoTipoVisita">
  <div style="font-size:11px;font-weight:700;color:var(--text-light);margin-bottom:4px">TIPO DE VISITA</div>
  <button id="btnTipoInspecao" class="pgp-tipo-btn" onclick="pgpSetTipoVisita('inspecao')">
    🔍 Inspeção
  </button>
  <button id="btnTipoPreventiva" class="pgp-tipo-btn" onclick="pgpSetTipoVisita('preventiva')">
    🔧 Preventiva
  </button>
</div>

<!-- Badge OS ativa (oculto por padrão) -->
<div id="pgpOSBadge" style="display:none; ..."></div>

<!-- SEÇÃO 2: modo OS (separado visualmente) -->
<div style="margin-top:8px;padding-top:8px;border-top:1px solid var(--gray-mid)">
  <button id="btnTipoOS" class="btn" style="width:100%;background:var(--blue);color:white;border-radius:8px"
          onclick="pgpOpenOSLogin()">
    🧾 Abrir Ordem de Serviço
  </button>
</div>
```

### 3.8 — Propagar `tipo_os` da OS para o payload da visita

Em `pgpAbrirOS(i)`, após `pgpSetTipoVisita('preventiva')`, sobrescrever com o tipo real da OS se disponível:

```js
const tipoOS = _pgpOSAtual.tipo_os || 'preventiva';
// mapear para o tipo interno do PCF
const tipoInterno = tipoOS.toLowerCase() === 'corretiva' ? 'preventiva' : 'preventiva';
// (Fase 1: todos executam como preventiva no campo; tipo_os é informativo)
// Exibir o tipo real no badge — não forçar lógica diferente por ora
```

Em `pgpBuildPreventivaPayload` (ou onde `tipo_visita` é definido no payload), adicionar:
```js
tipo_visita: _pgpOSAtual ? (_pgpOSAtual.tipo_os || 'Preventiva') : _pgpTipoVisita
```

> **Nota Fase 1:** para simplificar, todos os tipos de OS continuam usando o formulário de preventiva. O `tipo_os` é gravado como informação no payload de visita. A lógica de formulários distintos por tipo é Fase 2.

### Gate PCF

1. Abrir PCF
2. Verificar que "Inspeção" e "Preventiva" estão em seção separada de "Abrir Ordem de Serviço"
3. Clicar "Abrir Ordem de Serviço" → tela de PIN
4. Digitar PIN do TEC-001 (1234) → deve listar OS criada no PCM
5. Abrir OS → botões Inspeção/Preventiva devem ficar desabilitados + badge visível
6. Selecionar máquinas → "Usar máquinas selecionadas"
7. Executar preventiva → registrar
8. `_pgpOSUltimaVisita` no console deve ter valor não vazio
9. "Fechar OS" → assinar → "Concluir OS"
10. PCM: abrir OS → status deve ser `concluida` e `id_visita_resultado` preenchido

---

## BLOCO 4 — Teste E2E obrigatório após todos os deploys

Executar o ciclo completo na ordem:

| # | Passo | Resultado esperado |
|---|-------|--------------------|
| 1 | GAS: executar `seedTecnicosDefault()` | Log confirma TEC-001 criado |
| 2 | GAS: `?action=getTecnicos` | JSON com TEC-001 ativo |
| 3 | PCM: `+ Nova OS` | Modal abre (display: flex) |
| 4 | PCM: selecionar técnico TEC-001, tipo "Preventiva", cliente, máquinas | Formulário aceita |
| 5 | PCM: Salvar OS | Toast "OS salva", lista atualiza |
| 6 | GAS: `?action=getOS` | OS aparece com `tipo_os: "Preventiva"` |
| 7 | PCF: "Abrir Ordem de Serviço" | Tela de PIN |
| 8 | PCF: PIN 1234 | Lista OS do TEC-001 |
| 9 | PCF: abrir OS | Máquinas pré-selecionadas, botões tipo desabilitados, badge visível |
| 10 | PCF: registrar preventiva | `_pgpOSUltimaVisita` não vazio |
| 11 | PCF: fechar OS, assinar, concluir | Notificação de sucesso |
| 12 | PCM: abrir OS | Status `concluida`, `id_visita_resultado` preenchido |

---

## RESUMO DE MUDANÇAS POR ARQUIVO

### GAS v21
- `HEADERS.ORDENS_SERVICO`: campo `tipo_os` adicionado ao final
- `saveOS()`: `tipo_os` no payload de criação e no array `allowed` de atualização
- Switch `doPost/doGet`: cases `getTecnicos` e `saveTecnico`
- Novas funções: `getTecnicos()`, `saveTecnico()`, `seedTecnicosDefault()`

### PCM v30
- `openOSForm()`: `overlay show` → `overlay open` (**crítico**)
- `openOSDetail()`: `overlay show` → `overlay open` (**crítico**)
- `openOSForm()`: campo técnico → `<select multiple>` dinâmico via `getTecnicos`
- `openOSForm()`: `<select id="osTipo">` com 5 opções
- `openOSForm()`: máquinas com `data-client` + `osFilterMachines()` on change
- `saveOSFromPCM()`: inclui `tipo_os` e `id_cliente` no payload
- `saveOSFromPCM()`: valida técnico selecionado
- Nova seção "Técnicos" no menu com CRUD via `getTecnicos`/`saveTecnico`

### PCF v14
- Bug #1: `_pgpOSUltimaVisita` captura `d.visit_id` da resposta GAS
- Bug #2: `pgpSetTipoVisita('inspecao')` protegido por `if(!_pgpOSAtual)`
- Bug #3: `pgpConcluirOS()` aviso se `_pgpOSUltimaVisita` vazio, não bloqueia
- Bug #4: `pgpConcluirOS()` não lê checkboxes de máquinas
- Nova função `pgpSetOSModeUI(ativo)` desabilita botões de tipo durante OS
- `pgpAbrirOS()`: chama `pgpSetOSModeUI(true)` + atualiza badge
- `pgpConcluirOS()`: chama `pgpSetOSModeUI(false)` ao final
- Novo `<div id="pgpOSBadge">` oculto, visível ao entrar em OS
- Layout: separação visual "TIPO DE VISITA" vs "Abrir Ordem de Serviço"
- Payload da visita: inclui `tipo_os` da OS quando ativa

---

## O QUE NÃO ALTERAR

As seguintes funcionalidades estão corretas e **não devem ser tocadas**:
- Login por PIN com SHA-256 (SubtleCrypto) no PCF
- Numeração automática de OS pelo GAS
- Filtro de OS por técnico logado em `loginTecnico`
- Cronômetro `_pgpOSInicio` / `fim_atendimento`
- Assinaturas por canvas touch + upload via `pgpQueueAnexo`
- Schema `CAT_TECNICOS` no GAS (não alterar colunas existentes)
- Painel de OS no PCM: filtros de Status/Técnico/Cliente existentes
- Persistência de sessão de técnico em `sessionStorage`
- `pgpFlushAnexos()` e todo o pipeline de fila offline
