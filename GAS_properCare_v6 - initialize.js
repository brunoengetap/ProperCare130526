// ═══════════════════════════════════════════════════════════════════════════
// PROPERCARE — Google Apps Script Backend — PCM/PCF
// ═══════════════════════════════════════════════════════════════════════════
//
// ESTRUTURA DE ABAS DA PLANILHA:
//   MODELOS       → Catálogo de modelos de equipamentos
//   MAQUINAS      → Equipamentos cadastrados dos clientes (PGP)
//   CLIENTES      → Clientes cadastrados no admin
//   VISITAS       → Registro de cada visita/preventiva realizada
//   PECAS_LOG     → Detalhamento das peças por visita
//   MACHINE_PARTS → Último estado das peças por máquina
//   PARTS_MASTER  → Catálogo de peças por modelo
//   PART_SIMILARITIES → Referências similares por peça
//
// NOVIDADES v1.6 (2026-05-14):
//   - CLIENTES: novas colunas Contato, Filial, Próx.Visita, Últ.Visita
//   - getClientsForField: retorna clientsFull com dados cadastrais
//   - getMachinesByClient: enriquece resultado com dados do cliente
//   - saveClient: aceita contato, filial, nextVisit
//   - VISITAS: novas colunas Próx.Visita, Contato, CNPJ, Filial
//   - PECAS_LOG: novas colunas Intervalo_H, H.Rodadas, H.Restantes, Status, Valor_Mostrado, Tipo_Contador
//   - saveVisit: grava nextVisit/contact/cnpj em VISITAS e dados calculados em PECAS_LOG
//   - Todas as alterações são retrocompatíveis (colunas adicionadas ao final)
//
// NOVIDADES v1.4:
//   - Campos de auditoria em todas as tabelas (Ativo, Tipo_Registro, Created_At, Deleted_At, Deleted_By)
//   - Soft delete (Ativo = NÃO) em vez de exclusão física
//   - Hard delete restrito a Tipo_Registro = TESTE
//   - Schema CLIENTES alinhado com admin (cidade, drive_url, antecedencia_alerta_dias)
//   - Correção de bug de duplicidade em saveMachine (=== em vez de ==)
//   - getSheetDataActive filtra registros inativos nos GETs principais
//   - Suporte a includeDeleted=true nos GETs de MAQUINAS, MODELOS, CLIENTES
//
// IMPORTANTE: execute manualmente `migrateExistingIds()` 1x se vier da v1.3
// ═══════════════════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── MACHINE KEY — mesmo algoritmo do campo e admin ───────────────────────
function machineKey(client, brand, model, serial) {
  function norm(v) {
    return String(v || '').trim().toLowerCase()
      .replace(/[àáâãäå]/g,'a').replace(/[èéêë]/g,'e')
      .replace(/[ìíîï]/g,'i').replace(/[òóôõö]/g,'o')
      .replace(/[ùúûü]/g,'u').replace(/[ç]/g,'c')
      .replace(/[^a-z0-9]/g,'');
  }
  const parts = [norm(client), norm(brand), norm(model)];
  const ser = norm(serial);
  if (ser) parts.push(ser);
  return 'MK-' + parts.join('-');
}

// ── PROTEÇÃO POR TOKEN ────────────────────────────────────
const API_KEY = '8d969eef6ecad3c29a3a629280e686cf0c3f5d5a86aff3ca12020c923adc6c92'; // sha256 de 123456

function checkKey(params_or_body) {
  const k = params_or_body.key || params_or_body.k || '';
  if (k !== API_KEY) throw new Error('Acesso não autorizado');
}

// ── Cabeçalhos das abas ──────────────────────────────────────────────────
const HEADERS = {
  MAQUINAS: [
    'ID','Cliente','Filial','Marca','Modelo','Série','Ano','TAG','Localização',
    'Hor.Total','h/Semana','Observações',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Atualizado'
  ],
  MODELOS: [
    'ID','Marca','Modelo','Tipo','Potência','Pressão','Observações',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Atualizado'
  ],
  CLIENTES: [
    'ID','Nome','CNPJ','Cidade','Telefone','Email','Observações',
    'Antecedencia_Alerta_Dias','Drive_URL',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Atualizado',
    // v1.5 — novas colunas ao final (retrocompatível)
    'Contato','Filial','Prox_Visita','Ult_Visita'
  ],
  VISITAS: [
    'ID','Machine_ID','Cliente','Filial','Marca','Modelo','Série','TAG',
    'Hor.Visita','h/Semana','Cenário','Técnico','Data Visita','Tipo','Obs.Gerais',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Enviado',
    // v1.5 — novas colunas ao final
    'Prox_Visita','Contato','CNPJ'
  ],
  PECAS_LOG: [
    'ID_Visita','ID_Peça','Nome Peça','Ref.','Subsistema','Últ.Troca(h)',
    'N/A','Observação','Ref_Nova','Ref_Anterior','Tipo_Referencia','Acao',
    'Horimetro','Data_Troca',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By',
    // v1.5 — novas colunas ao final
    'Intervalo_H','H_Rodadas','H_Restantes','Status','Valor_Mostrado','Tipo_Contador'
  ],
  MACHINE_PARTS: [
    'Machine_ID','Serial','TAG','Part_ID','Part_Name','Last_Change_H',
    'Interval_H','Ref','NA','Ref_Anterior','Created_At',
    'Ativo','Tipo_Registro','Deleted_At','Deleted_By','Atualizado',
    'Valor_Mostrado','Contador'
  ],
  PARTS_MASTER: [
    'Part_ID','Model_ID','Name','OEM_Ref','Part_Brand','Supplier_Primary',
    'Slot','Qty_Default','Interval_H','Criticality','Cost','Obs',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Updated_At',
    'Part_Scope','Sub_ID','Sub_Name','Sub_Category','Sub_Desc','Sub_Interval_H'
  ],
  PART_SIMILARITIES: [
    'Sim_ID','Part_ID','Model_ID','Ref_Similar','Brand_Similar','Obs',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Updated_At'
  ]
};

// ── Valores padrão de auditoria para novas linhas ────────────────────────
function auditDefaults(tipoRegistro) {
  const now = new Date().toISOString();
  return {
    ativo: 'SIM',
    tipoRegistro: tipoRegistro || 'PRODUCAO',
    createdAt: now,
    deletedAt: '',
    deletedBy: ''
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// ENTRY POINT — GET (ping e consultas)
// ═══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  const params = e.parameter;
  const action = params.action || 'ping';
  let result;

  try {
    if (action !== 'ping') checkKey(params);
    switch (action) {
      case 'ping':
        result = { status: 'ok', version: '1.4', spreadsheet: SS.getName(), ts: new Date().toISOString() };
        break;
      case 'searchMachine':
        result = searchMachine(params.q || '');
        break;
      case 'getMachines': {
        const inclDel = params.includeDeleted === 'true';
        result = { status: 'ok', machines: inclDel ? getSheetData('MAQUINAS') : getSheetDataActive('MAQUINAS') };
        break;
      }
      case 'getMachinesWithParts':
        result = getMachinesWithParts();
        break;
      case 'getModels': {
        const inclDel = params.includeDeleted === 'true';
        result = { status: 'ok', models: inclDel ? getSheetData('MODELOS') : getSheetDataActive('MODELOS') };
        break;
      }
      case 'getVisits':
        result = { status: 'ok', visits: getVisitsNormalized() };
        break;
      case 'getMachineParts':
        result = { status: 'ok', parts: getSheetData('MACHINE_PARTS') };
        break;
      case 'getAllMachineParts':
        result = { status: 'ok', parts: getSheetData('MACHINE_PARTS') };
        break;
      case 'getClients': {
        const inclDel = params.includeDeleted === 'true';
        result = { status: 'ok', clients: inclDel ? getSheetData('CLIENTES') : getSheetDataActive('CLIENTES') };
        break;
      }
      case 'getPartsMaster':
        result = { status: 'ok', parts: getSheetDataActive('PARTS_MASTER') };
        break;
      case 'getPartSimilarities':
        result = { status: 'ok', similarities: getSheetData('PART_SIMILARITIES') };
        break;
      case 'getCatalogFull':
        result = getCatalogFull();
        break;
      case 'getVisitsByMachine':
        result = getVisitsByMachine(params.machine_id || '');
        break;
      case 'getMachinesByClient':
        result = getMachinesByClient(params.client || '');
        break;
      case 'getClientsForField':
        result = getClientsForField();
        break;
      case 'getSystemHealth':
        result = getSystemHealth();
        break;
      case 'getDuplicates':
        result = getDuplicates();
        break;
      case 'getDeletedRecords':
        result = getDeletedRecords(params.sheetName || '');
        break;
      case 'createBackupSnapshot':
        result = createBackupSnapshot();
        break;
      default:
        result = { status: 'error', error: 'Ação desconhecida: ' + action };
    }
  } catch (err) {
    result = { status: 'error', error: err.message };
  }

  return jsonResponse(result);
}

// ═══════════════════════════════════════════════════════════════════════════
// ENTRY POINT — POST (gravações)
// ═══════════════════════════════════════════════════════════════════════════
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ status: 'error', error: 'JSON inválido' }); }

  const action = body.action;
  let result;

  try {
    checkKey(body);
    switch (action) {
      case 'saveVisit':
        result = saveVisit(body);
        break;
      case 'savePreventiva':
        result = savePreventiva(body);
        break;
      case 'saveMachine':
        result = saveMachine({ ...body.machine, tipoRegistro: body.machine?.tipoRegistro || 'PRODUCAO' });
        break;
      case 'deleteMachine':
        result = softDelete('MAQUINAS', body.id, body.deletedBy || 'admin');
        break;
      case 'saveModel':
        result = saveModel(body.model);
        break;
      case 'deleteModel':
        result = softDelete('MODELOS', body.id, body.deletedBy || 'admin');
        break;
      case 'updateMachineParts':
        result = updateMachineParts(body);
        break;
      case 'savePartMaster':
        result = savePartMaster(body.part);
        break;
      case 'replacePartSimilarities':
        result = replacePartSimilarities(body.partId, body.modelId, body.similarities || []);
        break;
      case 'saveClient':
        result = saveClient(body.client);
        break;
      case 'deleteClient':
        result = softDelete('CLIENTES', body.id, body.deletedBy || 'admin');
        break;
      case 'hardDeleteTestRecord':
        result = hardDeleteTestRecord(body.sheetName, body.id);
        break;
      default:
        result = { status: 'error', error: 'Ação desconhecida: ' + action };
    }
  } catch (err) {
    result = { status: 'error', error: err.message };
  }

  return jsonResponse(result);
}

// ═══════════════════════════════════════════════════════════════════════════
// SEARCH MACHINE
// ═══════════════════════════════════════════════════════════════════════════
function searchMachine(query) {
  if (!query) return { status: 'error', error: 'Query vazia' };
  const sheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { status: 'ok', machine: null };

  const headers = data[0];
  const qLower = query.toLowerCase().trim();

  const idxSerie  = headers.indexOf('Série');
  const idxTag    = headers.indexOf('TAG');
  const idxClient = headers.indexOf('Cliente');
  const idxId     = headers.indexOf('ID');
  const idxAtivo  = headers.indexOf('Ativo');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Ignorar registros inativos na busca
    const ativo = String(row[idxAtivo] || 'SIM').toUpperCase();
    if (ativo === 'NÃO') continue;

    const serie  = String(row[idxSerie]  || '').toLowerCase();
    const tag    = String(row[idxTag]    || '').toLowerCase();
    const client = String(row[idxClient] || '').toLowerCase();
    const id     = String(row[idxId]     || '').toLowerCase();
    const brand  = String(row[headers.indexOf('Marca')] || '').toLowerCase();
    const model  = String(row[headers.indexOf('Modelo')]|| '').toLowerCase();

    if (serie === qLower || tag === qLower || id === qLower ||
        client.includes(qLower) || (brand+' '+model).includes(qLower)) {
      const machine = {};
      headers.forEach((h, j) => machine[h] = row[j]);
      const result = rowToMachine(machine);
      result.parts = getMachinePartsById(result.id, result.serial, result.tag);
      return { status: 'ok', machine: result };
    }
  }
  return { status: 'ok', machine: null };
}

// ═══════════════════════════════════════════════════════════════════════════
// SAVE VISIT (from field checklist)
// ═══════════════════════════════════════════════════════════════════════════
function saveVisit(body) {
  ensureSheetHeaders('VISITAS', HEADERS.VISITAS);
  ensureSheetHeaders('PECAS_LOG', HEADERS.PECAS_LOG);
  const visitId = String(body.visit_id || ('VIS-' + new Date().getTime())).trim();
  const now = new Date().toISOString();
  const visitDate = body.visitDate || new Date().toLocaleDateString('pt-BR');
  const machineId = body.machine_id ||
    machineKey(body.client||'', body.brand||'', body.model||'', body.serial||'');

  // Verificar idempotência: se visit_id já existe, retornar sem duplicar
  const visitSheet = getOrCreateSheet('VISITAS', HEADERS.VISITAS);
  const visitsData_check = visitSheet.getDataRange().getValues();
  const visitHeaders_check = visitsData_check[0] || HEADERS.VISITAS;
  const idxVid_check = visitHeaders_check.indexOf('ID');
  if (idxVid_check >= 0) {
    for (let i = 1; i < visitsData_check.length; i++) {
      if (String(visitsData_check[i][idxVid_check] || '').trim() === visitId) {
        return { status: 'ok', visitId, duplicate: true };
      }
    }
  }

  ensureMachineFromVisit(body, machineId, visitDate, now);

  visitSheet.appendRow([
    visitId, machineId,
    body.client   || '', body.branch   || '',
    body.brand    || '', body.model    || '',
    body.serial   || '', body.tag      || '',
    parseInt(body.hourTotal) || 0,
    parseInt(body.hpw)       || 0,
    body.scenario || '',
    body.tech     || '',
    visitDate,
    body.tipo || 'inspecao',
    body.generalObs || '',
    'SIM', 'PRODUCAO', now, '', '', now,
    // v1.5: novas colunas
    body.nextVisit || '', body.contact || '', body.cnpj || ''
  ]);

  const partsSheet = getOrCreateSheet('PECAS_LOG', HEADERS.PECAS_LOG);
  const parts = body.parts || {};
  Object.entries(parts).forEach(([partId, ps]) => {
    partsSheet.appendRow([
      visitId, partId,
      ps.name  || partId,
      ps.ref   || '',
      ps.sub   || '',
      parseInt(ps.lastChange) || 0,
      ps.na ? 'SIM' : 'NÃO',
      ps.obs   || '',
      '', '', '', '',
      parseInt(body.hourTotal) || 0,
      now,
      'SIM', 'PRODUCAO', now, '', '',
      // v1.5: novas colunas calculadas
      parseInt(ps.interval)      || 0,
      ps.horasRodadas            || '',
      ps.horasRestantes          || '',
      ps.status                  || '',
      ps.valorMostrado           || '',
      ps.contador                || ''
    ]);
  });

  // MACHINE_PARTS só é mutada por preventiva, nunca por inspeção
  const tipoVisita = String(body.tipo || 'inspecao').trim().toLowerCase();
  if (tipoVisita === 'preventiva') {
    updateMachineParts({
      machine_id: machineId,
      serial: body.serial || '',
      tag: body.tag || '',
      parts: parts
    });
  }

  return { status: 'ok', visitId, machineId };
}

// ═══════════════════════════════════════════════════════════════════════════
// UPDATE MACHINE PARTS
// ═══════════════════════════════════════════════════════════════════════════
function updateMachineParts(body) {
  ensureSheetHeaders('MACHINE_PARTS', HEADERS.MACHINE_PARTS);
  const machineId = String(body.machine_id || '').trim();
  const serial    = String(body.serial     || '').trim();
  const tag       = String(body.tag        || '').trim();
  const parts     = body.parts || {};

  if (!machineId && !serial && !tag) {
    return { status: 'error', error: 'machine_id, serial ou tag obrigatório' };
  }

  const sheet = getOrCreateSheet('MACHINE_PARTS', HEADERS.MACHINE_PARTS);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxMid  = headers.indexOf('Machine_ID');
  const idxSer  = headers.indexOf('Serial');
  const idxTag  = headers.indexOf('TAG');
  const idxPid  = headers.indexOf('Part_ID');
  const idxRef  = headers.indexOf('Ref');
  const idxRefAnterior = headers.indexOf('Ref_Anterior');
  const idxCreatedAt   = headers.indexOf('Created_At');
  const now     = new Date().toISOString();

  Object.entries(parts).forEach(([partId, ps]) => {
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      const rowMid = String(data[i][idxMid] || '').trim();
      const rowSer = String(data[i][idxSer] || '').trim();
      const rowTag = String(data[i][idxTag] || '').trim();
      const rowPid = String(data[i][idxPid] || '').trim();

      const machineMatch = (machineId && rowMid === machineId) ||
                           (serial    && rowSer === serial)    ||
                           (tag       && rowTag === tag);
      if (machineMatch && rowPid === partId) {
        rowIdx = i + 1;
        break;
      }
    }

    if (rowIdx > 0) {
      const existingRef = String(sheet.getRange(rowIdx, idxRef + 1).getValue() || '');
      const incomingRef = String(ps.ref || '');
      const refAnterior = (incomingRef && existingRef && incomingRef !== existingRef)
        ? existingRef
        : String(sheet.getRange(rowIdx, idxRefAnterior + 1).getValue() || '');
      const createdAt = String(sheet.getRange(rowIdx, idxCreatedAt + 1).getValue() || '');
      const rowData = [
        machineId, serial, tag,
        partId,
        ps.name     || partId,
        parseInt(ps.lastChange) || 0,
        parseInt(ps.interval)   || 2000,
        incomingRef,
        ps.na  ? 'SIM' : 'NÃO',
        refAnterior, createdAt,
        'SIM', 'PRODUCAO', '', '', now,
        ps.valorMostrado || '',
        ps.contador      || ''
      ];
      sheet.getRange(rowIdx, 1, 1, HEADERS.MACHINE_PARTS.length).setValues([rowData]);
    } else {
      const rowData = [
        machineId, serial, tag,
        partId,
        ps.name     || partId,
        parseInt(ps.lastChange) || 0,
        parseInt(ps.interval)   || 2000,
        ps.ref || '',
        ps.na  ? 'SIM' : 'NÃO',
        '', now,
        'SIM', 'PRODUCAO', '', '', now,
        ps.valorMostrado || '',
        ps.contador      || ''
      ];
      sheet.appendRow(rowData);
    }
  });

  return { status: 'ok', updated: Object.keys(parts).length };
}

function enrichMachineWithClientData(machine, clientData) {
  if (!machine || !clientData) return machine;
  const out = Object.assign({}, machine);
  const map = [['client','nome'],['branch','filial'],['cnpj','cnpj'],['contact','contato'],['phone','telefone'],['email','email'],['city','cidade'],['lastVisit','lastVisit'],['nextVisit','nextVisit']];
  map.forEach(([k,ck])=>{ if(!String(out[k]||'').trim() && String(clientData[ck]||'').trim()) out[k]=String(clientData[ck]).trim(); });
  out.clientData = clientData;
  return out;
}

// ═══════════════════════════════════════════════════════════════════════════
// GET MACHINES WITH PARTS
// ═══════════════════════════════════════════════════════════════════════════
function getMachinesWithParts() {
  const clients = getSheetDataActive('CLIENTES');
  const norm = v => String(v || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  const cmap = {};
  clients.forEach(c=>{ const k=norm(c['Nome']); if(!k) return; cmap[k]={nome:String(c['Nome']||'').trim(),cnpj:String(c['CNPJ']||'').trim(),cidade:String(c['Cidade']||'').trim(),telefone:String(c['Telefone']||'').trim(),email:String(c['Email']||'').trim(),contato:String(c['Contato']||'').trim(),filial:String(c['Filial']||'').trim(),nextVisit:String(c['Prox_Visita']||'').trim(),lastVisit:String(c['Ult_Visita']||'').trim()}; });
  const machines = getSheetDataActive('MAQUINAS').map(row => {
    let m = rowToMachineFromObj(row);
    m.parts = getMachinePartsById(m.id, m.serial, m.tag);
    m = enrichMachineWithClientData(m, cmap[norm(m.client)] || null);
    return m;
  });
  return { status: 'ok', machines };
}

// ═══════════════════════════════════════════════════════════════════════════
// GET MACHINE PARTS BY ID/SERIAL/TAG
// ═══════════════════════════════════════════════════════════════════════════
function getMachinePartsById(machineId, serial, tag) {
  try {
    const sheet = SS.getSheetByName('MACHINE_PARTS');
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {};

    const headers = data[0];
    const idxMid = headers.indexOf('Machine_ID');
    const idxSer = headers.indexOf('Serial');
    const idxTag = headers.indexOf('TAG');
    const idxPid = headers.indexOf('Part_ID');
    const idxName= headers.indexOf('Part_Name');
    const idxLch = headers.indexOf('Last_Change_H');
    const idxInt = headers.indexOf('Interval_H');
    const idxRef = headers.indexOf('Ref');
    const idxNA  = headers.indexOf('NA');
    const idxVM  = headers.indexOf('Valor_Mostrado');
    const idxCnt = headers.indexOf('Contador');

    const result = {};
    for (let i = 1; i < data.length; i++) {
      const rowMid = String(data[i][idxMid] || '').trim();
      const rowSer = String(data[i][idxSer] || '').trim();
      const rowTag = String(data[i][idxTag] || '').trim();
      const rowPid = String(data[i][idxPid] || '').trim();

      const match = (machineId && rowMid === String(machineId).trim()) ||
                    (serial    && rowSer === String(serial).trim())    ||
                    (tag       && rowTag === String(tag).trim());
      if (match && rowPid) {
        result[rowPid] = {
          name:          String(data[i][idxName] || rowPid),
          lastChange:    parseInt(data[i][idxLch]) || 0,
          interval:      parseInt(data[i][idxInt]) || 2000,
          ref:           String(data[i][idxRef] || ''),
          na:            String(data[i][idxNA] || '').toUpperCase() === 'SIM',
          valorMostrado: idxVM  >= 0 ? String(data[i][idxVM]  || '') : '',
          contador:      idxCnt >= 0 ? String(data[i][idxCnt] || '') : ''
        };
      }
    }
    return result;
  } catch(e) {
    return {};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// GET VISITS NORMALIZED
// ═══════════════════════════════════════════════════════════════════════════
function getVisitsNormalized() {
  ensureSheetHeaders('VISITAS', HEADERS.VISITAS);
  const sheet = getOrCreateSheet('VISITAS', HEADERS.VISITAS);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];

  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, j) => obj[h] = row[j]);
    return {
      visitId:    obj['ID']           || '',
      machine_id: obj['Machine_ID']   || '',
      client:     obj['Cliente']      || '',
      branch:     obj['Filial']       || '',
      brand:      obj['Marca']        || '',
      model:      obj['Modelo']       || '',
      serial:     obj['Série']        || '',
      tag:        obj['TAG']          || '',
      hourTotal:  parseInt(obj['Hor.Visita']) || 0,
      hpw:        parseInt(obj['h/Semana'])   || 0,
      scenario:   obj['Cenário']      || '',
      tech:       obj['Técnico']      || '',
      visitDate:  obj['Data Visita']  || '',
      tipo:       obj['Tipo']         || 'inspecao',
      generalObs: obj['Obs.Gerais']   || '',
      // v1.5:
      nextVisit:  obj['Prox_Visita']  || '',
      contact:    obj['Contato']      || '',
      cnpj:       obj['CNPJ']         || '',
      'Machine_ID':  obj['Machine_ID']  || '',
      'Série':       obj['Série']       || '',
      'TAG':         obj['TAG']         || '',
      'Hor.Total':   parseInt(obj['Hor.Visita']) || 0,
      'h/Semana':    parseInt(obj['h/Semana'])   || 0,
      'Data':        obj['Data Visita'] || '',
      'Visit_Date':  obj['Data Visita'] || '',
    };
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SAVE PREVENTIVA
// ═══════════════════════════════════════════════════════════════════════════
function savePreventiva(body) {
  ensureSheetHeaders('VISITAS', HEADERS.VISITAS);
  ensureSheetHeaders('PECAS_LOG', HEADERS.PECAS_LOG);
  ensureSheetHeaders('MACHINE_PARTS', HEADERS.MACHINE_PARTS);

  if (!body || !body.visitDate || body.hourTotal === undefined || body.hourTotal === null || !body.parts) {
    return { status: 'error', error: 'Campos obrigatórios: machine_id/tipo/visitDate/hourTotal/parts' };
  }

  const parts = body.parts || {};
  const validAcoes = { trocada: true, conferida: true, na: true };
  for (const [partId, ps] of Object.entries(parts)) {
    const acao = String((ps && ps.acao) || '').trim().toLowerCase();
    const partName = (ps && ps.name) || partId;
    if (!validAcoes[acao]) {
      return { status: 'error', error: 'Peça ' + partName + ': acao inválida' };
    }
    if (acao === 'trocada' && !String((ps && ps.ref) || '').trim()) {
      return { status: 'error', error: 'Peça ' + partName + ': ref obrigatória para acao=trocada' };
    }
  }

  const visitSheet = getOrCreateSheet('VISITAS', HEADERS.VISITAS);
  const visitsData = visitSheet.getDataRange().getValues();
  const visitHeaders = visitsData[0] || HEADERS.VISITAS;
  const idxVisitId = visitHeaders.indexOf('ID');
  const visitId = String(body.visit_id || ('VIS-' + new Date().getTime() + '-' + Math.floor(Math.random() * 100000))).trim();

  for (let i = 1; i < visitsData.length; i++) {
    if (String(visitsData[i][idxVisitId] || '').trim() === visitId) {
      return { status: 'ok', visitId, duplicate: true };
    }
  }

  let machineId = String(body.machine_id || '').trim();
  if (!machineId || !machineId.startsWith('MK-')) {
    machineId = machineKey(body.client || '', body.brand || '', body.model || '', body.serial || '');
  }

  const now = new Date().toISOString();
  ensureMachineFromVisit(body, machineId, body.visitDate, now);

  visitSheet.appendRow([
    visitId, machineId,
    body.client   || '', body.branch   || '',
    body.brand    || '', body.model    || '',
    body.serial   || '', body.tag      || '',
    parseInt(body.hourTotal) || 0,
    parseInt(body.hpw)       || 0,
    body.scenario || '',
    body.tech     || '',
    body.visitDate || now,
    body.tipo || 'preventiva',
    body.generalObs || '',
    'SIM', 'PRODUCAO', now, '', '', now,
    // v1.5:
    body.nextVisit || '', body.contact || '', body.cnpj || ''
  ]);

  const partsSheet = getOrCreateSheet('PECAS_LOG', HEADERS.PECAS_LOG);
  let pecasTrocadas = 0;
  Object.entries(parts).forEach(([partId, ps]) => {
    const acao = String(ps.acao || '').trim().toLowerCase();
    const refNova = String(ps.ref || '').trim();
    const refAnterior = String(ps.refAnterior || '').trim() ||
      getCurrentRefFromMachineParts(machineId, body.serial || '', body.tag || '', partId);

    partsSheet.appendRow([
      visitId, partId,
      ps.name || partId,
      refNova || refAnterior || '',
      ps.sub || '',
      parseInt(ps.lastChange) || 0,
      ps.na ? 'SIM' : 'NÃO',
      ps.obs || '',
      refNova, refAnterior,
      ps.tipoReferencia || '',
      acao,
      parseInt(body.hourTotal) || 0,
      now,
      'SIM', 'PRODUCAO', now, '', '',
      // v1.5:
      parseInt(ps.interval)     || 0,
      ps.horasRodadas           || '',
      ps.horasRestantes         || '',
      ps.status                 || '',
      ps.valorMostrado          || '',
      ps.contador               || ''
    ]);

    if (acao === 'trocada') {
      updateMachinePartFromPreventiva({
        machine_id: machineId,
        serial: body.serial || '',
        tag: body.tag || '',
        partId,
        name: ps.name || partId,
        interval: parseInt(ps.interval) || 2000,
        refNova: refNova,
        refAnterior,
        na: ps.na ? 'SIM' : 'NÃO',
        hourTotal: parseInt(body.hourTotal) || 0,
        now
      });
      pecasTrocadas++;
    } else if (acao === 'conferida' || acao === 'na') {
      // Não altera lastChange nem ref, mas persiste valorMostrado e contador
      updateMachinePartValorMostrado(machineId, body.serial || '', body.tag || '', partId, ps.valorMostrado || '', ps.contador || '', now);
    }
  });

  updateMachineHorímetro(machineId, body.client || '', body.serial || '', body.tag || '', parseInt(body.hourTotal) || 0, body.visitDate || now);
  return { status: 'ok', visitId, machineId, pecasTrocadas, duplicate: false };
}

// ═══════════════════════════════════════════════════════════════════════════
// GET VISITS BY MACHINE
// ═══════════════════════════════════════════════════════════════════════════
function getVisitsByMachine(machineId) {
  const mId = String(machineId || '').trim();
  if (!mId) return { status: 'error', error: 'machine_id obrigatório' };
  ensureSheetHeaders('VISITAS', HEADERS.VISITAS);
  ensureSheetHeaders('PECAS_LOG', HEADERS.PECAS_LOG);

  const allVisits = getSheetData('VISITAS');
  const visits = allVisits
    .filter(v => String(v['Machine_ID'] || '').trim() === mId)
    .map(v => ({
      visitId:    v['ID'] || '',
      machine_id: v['Machine_ID'] || '',
      tipo:       v['Tipo'] || 'inspecao',
      visitDate:  v['Data Visita'] || '',
      hourTotal:  parseInt(v['Hor.Visita']) || 0,
      tech:       v['Técnico'] || '',
      scenario:   v['Cenário'] || '',
      generalObs: v['Obs.Gerais'] || '',
      client:     v['Cliente'] || '',
      brand:      v['Marca'] || '',
      model:      v['Modelo'] || '',
      // v1.5:
      nextVisit:  v['Prox_Visita'] || '',
      contact:    v['Contato'] || '',
      cnpj:       v['CNPJ'] || '',
    }));

  visits.sort((a, b) => new Date(b.visitDate) - new Date(a.visitDate));
  const allPecasLog = getSheetData('PECAS_LOG');
  const visitIds = new Set(visits.map(v => String(v.visitId || '').trim()));
  const pecasLog = {};
  allPecasLog.forEach(p => {
    const vid = String(p['ID_Visita'] || '').trim();
    if (!visitIds.has(vid)) return;
    if (!pecasLog[vid]) pecasLog[vid] = [];
    pecasLog[vid].push({
      partId:         p['ID_Peça'] || '',
      name:           p['Nome Peça'] || '',
      acao:           p['Acao'] || '',
      refNova:        p['Ref_Nova'] || p['Ref.'] || '',
      refAnterior:    p['Ref_Anterior'] || '',
      tipoReferencia: p['Tipo_Referencia'] || '',
      horimetro:      parseInt(p['Horimetro']) || 0,
      dataTroca:      p['Data_Troca'] || '',
      na:             String(p['N/A'] || '').toUpperCase() === 'SIM',
      obs:            p['Observação'] || '',
      // v1.5:
      intervalo:      parseInt(p['Intervalo_H']) || 0,
      horasRodadas:   p['H_Rodadas']   || '',
      horasRestantes: p['H_Restantes'] || '',
      status:         p['Status']      || '',
      valorMostrado:  p['Valor_Mostrado'] || '',
      tipoContador:   p['Tipo_Contador']  || '',
    });
  });
  return { status: 'ok', visits, pecasLog };
}

// ═══════════════════════════════════════════════════════════════════════════
// GET MACHINES BY CLIENT
// ═══════════════════════════════════════════════════════════════════════════
function getMachinesByClient(clientQuery) {
  const q = String(clientQuery || '').trim();
  if (!q) return { status: 'error', error: 'client obrigatório' };
  const norm = v => String(v || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  const machines = getSheetDataActive('MAQUINAS')
    .filter(m => norm(m['Cliente']).includes(norm(q)))
    .map(m => {
      const eq = rowToMachineFromObj(m);
      eq.parts = getMachinePartsById(eq.id, eq.serial, eq.tag);
      return eq;
    });

  // v1.5: buscar dados do cliente para o PCF pré-preencher campos cadastrais
  const clientsData = getSheetDataActive('CLIENTES');
  const matchClient = clientsData.find(c => norm(String(c['Nome'] || '')) === norm(q));
  let clientData = null;
  if (matchClient) {
    clientData = {
      nome:      String(matchClient['Nome']        || '').trim(),
      cnpj:      String(matchClient['CNPJ']        || '').trim(),
      cidade:    String(matchClient['Cidade']      || '').trim(),
      telefone:  String(matchClient['Telefone']    || '').trim(),
      email:     String(matchClient['Email']       || '').trim(),
      contato:   String(matchClient['Contato']     || '').trim(),
      filial:    String(matchClient['Filial']      || '').trim(),
      nextVisit: String(matchClient['Prox_Visita'] || '').trim(),
      lastVisit: String(matchClient['Ult_Visita']  || '').trim(),
    };
  }

  const enrichedMachines = machines.map(m => enrichMachineWithClientData(m, clientData));
  return { status: 'ok', machines: enrichedMachines, clientData };
}

// ═══════════════════════════════════════════════════════════════════════════
// GET CLIENTS FOR FIELD
// ═══════════════════════════════════════════════════════════════════════════
function getClientsForField() {
  const clientsData = getSheetDataActive('CLIENTES');
  const fromMachines = getSheetDataActive('MAQUINAS')
    .map(m => String(m['Cliente'] || '').trim())
    .filter(Boolean);
  const fromClients = clientsData
    .map(c => String(c['Nome'] || '').trim())
    .filter(Boolean);
  const unique = [...new Set([...fromClients, ...fromMachines])]
    .sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));

  // v1.5: clientsFull com dados cadastrais para o PCF pré-preencher campos
  const clientsFull = clientsData.map(c => ({
    nome:      String(c['Nome']       || '').trim(),
    cnpj:      String(c['CNPJ']       || '').trim(),
    cidade:    String(c['Cidade']     || '').trim(),
    telefone:  String(c['Telefone']   || '').trim(),
    email:     String(c['Email']      || '').trim(),
    contato:   String(c['Contato']    || '').trim(),
    filial:    String(c['Filial']     || '').trim(),
    nextVisit: String(c['Prox_Visita']|| '').trim(),
    lastVisit: String(c['Ult_Visita'] || '').trim(),
  })).filter(c => c.nome);

  return { status: 'ok', clients: unique, clientsFull };
}


// ═══════════════════════════════════════════════════════════════════════════
// AUDITORIA — getSystemHealth
// ═══════════════════════════════════════════════════════════════════════════
function getSystemHealth() {
  const tabelas = ['MAQUINAS', 'MODELOS', 'CLIENTES', 'VISITAS', 'PECAS_LOG', 'MACHINE_PARTS'];
  const summary = {};

  tabelas.forEach(nome => {
    const rows = getSheetData(nome);
    const ativos   = rows.filter(r => String(r['Ativo'] || 'SIM').toUpperCase() !== 'NÃO');
    const inativos = rows.filter(r => String(r['Ativo'] || 'SIM').toUpperCase() === 'NÃO');
    const teste    = rows.filter(r => String(r['Tipo_Registro'] || '').toUpperCase() === 'TESTE');
    summary[nome] = {
      total:    rows.length,
      ativos:   ativos.length,
      inativos: inativos.length,
      teste:    teste.length
    };
  });

  const machineIds = new Set(
    getSheetData('MAQUINAS')
      .filter(r => String(r['Ativo'] || 'SIM').toUpperCase() !== 'NÃO')
      .map(r => String(r['ID'] || '').trim())
      .filter(Boolean)
  );
  const orphanParts = getSheetData('MACHINE_PARTS').filter(r => {
    const mid = String(r['Machine_ID'] || '').trim();
    return mid && !machineIds.has(mid);
  });
  const orphanVisits = getSheetData('VISITAS').filter(r => {
    const mid = String(r['Machine_ID'] || '').trim();
    return mid && !machineIds.has(mid);
  });

  return {
    status: 'ok',
    ts: new Date().toISOString(),
    tabelas: summary,
    orfaos: {
      machine_parts: orphanParts.length,
      visitas: orphanVisits.length
    }
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// AUDITORIA — getDuplicates
// ═══════════════════════════════════════════════════════════════════════════
function getDuplicates() {
  const machines = getSheetData('MAQUINAS');
  const dupsByMachineId    = {};
  const dupsByClientSerial = {};
  const dupsByClientTag    = {};

  machines.forEach(m => {
    const mid = String(m['ID'] || '').trim();
    const clientSerial = (String(m['Cliente'] || '').trim() + '|' + String(m['Série'] || '').trim()).toLowerCase();
    const clientTag    = (String(m['Cliente'] || '').trim() + '|' + String(m['TAG']    || '').trim()).toLowerCase();
    if (mid) dupsByMachineId[mid] = (dupsByMachineId[mid] || 0) + 1;
    if (clientSerial && clientSerial !== '|') dupsByClientSerial[clientSerial] = (dupsByClientSerial[clientSerial] || 0) + 1;
    if (clientTag    && clientTag    !== '|') dupsByClientTag[clientTag]        = (dupsByClientTag[clientTag]    || 0) + 1;
  });

  const filterDups = obj =>
    Object.entries(obj).filter(([, c]) => c > 1).map(([key, count]) => ({ key, count }));

  const porMachineId     = filterDups(dupsByMachineId);
  const porClienteSerial = filterDups(dupsByClientSerial);
  const porClienteTag    = filterDups(dupsByClientTag);

  return {
    status: 'ok',
    ts: new Date().toISOString(),
    duplicatas: {
      por_machine_id:    porMachineId,
      por_cliente_serie: porClienteSerial,
      por_cliente_tag:   porClienteTag
    },
    total_duplicatas: porMachineId.length + porClienteSerial.length + porClienteTag.length
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// AUDITORIA — getDeletedRecords
// ═══════════════════════════════════════════════════════════════════════════
function getDeletedRecords(sheetName) {
  const tabelasPermitidas = ['MAQUINAS', 'MODELOS', 'CLIENTES', 'VISITAS'];

  if (sheetName) {
    if (!tabelasPermitidas.includes(sheetName)) {
      return { status: 'error', error: 'Tabela inválida. Permitidas: ' + tabelasPermitidas.join(', ') };
    }
    const rows = getSheetData(sheetName).filter(r => String(r['Ativo'] || 'SIM').toUpperCase() === 'NÃO');
    return { status: 'ok', sheetName, records: rows, total: rows.length };
  }

  const result = {};
  tabelasPermitidas.forEach(nome => {
    const rows = getSheetData(nome).filter(r => String(r['Ativo'] || 'SIM').toUpperCase() === 'NÃO');
    result[nome] = { total: rows.length, records: rows };
  });
  return { status: 'ok', ts: new Date().toISOString(), deletados: result };
}

// ═══════════════════════════════════════════════════════════════════════════
// AUDITORIA — createBackupSnapshot
// ═══════════════════════════════════════════════════════════════════════════
function createBackupSnapshot() {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const tabelas = ['MAQUINAS', 'MODELOS', 'CLIENTES', 'VISITAS', 'PECAS_LOG', 'MACHINE_PARTS'];
  const created = [];

  tabelas.forEach(nome => {
    const source = SS.getSheetByName(nome);
    if (!source) {
      created.push({ sheet: 'BAK_' + nome + '_' + ts, status: 'aba não encontrada' });
      return;
    }
    const snapName = 'BAK_' + nome + '_' + ts;
    if (SS.getSheetByName(snapName)) {
      created.push({ sheet: snapName, status: 'já existia' });
      return;
    }
    const snap = source.copyTo(SS);
    snap.setName(snapName);
    SS.setActiveSheet(snap);
    SS.moveActiveSheet(SS.getNumSheets());
    created.push({ sheet: snapName, status: 'criado', rows: source.getLastRow() - 1 });
  });

  return {
    status: 'ok',
    ts: new Date().toISOString(),
    snapshot_prefix: 'BAK_*_' + ts,
    abas_criadas: created
  };
}
// ═══════════════════════════════════════════════════════════════════════════
// HELPERS DE MACHINE PARTS
// ═══════════════════════════════════════════════════════════════════════════
function getCurrentRefFromMachineParts(machineId, serial, tag, partId) {
  const sheet = getOrCreateSheet('MACHINE_PARTS', HEADERS.MACHINE_PARTS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return '';
  const headers = data[0];
  const idxMid = headers.indexOf('Machine_ID');
  const idxSer = headers.indexOf('Serial');
  const idxTag = headers.indexOf('TAG');
  const idxPid = headers.indexOf('Part_ID');
  const idxRef = headers.indexOf('Ref');

  for (let i = 1; i < data.length; i++) {
    const rowMid = String(data[i][idxMid] || '').trim();
    const rowSer = String(data[i][idxSer] || '').trim();
    const rowTag = String(data[i][idxTag] || '').trim();
    const rowPid = String(data[i][idxPid] || '').trim();
    const match = (machineId && rowMid === String(machineId).trim()) ||
                  (serial && rowSer === String(serial).trim()) ||
                  (tag && rowTag === String(tag).trim());
    if (match && rowPid === String(partId).trim()) {
      return String(data[i][idxRef] || '').trim();
    }
  }
  return '';
}

function updateMachinePartFromPreventiva(payload) {
  const machineId = String(payload.machine_id || '').trim();
  const serial = String(payload.serial || '').trim();
  const tag = String(payload.tag || '').trim();
  const partId = String(payload.partId || '').trim();
  if (!partId) return;

  const sheet = getOrCreateSheet('MACHINE_PARTS', HEADERS.MACHINE_PARTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxMid = headers.indexOf('Machine_ID');
  const idxSer = headers.indexOf('Serial');
  const idxTag = headers.indexOf('TAG');
  const idxPid = headers.indexOf('Part_ID');
  const idxRef = headers.indexOf('Ref');
  const idxRefAnterior = headers.indexOf('Ref_Anterior');
  const idxCreatedAt   = headers.indexOf('Created_At');

  for (let i = 1; i < data.length; i++) {
    const rowMid = String(data[i][idxMid] || '').trim();
    const rowSer = String(data[i][idxSer] || '').trim();
    const rowTag = String(data[i][idxTag] || '').trim();
    const rowPid = String(data[i][idxPid] || '').trim();
    const match = (machineId && rowMid === machineId) ||
                  (serial && rowSer === serial) ||
                  (tag && rowTag === tag);
    if (match && rowPid === partId) {
      const currentRef = String(data[i][idxRef] || '').trim();
      const refAnterior = String(payload.refAnterior || '').trim() ||
        ((currentRef && currentRef !== String(payload.refNova || '').trim()) ? currentRef : String(data[i][idxRefAnterior] || '').trim());
      const createdAt = String(data[i][idxCreatedAt] || '').trim();
      const updated = [
        machineId || rowMid,
        serial || rowSer,
        tag || rowTag,
        partId,
        payload.name || rowPid,
        parseInt(payload.hourTotal) || 0,
        parseInt(payload.interval) || 2000,
        String(payload.refNova || '').trim(),
        String(payload.na || '').toUpperCase() === 'SIM' ? 'SIM' : 'NÃO',
        refAnterior,
        createdAt || (payload.now || new Date().toISOString()),
        'SIM', 'PRODUCAO', '', '', payload.now || new Date().toISOString(),
        '',  // Valor_Mostrado: zerado após troca
        ''   // Contador: zerado após troca
      ];
      sheet.getRange(i + 1, 1, 1, HEADERS.MACHINE_PARTS.length).setValues([updated]);
      return;
    }
  }

  sheet.appendRow([
    machineId, serial, tag,
    partId,
    payload.name || partId,
    parseInt(payload.hourTotal) || 0,
    parseInt(payload.interval) || 2000,
    String(payload.refNova || '').trim(),
    String(payload.na || '').toUpperCase() === 'SIM' ? 'SIM' : 'NÃO',
    String(payload.refAnterior || '').trim(),
    payload.now || new Date().toISOString(),
    'SIM', 'PRODUCAO', '', '', payload.now || new Date().toISOString(),
    '',  // Valor_Mostrado
    ''   // Contador
  ]);
}

// ═══════════════════════════════════════════════════════════════════════════
// UPDATE VALOR_MOSTRADO E CONTADOR (peças conferidas/na — sem alterar lastChange)
// ═══════════════════════════════════════════════════════════════════════════
function updateMachinePartValorMostrado(machineId, serial, tag, partId, valorMostrado, contador, now) {
  try {
    const sheet = getOrCreateSheet('MACHINE_PARTS', HEADERS.MACHINE_PARTS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxMid = headers.indexOf('Machine_ID');
    const idxSer = headers.indexOf('Serial');
    const idxTag = headers.indexOf('TAG');
    const idxPid = headers.indexOf('Part_ID');
    const idxVM  = headers.indexOf('Valor_Mostrado');
    const idxCnt = headers.indexOf('Contador');
    const idxAtual = headers.indexOf('Atualizado');
    if (idxVM < 0 || idxCnt < 0) return; // colunas ainda não existem
    for (let i = 1; i < data.length; i++) {
      const rowMid = String(data[i][idxMid] || '').trim();
      const rowSer = String(data[i][idxSer] || '').trim();
      const rowTag = String(data[i][idxTag] || '').trim();
      const rowPid = String(data[i][idxPid] || '').trim();
      const match = (machineId && rowMid === String(machineId).trim()) ||
                    (serial    && rowSer === String(serial).trim())    ||
                    (tag       && rowTag === String(tag).trim());
      if (match && rowPid === String(partId).trim()) {
        sheet.getRange(i + 1, idxVM  + 1).setValue(valorMostrado || '');
        sheet.getRange(i + 1, idxCnt + 1).setValue(contador      || '');
        if (idxAtual >= 0) sheet.getRange(i + 1, idxAtual + 1).setValue(now || new Date().toISOString());
        return;
      }
    }
  } catch(e) { /* silencioso */ }
}

// ═══════════════════════════════════════════════════════════════════════════
// SAVE / UPDATE MACHINE
// ═══════════════════════════════════════════════════════════════════════════
function saveMachine(m) {
  if (!m) return { status: 'error', error: 'Dados ausentes' };
  const sheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const now = new Date().toISOString();
  const ad = auditDefaults(m.tipoRegistro);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId    = headers.indexOf('ID');
  const idxSerie = headers.indexOf('Série');
  const idxAtivo = headers.indexOf('Ativo');

  for (let i = 1; i < data.length; i++) {
    const idxCli = headers.indexOf('Cliente');
    const idxTag = headers.indexOf('TAG');
    const norm = v => String(v||'').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
    const sameId     = String(data[i][idxId]    || '').trim() === String(m.id    || '').trim();
    const sameClientSerial = m.serial && m.client && norm(data[i][idxCli])===norm(m.client) && norm(data[i][idxSerie])===norm(m.serial);
    const sameClientTag = !m.serial && m.tag && m.client && norm(data[i][idxCli])===norm(m.client) && norm(data[i][idxTag])===norm(m.tag);
    if (sameId || sameClientSerial || sameClientTag) {
      const idxCreated = headers.indexOf('Created_At');
      const idxTipo    = headers.indexOf('Tipo_Registro');
      const originalCreated = String(data[i][idxCreated] || now);
      const originalTipo    = String(data[i][idxTipo]    || 'PRODUCAO');
      const originalAtivo   = String(data[i][idxAtivo]   || 'SIM');
      sheet.getRange(i+1, 1, 1, HEADERS.MAQUINAS.length).setValues([[
        m.id || data[i][idxId],
        m.client || '', m.branch || '',
        m.brand  || '', m.model  || '',
        m.serial || '', m.year   || '',
        m.tag    || '', m.location || '',
        parseInt(m.hourTotal) || 0,
        parseInt(m.hpw)       || 0,
        m.obs || '',
        originalAtivo, originalTipo, originalCreated, '', '', now
      ]]);
      return { status: 'ok', action: 'updated' };
    }
  }

  sheet.appendRow([
    m.id || (m.serial ? machineKey(m.client||'', m.brand||'', m.model||'', m.serial||'') : '') || (m.tag ? machineKey(m.client||'', m.brand||'', m.model||'', 'TAG-'+m.tag) : '') || ('EQ-' + new Date().getTime() + '-' + Math.random().toString(36).slice(2,6)),
    m.client || '', m.branch || '',
    m.brand  || '', m.model  || '',
    m.serial || '', m.year   || '',
    m.tag    || '', m.location || '',
    parseInt(m.hourTotal) || 0,
    parseInt(m.hpw)       || 0,
    m.obs || '',
    ad.ativo, ad.tipoRegistro, ad.createdAt, ad.deletedAt, ad.deletedBy, now
  ]);
  return { status: 'ok', action: 'inserted' };
}

// ═══════════════════════════════════════════════════════════════════════════
// SAVE MODEL
// ═══════════════════════════════════════════════════════════════════════════
function saveModel(m) {
  if (!m) return { status: 'error', error: 'Dados ausentes' };
  const sheet = getOrCreateSheet('MODELOS', HEADERS.MODELOS);
  const now = new Date().toISOString();
  const ad = auditDefaults(m.tipoRegistro);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('ID');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId] || '').trim() === String(m.id || '').trim()) {
      const idxCreated = headers.indexOf('Created_At');
      const idxTipo    = headers.indexOf('Tipo_Registro');
      const idxAtivo   = headers.indexOf('Ativo');
      sheet.getRange(i+1, 1, 1, HEADERS.MODELOS.length).setValues([[
        m.id, m.brand || '', m.model || '', m.type || '',
        m.power || '', m.pressure || '', m.obs || '',
        String(data[i][idxAtivo]   || 'SIM'),
        String(data[i][idxTipo]    || 'PRODUCAO'),
        String(data[i][idxCreated] || now),
        '', '', now
      ]]);
      return { status: 'ok', action: 'updated' };
    }
  }

  sheet.appendRow([
    m.id || 'MOD-' + new Date().getTime(),
    m.brand || '', m.model || '', m.type || '',
    m.power || '', m.pressure || '', m.obs || '',
    ad.ativo, ad.tipoRegistro, ad.createdAt, ad.deletedAt, ad.deletedBy, now
  ]);
  return { status: 'ok', action: 'inserted' };
}

// ═══════════════════════════════════════════════════════════════════════════
// SAVE PART MASTER
// ═══════════════════════════════════════════════════════════════════════════
function savePartMaster(p) {
  if (!p || !p.id) return { status: 'error', error: 'id da peça obrigatório' };

  ensureSheetHeaders('PARTS_MASTER', HEADERS.PARTS_MASTER);
  const sheet = getOrCreateSheet('PARTS_MASTER', HEADERS.PARTS_MASTER);
  const now = new Date().toISOString();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxPartId  = headers.indexOf('Part_ID');
  const idxModelId = headers.indexOf('Model_ID');

  const scopeRaw = String(p.scope || p.partScope || 'direct').trim().toLowerCase();
  const partScope = scopeRaw === 'sub' ? 'sub' : 'direct';
  const row = [
    p.id, p.modelId || '', p.name || '', p.ref || '',
    p.partBrand || '', p.supplierPrimary || '',
    p.slot || '', parseInt(p.qty) || 1,
    parseInt(p.interval) || 0,
    p.criticality || 'normal',
    parseFloat(p.cost) || 0,
    p.obs || '',
    'SIM', 'PRODUCAO', now, '', '', now,
    partScope,
    partScope === 'sub' ? (p.subId || '') : '',
    partScope === 'sub' ? (p.subName || '') : '',
    partScope === 'sub' ? (p.subCategory || '') : '',
    partScope === 'sub' ? (p.subDesc || '') : '',
    partScope === 'sub' ? (parseInt(p.subInterval, 10) || 0) : ''
  ];

  for (let i = 1; i < data.length; i++) {
    const samePartId  = data[i][idxPartId]  === p.id;
    const sameModelId = data[i][idxModelId] === (p.modelId || '');
    if (samePartId && sameModelId) {
      sheet.getRange(i + 1, 1, 1, HEADERS.PARTS_MASTER.length).setValues([row]);
      return { status: 'ok', action: 'updated' };
    }
  }

  sheet.appendRow(row);
  return { status: 'ok', action: 'inserted' };
}

// ═══════════════════════════════════════════════════════════════════════════
// REPLACE PART SIMILARITIES
// ═══════════════════════════════════════════════════════════════════════════
function replacePartSimilarities(partId, modelId, similarities) {
  if (!partId)  return { status: 'error', error: 'partId obrigatório' };
  if (!modelId) return { status: 'error', error: 'modelId obrigatório' };

  const sheet = getOrCreateSheet('PART_SIMILARITIES', HEADERS.PART_SIMILARITIES);
  const now = new Date().toISOString();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxPid = headers.indexOf('Part_ID');
  const idxMid = headers.indexOf('Model_ID');

  const rowsToKeep = [headers];
  for (let i = 1; i < data.length; i++) {
    const samePart  = String(data[i][idxPid]).trim() === String(partId).trim();
    const sameModel = String(data[i][idxMid]).trim() === String(modelId).trim();
    if (!(samePart && sameModel)) rowsToKeep.push(data[i]);
  }

  sheet.clearContents();
  if (rowsToKeep.length > 0) {
    sheet.getRange(1, 1, rowsToKeep.length, HEADERS.PART_SIMILARITIES.length).setValues(rowsToKeep);
  }

  const newRows = (similarities || []).map((s, idx) => [
    'SIM-' + partId + '-' + modelId + '-' + idx,
    partId, modelId,
    typeof s === 'string' ? s : (s.ref   || ''),
    typeof s === 'object'  ? (s.brand || '') : '',
    typeof s === 'object'  ? (s.obs   || '') : '',
    'SIM', 'PRODUCAO', now, '', '', now
  ]);

  if (newRows.length > 0) {
    sheet.getRange(rowsToKeep.length + 1, 1, newRows.length, HEADERS.PART_SIMILARITIES.length).setValues(newRows);
  }

  return { status: 'ok', replaced: newRows.length };
}

// ═══════════════════════════════════════════════════════════════════════════
// GET CATALOG FULL
// ═══════════════════════════════════════════════════════════════════════════
function getCatalogFull() {
  ensureSheetHeaders('PARTS_MASTER', HEADERS.PARTS_MASTER);
  return {
    status: 'ok',
    models:       getSheetDataActive('MODELOS'),
    parts:        getSheetDataActive('PARTS_MASTER'),
    similarities: getSheetData('PART_SIMILARITIES')
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// SAVE / UPDATE CLIENT
// ═══════════════════════════════════════════════════════════════════════════
function saveClient(c) {
  if (!c) return { status: 'error', error: 'Dados ausentes' };
  const sheet = getOrCreateSheet('CLIENTES', HEADERS.CLIENTES);
  const now = new Date().toISOString();
  const ad = auditDefaults(c.tipoRegistro);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('ID');

  // v1.5: helper para não sobrescrever campo existente com vazio
  const mergeVal = (newVal, existVal) => (newVal && String(newVal).trim()) ? String(newVal).trim() : String(existVal || '');

  const buildRow = (ativo, tipo, created, existing) => {
    const ex = existing || {};
    return [
      c.id || ex['ID'] || 'CLI-' + new Date().getTime(),
      mergeVal(c.nome,         ex['Nome']),
      mergeVal(c.cnpj,         ex['CNPJ']),
      mergeVal(c.cidade,       ex['Cidade']),
      mergeVal(c.telefone,     ex['Telefone']),
      mergeVal(c.email,        ex['Email']),
      mergeVal(c.observacoes || c.obs, ex['Observações']),
      parseInt(c.antecedencia_alerta_dias) || parseInt(ex['Antecedencia_Alerta_Dias']) || 0,
      mergeVal(c.drive_url,    ex['Drive_URL']),
      ativo, tipo, created, '', '', now,
      // v1.5: novas colunas
      mergeVal(c.contato,      ex['Contato']),
      mergeVal(c.filial,       ex['Filial']),
      mergeVal(c.nextVisit || c.prox_visita, ex['Prox_Visita']),
      mergeVal(c.lastVisit  || c.ult_visita, ex['Ult_Visita']),
    ];
  };

  const _normEnt = v => String(v || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\b(ltda|s\.?a\.?|eireli|me|epp|ss|lda|inc|llc|do brasil|do norte|do sul|e cia)\.?\b/gi, '')
    .replace(/[^a-z0-9\s]/g, '').replace(/\s+/g, ' ').trim();

  const idxNome = headers.indexOf('Nome');
  const normNovo = _normEnt(c.nome || '');
  let _foundRow = -1;
  let _matchType = '';

  for (let i = 1; i < data.length; i++) {
    const rowId   = String(data[i][idxId]   || '').trim();
    const rowNome = idxNome >= 0 ? String(data[i][idxNome] || '').trim() : '';
    if (rowId && rowId === String(c.id || '').trim()) {
      _foundRow = i; _matchType = 'id'; break;
    }
    if (normNovo && _normEnt(rowNome) === normNovo && _foundRow < 0) {
      _foundRow = i; _matchType = 'nome';
    }
  }

  if (_foundRow >= 0) {
    const idxCreated = headers.indexOf('Created_At');
    const idxTipo    = headers.indexOf('Tipo_Registro');
    const idxAtivo   = headers.indexOf('Ativo');
    if (_matchType === 'nome') {
      c.id = String(data[_foundRow][idxId] || c.id).trim();
      Logger.log('saveClient GAS: match por nome — usando ID da planilha: ' + c.id);
    }
    // v1.5: montar objeto da linha existente para merge seguro
    const existingObj = {};
    headers.forEach((h, i) => { existingObj[h] = data[_foundRow][i]; });
    const rowData = buildRow(
      String(data[_foundRow][idxAtivo]   || 'SIM'),
      String(data[_foundRow][idxTipo]    || 'PRODUCAO'),
      String(data[_foundRow][idxCreated] || now),
      existingObj
    );
    // Garantir que o row tenha comprimento suficiente para colunas novas
    while (rowData.length < HEADERS.CLIENTES.length) rowData.push('');
    sheet.getRange(_foundRow + 1, 1, 1, HEADERS.CLIENTES.length).setValues([rowData]);
    return { status: 'ok', action: 'updated', matchedBy: _matchType };
  }

  const newRow = buildRow(ad.ativo, ad.tipoRegistro, ad.createdAt, {});
  while (newRow.length < HEADERS.CLIENTES.length) newRow.push('');
  sheet.appendRow(newRow);
  return { status: 'ok', action: 'inserted' };
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════
function ensureMachineFromVisit(body, machineId, visitDate, now) {
  const sheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxId  = headers.indexOf('ID');
  const idxCli = headers.indexOf('Cliente');
  const idxFil = headers.indexOf('Filial');
  const idxMar = headers.indexOf('Marca');
  const idxMod = headers.indexOf('Modelo');
  const idxSer = headers.indexOf('Série');
  const idxAno = headers.indexOf('Ano');
  const idxTag = headers.indexOf('TAG');
  const idxLoc = headers.indexOf('Localização');
  const idxHor = headers.indexOf('Hor.Total');
  const idxHpw = headers.indexOf('h/Semana');
  const idxObs = headers.indexOf('Observações');
  const idxUpd = headers.indexOf('Atualizado');

  const client    = body.client    || '';
  const branch    = body.branch    || '';
  const brand     = body.brand     || '';
  const model     = body.model     || '';
  const serial    = body.serial    || '';
  const year      = body.year      || '';
  const tag       = body.tag       || '';
  const location  = body.location  || '';
  const hourTotal = parseInt(body.hourTotal) || 0;
  const hpw       = parseInt(body.hpw)       || 0;
  const obs       = body.generalObs || body.obs || '';

  for (let i = 1; i < data.length; i++) {
    const rowId  = String(data[i][idxId]  || '').trim();
    const rowSer = String(data[i][idxSer] || '').trim();
    const rowTag = String(data[i][idxTag] || '').trim();

    const norm = v => String(v||'').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
    const rowCli = String(data[i][headers.indexOf('Cliente')] || '').trim();
    const match = (machineId && rowId  === String(machineId).trim()) ||
                  (serial && client && norm(rowSer)===norm(serial) && norm(rowCli)===norm(client)) ||
                  (!serial && tag && client && norm(rowTag)===norm(tag) && norm(rowCli)===norm(client));

    if (match) {
      const existingHour = parseInt(data[i][idxHor]) || 0;
      sheet.getRange(i + 1, idxId  + 1).setValue(machineId || rowId);
      if (client)   sheet.getRange(i + 1, idxCli + 1).setValue(client);
      if (branch)   sheet.getRange(i + 1, idxFil + 1).setValue(branch);
      if (brand)    sheet.getRange(i + 1, idxMar + 1).setValue(brand);
      if (model)    sheet.getRange(i + 1, idxMod + 1).setValue(model);
      if (serial)   sheet.getRange(i + 1, idxSer + 1).setValue(serial);
      if (year)     sheet.getRange(i + 1, idxAno + 1).setValue(year);
      if (tag)      sheet.getRange(i + 1, idxTag + 1).setValue(tag);
      if (location) sheet.getRange(i + 1, idxLoc + 1).setValue(location);
      if (hourTotal > 0) sheet.getRange(i + 1, idxHor + 1).setValue(hourTotal);
      if (hpw)      sheet.getRange(i + 1, idxHpw + 1).setValue(hpw);
      if (obs)      sheet.getRange(i + 1, idxObs + 1).setValue(obs);
      sheet.getRange(i + 1, idxUpd + 1).setValue(now);
      return { status: 'ok', action: 'updated', machineId: machineId || rowId };
    }
  }

  // Fallback: buscar por serial normalizado + cliente normalizado
  if (serial) {
    const _normStr = v => String(v || '').trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
    const _normSerBk = _normStr(serial);
    const _normCliBk = _normStr(client);
    for (let j = 1; j < data.length; j++) {
      if (_normStr(String(data[j][idxSer] || '')) === _normSerBk &&
          _normStr(String(data[j][idxCli] || '')) === _normCliBk) {
        Logger.log('ensureMachineFromVisit: match serial+cliente normalizado, linha ' + j);
        const existingHour = parseInt(data[j][idxHor]) || 0;
        const rowId = String(data[j][idxId] || '').trim();
        sheet.getRange(j + 1, idxId  + 1).setValue(machineId || rowId);
        if (client)   sheet.getRange(j + 1, idxCli + 1).setValue(client);
        if (branch)   sheet.getRange(j + 1, idxFil + 1).setValue(branch);
        if (brand)    sheet.getRange(j + 1, idxMar + 1).setValue(brand);
        if (model)    sheet.getRange(j + 1, idxMod + 1).setValue(model);
        if (serial)   sheet.getRange(j + 1, idxSer + 1).setValue(serial);
        if (year)     sheet.getRange(j + 1, idxAno + 1).setValue(year);
        if (tag)      sheet.getRange(j + 1, idxTag + 1).setValue(tag);
        if (location) sheet.getRange(j + 1, idxLoc + 1).setValue(location);
        if (hourTotal > 0) sheet.getRange(j + 1, idxHor + 1).setValue(hourTotal);
        if (hpw)      sheet.getRange(j + 1, idxHpw + 1).setValue(hpw);
        if (obs)      sheet.getRange(j + 1, idxObs + 1).setValue(obs);
        sheet.getRange(j + 1, idxUpd + 1).setValue(now);
        return { status: 'ok', action: 'updated', machineId: machineId || rowId };
      }
    }
  }

  sheet.appendRow([
    machineId || machineKey(client, brand, model, serial) || ('EQ-' + new Date().getTime()),
    client, branch, brand, model, serial, year, tag, location,
    hourTotal, hpw, obs,
    'SIM', 'PRODUCAO', now, '', '', now
  ]);
  return { status: 'ok', action: 'inserted', machineId: machineId };
}

function updateMachineHorímetro(machineId, client, serial, tag, hourTotal, visitDate) {
  const sheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId    = headers.indexOf('ID');
  const idxSerie = headers.indexOf('Série');
  const idxTag   = headers.indexOf('TAG');
  const idxHor   = headers.indexOf('Hor.Total');
  const idxUpd   = headers.indexOf('Atualizado');

  for (let i = 1; i < data.length; i++) {
    const rowId  = String(data[i][idxId]    || '').trim();
    const rowSer = String(data[i][idxSerie] || '').trim();
    const rowTag = String(data[i][idxTag]   || '').trim();

    const norm = v => String(v||'').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
    const rowCli = String(data[i][headers.indexOf('Cliente')] || '').trim();
    const match = (machineId && rowId  === String(machineId).trim()) ||
                  (serial && client && norm(rowSer)===norm(serial) && norm(rowCli)===norm(client)) ||
                  (!serial && tag && client && norm(rowTag)===norm(tag) && norm(rowCli)===norm(client));

    if (match) {
      if (hourTotal > 0) {
        sheet.getRange(i+1, idxHor+1).setValue(hourTotal);
      }
      sheet.getRange(i+1, idxUpd+1).setValue(new Date().toISOString());
      return;
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SOFT DELETE (substitui deleteRow para MAQUINAS, MODELOS, CLIENTES)
// ═══════════════════════════════════════════════════════════════════════════
function softDelete(sheetName, id, deletedBy) {
  const sheet = getOrCreateSheet(sheetName, HEADERS[sheetName]);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId  = headers.indexOf('ID');
  const idxAlt = headers.indexOf('ID_Visita');
  const idxFinal = idxId >= 0 ? idxId : idxAlt;

  const idxAtivo      = headers.indexOf('Ativo');
  const idxDeletedAt  = headers.indexOf('Deleted_At');
  const idxDeletedBy  = headers.indexOf('Deleted_By');
  const idxAtualizado = headers.indexOf('Atualizado');
  const now = new Date().toISOString();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxFinal] || '').trim() === String(id).trim()) {
      if (idxAtivo      >= 0) sheet.getRange(i+1, idxAtivo+1).setValue('NÃO');
      if (idxDeletedAt  >= 0) sheet.getRange(i+1, idxDeletedAt+1).setValue(now);
      if (idxDeletedBy  >= 0) sheet.getRange(i+1, idxDeletedBy+1).setValue(deletedBy || 'admin');
      if (idxAtualizado >= 0) sheet.getRange(i+1, idxAtualizado+1).setValue(now);
      return { status: 'ok', action: 'soft_deleted' };
    }
  }
  return { status: 'ok', note: 'Não encontrado' };
}

// ═══════════════════════════════════════════════════════════════════════════
// HARD DELETE — apenas registros com Tipo_Registro = TESTE
// ═══════════════════════════════════════════════════════════════════════════
function hardDeleteTestRecord(sheetName, id) {
  if (!sheetName || !id) return { status: 'error', error: 'sheetName e id obrigatórios' };
  if (!HEADERS[sheetName]) return { status: 'error', error: 'Tabela inválida: ' + sheetName };

  const sheet = getOrCreateSheet(sheetName, HEADERS[sheetName]);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId   = headers.indexOf('ID');
  const idxTipo = headers.indexOf('Tipo_Registro');

  if (idxTipo < 0) return { status: 'error', error: 'Tabela sem coluna Tipo_Registro' };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId] || '').trim() === String(id).trim()) {
      const tipo = String(data[i][idxTipo] || '').trim().toUpperCase();
      if (tipo !== 'TESTE') {
        return { status: 'error', error: 'Hard delete permitido apenas para TESTE. Este é: ' + (tipo || 'PRODUCAO') };
      }
      sheet.deleteRow(i + 1);
      return { status: 'ok', action: 'hard_deleted' };
    }
  }
  return { status: 'ok', note: 'Não encontrado' };
}

// ═══════════════════════════════════════════════════════════════════════════
// GET SHEET DATA
// ═══════════════════════════════════════════════════════════════════════════
function getSheetData(sheetName) {
  const sheet = getOrCreateSheet(sheetName, HEADERS[sheetName]);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, j) => obj[h] = row[j]);
    return obj;
  });
}

// Retorna apenas linhas onde Ativo != 'NÃO'
function getSheetDataActive(sheetName) {
  const all = getSheetData(sheetName);
  return all.filter(row => String(row['Ativo'] || 'SIM').trim().toUpperCase() !== 'NÃO');
}

function rowToMachine(row) {
  return {
    id:        row['ID']           || '',
    client:    row['Cliente']      || '',
    branch:    row['Filial']       || '',
    brand:     row['Marca']        || '',
    model:     row['Modelo']       || '',
    serial:    String(row['Série'] || ''),
    year:      row['Ano']          || '',
    tag:       row['TAG']          || '',
    location:  row['Localização']  || '',
    hourTotal: parseInt(row['Hor.Total']) || 0,
    hpw:       parseInt(row['h/Semana'])  || 0,
    obs:       row['Observações']  || ''
  };
}

function rowToMachineFromObj(row) {
  return rowToMachine(row);
}

function ensureSheetHeaders(sheetName, expectedHeaders) {
  const sheet = getOrCreateSheet(sheetName, expectedHeaders);
  if (sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    sheet.getRange(1, 1, 1, expectedHeaders.length)
      .setFontWeight('bold')
      .setBackground('#1a3a6b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return;
  }

  const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  expectedHeaders.forEach(col => {
    if (!current.includes(col)) {
      const newColIdx = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColIdx)
        .setValue(col)
        .setFontWeight('bold')
        .setBackground('#1a3a6b')
        .setFontColor('#ffffff');
    }
  });
}

function getOrCreateSheet(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a3a6b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 150);
  }
  return sheet;
}

// ═══════════════════════════════════════════════════════════════════════════
// MIGRAÇÃO — Atualizar IDs existentes para machine_key (rodar uma vez)
// Execute manualmente no Apps Script Editor: migrateExistingIds()
// ═══════════════════════════════════════════════════════════════════════════
function migrateExistingIds() {
  const sheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return 'Sem dados';
  const headers   = data[0];
  const idxId     = headers.indexOf('ID');
  const idxClient = headers.indexOf('Cliente');
  const idxBrand  = headers.indexOf('Marca');
  const idxModel  = headers.indexOf('Modelo');
  const idxSerial = headers.indexOf('Série');
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const currentId = String(data[i][idxId] || '').trim();
    if (currentId.startsWith('MK-')) continue;
    const mk = machineKey(
      data[i][idxClient] || '',
      data[i][idxBrand]  || '',
      data[i][idxModel]  || '',
      data[i][idxSerial] || ''
    );
    sheet.getRange(i + 1, idxId + 1).setValue(mk);
    updated++;
  }
  const vSheet = SS.getSheetByName('VISITAS');
  if (vSheet) {
    const vData = vSheet.getDataRange().getValues();
    const vH    = vData[0];
    const viMid = vH.indexOf('Machine_ID');
    const viCli = vH.indexOf('Cliente');
    const viMar = vH.indexOf('Marca');
    const viMod = vH.indexOf('Modelo');
    const viSer = vH.indexOf('Série');
    if (viMid >= 0) {
      for (let i = 1; i < vData.length; i++) {
        const mid = String(vData[i][viMid] || '').trim();
        if (mid) continue;
        const mk = machineKey(
          vData[i][viCli] || '',
          vData[i][viMar] || '',
          vData[i][viMod] || '',
          vData[i][viSer] || ''
        );
        vSheet.getRange(i + 1, viMid + 1).setValue(mk);
      }
    }
  }
  return 'Migração concluída: ' + updated + ' máquina(s) atualizada(s)';
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
// ═══════════════════════════════════════════════════════════════════════════
// initializeDatabase() — Proper Care GAS v6
// ───────────────────────────────────────────────────────────────────────────
// QUANDO EXECUTAR:
//   Execute manualmente no Apps Script Editor após cada nova implantação
//   da GAS. Também pode ser chamado repetidamente — é idempotente.
//
// O QUE FAZ:
//   1. Cria todas as abas que faltam com seus cabeçalhos corretos
//   2. Adiciona colunas novas em abas existentes (sem apagar dados)
//   3. Detecta PARTS_MASTER com cabeçalho corrompido e aborta com erro claro
//   4. Valida integridade básica (orphans, duplicatas de modelo)
//   5. Gera relatório completo no Logger do Apps Script
//
// SEGURO: nunca apaga dados, nunca move linhas, nunca altera registros.
// ═══════════════════════════════════════════════════════════════════════════

function initializeDatabase() {
  const log = [];
  const warnings = [];
  const errors = [];
  const now = new Date().toISOString();

  log.push('╔══════════════════════════════════════════════╗');
  log.push('║  Proper Care — initializeDatabase()          ║');
  log.push('║  ' + now + '  ║');
  log.push('╚══════════════════════════════════════════════╝');

  // ── 1. Garantir existência e cabeçalhos de todas as abas ──────────────
  const SHEET_DEFS = [
    { name: 'MAQUINAS',         headers: HEADERS.MAQUINAS         },
    { name: 'MODELOS',          headers: HEADERS.MODELOS           },
    { name: 'CLIENTES',         headers: HEADERS.CLIENTES          },
    { name: 'VISITAS',          headers: HEADERS.VISITAS           },
    { name: 'PECAS_LOG',        headers: HEADERS.PECAS_LOG         },
    { name: 'MACHINE_PARTS',    headers: HEADERS.MACHINE_PARTS     },
    { name: 'PARTS_MASTER',     headers: HEADERS.PARTS_MASTER      },
    { name: 'PART_SIMILARITIES',headers: HEADERS.PART_SIMILARITIES },
  ];

  SHEET_DEFS.forEach(def => {
    const result = _ensureSheetComplete(def.name, def.headers);
    log.push(result.message);
    if (result.warning) warnings.push(result.warning);
    if (result.error)   errors.push(result.error);
  });

  // ── 2. Validação crítica: PARTS_MASTER com cabeçalho corrompido ───────
  const pmCheck = _checkPartsMasterHeader();
  if (pmCheck.corrupted) {
    const msg = '🔴 PARTS_MASTER: cabeçalho CORROMPIDO — linha 1 contém dado, não cabeçalho. ' +
                'Corrija manualmente antes de continuar. Dado encontrado na L1: ' + pmCheck.firstCellValue;
    errors.push(msg);
    log.push(msg);
  } else {
    log.push('✅ PARTS_MASTER: cabeçalho na linha 1 validado.');
  }

  // ── 3. Relatório de orphans ────────────────────────────────────────────
  const orphanResult = _checkOrphans();
  log.push('');
  log.push('── Verificação de orphans ──');
  log.push('  MACHINE_PARTS órfãos: ' + orphanResult.machinePartsOrphans);
  log.push('  VISITAS órfãs:        ' + orphanResult.visitasOrphans);
  if (orphanResult.machinePartsOrphans > 0 || orphanResult.visitasOrphans > 0) {
    warnings.push('Existem registros órfãos. Execute getDuplicates() e getSystemHealth() para detalhes.');
  }

  // ── 4. Relatório de duplicatas em MODELOS ─────────────────────────────
  const dupModelos = _checkModeloDuplicates();
  log.push('');
  log.push('── Verificação de modelos duplicados ──');
  if (dupModelos.length > 0) {
    dupModelos.forEach(d => {
      const msg = '  ⚠️  Modelo duplicado (ATIVO): Marca=' + d.brand + ' Modelo=' + d.model +
                  ' | IDs: ' + d.ids.join(', ');
      log.push(msg);
      warnings.push(msg);
    });
  } else {
    log.push('  Nenhum modelo duplicado ativo encontrado.');
  }

  // ── 5. Resumo final ───────────────────────────────────────────────────
  log.push('');
  log.push('── Resumo ──');
  log.push('  Erros críticos: '   + errors.length);
  log.push('  Avisos:         '   + warnings.length);
  log.push(errors.length === 0 ? '✅ Banco de dados pronto para uso.' : '🔴 Corrija os erros antes de usar.');

  const fullLog = log.join('\n');
  Logger.log(fullLog);

  return {
    status:   errors.length === 0 ? 'ok' : 'error',
    errors,
    warnings,
    log:      fullLog,
    ts:       now
  };
}

// ── Garante que a aba existe e tem todas as colunas do schema ─────────────
function _ensureSheetComplete(sheetName, expectedHeaders) {
  let sheet = SS.getSheetByName(sheetName);
  const result = { message: '', warning: null, error: null };

  // Aba não existe: criar com cabeçalho completo
  if (!sheet) {
    sheet = SS.insertSheet(sheetName);
    sheet.appendRow(expectedHeaders);
    sheet.getRange(1, 1, 1, expectedHeaders.length)
      .setFontWeight('bold')
      .setBackground('#1a3a6b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, expectedHeaders.length, 150);
    result.message = '✅ ' + sheetName + ': aba criada com ' + expectedHeaders.length + ' colunas.';
    return result;
  }

  // Aba existe: verificar e adicionar colunas faltantes
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    // Aba totalmente vazia
    sheet.appendRow(expectedHeaders);
    sheet.getRange(1, 1, 1, expectedHeaders.length)
      .setFontWeight('bold')
      .setBackground('#1a3a6b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    result.message = '✅ ' + sheetName + ': cabeçalho inicial inserido (' + expectedHeaders.length + ' colunas).';
    return result;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const missing = expectedHeaders.filter(h => !currentHeaders.includes(h));

  if (missing.length === 0) {
    result.message = '✅ ' + sheetName + ': cabeçalho completo (' + currentHeaders.length + ' colunas).';
  } else {
    missing.forEach(col => {
      const newColIdx = sheet.getLastColumn() + 1;
      const cell = sheet.getRange(1, newColIdx);
      cell.setValue(col)
          .setFontWeight('bold')
          .setBackground('#1a3a6b')
          .setFontColor('#ffffff');
    });
    result.message = '🔧 ' + sheetName + ': ' + missing.length + ' coluna(s) adicionada(s): [' + missing.join(', ') + ']';
  }

  return result;
}

// ── Detecta PARTS_MASTER com cabeçalho corrompido ────────────────────────
function _checkPartsMasterHeader() {
  const sheet = SS.getSheetByName('PARTS_MASTER');
  if (!sheet || sheet.getLastRow() === 0) {
    return { corrupted: false };
  }
  const firstCell = sheet.getRange(1, 1).getValue();
  const firstCellStr = String(firstCell || '').trim();

  // Se a célula A1 começa com 'dp-', 'imp', 'EQ-' ou 'MK-' é dado, não cabeçalho
  const looksLikeData = /^(dp-|imp|EQ-|MK-|VIS-|CLI-|MOD-|SIM-)/.test(firstCellStr)
    || (firstCellStr.length > 0 && firstCellStr !== 'Part_ID');

  return {
    corrupted:      looksLikeData,
    firstCellValue: firstCellStr
  };
}

// ── Contagem rápida de orphans ────────────────────────────────────────────
function _checkOrphans() {
  try {
    const maqSheet = SS.getSheetByName('MAQUINAS');
    if (!maqSheet) return { machinePartsOrphans: 0, visitasOrphans: 0 };

    const maqData    = maqSheet.getDataRange().getValues();
    const maqHeaders = maqData[0];
    const idxId      = maqHeaders.indexOf('ID');
    const idxAtivo   = maqHeaders.indexOf('Ativo');

    const activeMachineIds = new Set();
    for (let i = 1; i < maqData.length; i++) {
      const ativo = String(maqData[i][idxAtivo] || 'SIM').toUpperCase();
      if (ativo !== 'NÃO') {
        const id = String(maqData[i][idxId] || '').trim();
        if (id) activeMachineIds.add(id);
      }
    }

    let mpOrphans = 0;
    const mpSheet = SS.getSheetByName('MACHINE_PARTS');
    if (mpSheet && mpSheet.getLastRow() > 1) {
      const mpData    = mpSheet.getDataRange().getValues();
      const mpHeaders = mpData[0];
      const mpIdxMid  = mpHeaders.indexOf('Machine_ID');
      for (let i = 1; i < mpData.length; i++) {
        const mid = String(mpData[i][mpIdxMid] || '').trim();
        if (mid && !activeMachineIds.has(mid)) mpOrphans++;
      }
    }

    let visitOrphans = 0;
    const vSheet = SS.getSheetByName('VISITAS');
    if (vSheet && vSheet.getLastRow() > 1) {
      const vData    = vSheet.getDataRange().getValues();
      const vHeaders = vData[0];
      const vIdxMid  = vHeaders.indexOf('Machine_ID');
      for (let i = 1; i < vData.length; i++) {
        const mid = String(vData[i][vIdxMid] || '').trim();
        if (mid && !activeMachineIds.has(mid)) visitOrphans++;
      }
    }

    return { machinePartsOrphans: mpOrphans, visitasOrphans: visitOrphans };
  } catch (e) {
    return { machinePartsOrphans: -1, visitasOrphans: -1, error: e.message };
  }
}

// ── Detecta modelos com mesma Marca+Modelo ambos ativos ──────────────────
function _checkModeloDuplicates() {
  try {
    const sheet = SS.getSheetByName('MODELOS');
    if (!sheet) return [];
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxId    = headers.indexOf('ID');
    const idxBrand = headers.indexOf('Marca');
    const idxModel = headers.indexOf('Modelo');
    const idxAtivo = headers.indexOf('Ativo');

    const seen = {};
    for (let i = 1; i < data.length; i++) {
      const ativo = String(data[i][idxAtivo] || 'SIM').toUpperCase();
      if (ativo === 'NÃO') continue;
      const key = String(data[i][idxBrand] || '').trim().toLowerCase()
                + '|' + String(data[i][idxModel] || '').trim().toLowerCase();
      if (!seen[key]) seen[key] = { brand: data[i][idxBrand], model: data[i][idxModel], ids: [] };
      seen[key].ids.push(String(data[i][idxId] || '').trim());
    }

    return Object.values(seen).filter(v => v.ids.length > 1);
  } catch (e) {
    return [];
  }
}

