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
// NOVIDADES v13 (2026-05-26):
//   - version retorna '13' para alinhamento com PCF v13
//   - PCF v13: cache de clientes (pgpEnsureClientsCache) — busca instantânea e completa
//   - PCF v13: regime de trabalho copiado automaticamente para novas máquinas
//   - PCF v13: campos histórico/confiança/fonte removidos do detail de peças
//   - PCF v13: foto do horímetro por peça (hori_p1..p9) com label descritivo no PDF
//
// NOVIDADES v1.8 (2026-05-29):
//   - sameClient(): match por clientKey(nome) OU CNPJ numérico
//   - sameMachine(): match obrigatório por marca+modelo+série/tag (evita falso positivo)
//   - ensureMachineFromVisit: reescrito em 4 passagens; NUNCA sobrescreve Machine_ID existente
//   - saveMachine: usa sameMachine; fallback serial_unique valida marca/modelo antes de aceitar
//   - updateMachineHorímetro: usa sameMachine robusto
//   - updateMachinePartValorMostrado: usa sameMachine + Logger.log de erro
//   - getVisitsNormalized: JOIN com PECAS_LOG — retorna parts por visita para sync do PCM
//
// NOVIDADES v1.7 (2026-05-22):
//   - Identidade canônica de clientes: normalizeTextKey, clientKey, compactKey, isSuspectClientTerm
//   - resolveCanonicalClientName: resolve grafia canônica via CLIENTES → MAQUINAS
//   - getClientsForField: deduplicado por clientKey, filtra termos suspeitos, retorna aliases
//   - getMachinesByClient: igualdade por clientKey (não includes)
//   - searchMachine: sistema de pontuação, série/TAG > cliente > marca/modelo
//   - ensureMachineFromVisit: fallback por série única independente de cliente
//   - saveMachine: resolução canônica + fallback por série única
//   - saveVisit/savePreventiva: resolve cliente canônico antes de calcular machineId
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

// ── IDENTIDADE CANÔNICA v1.7 ─────────────────────────────────────────────
function normalizeTextKey(v) {
  return String(v || '')
    .trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function clientKey(v) {
  return normalizeTextKey(v)
    .replace(/\b(ltda|s\.?a\.?|eireli|me|epp|ss|lda|inc|llc|do brasil|do norte|do sul|e cia|cia)\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function compactKey(v) {
  return normalizeTextKey(v).replace(/[^a-z0-9]/g, '');
}

function isSuspectClientTerm(v) {
  const s = String(v || '').trim();
  if (!s || s.length < 2) return true;
  if (/^[\d\s\-\/]+$/.test(s)) return true;
  if (/^[a-z]{0,2}\d/i.test(s) && s.length < 6) return true;
  if (normalizeTextKey(s).replace(/[^a-z]/g, '').length < 2) return true;
  return false;
}

// Helper: preserva valor 0 (zero é dado válido de horímetro)
function safeVal(v) {
  if (v === undefined || v === null || v === '') return '';
  return v;
}

// Preserva 0 numérico/string em leituras vindas do Sheets.
// Use para Valor_Mostrado, H_Rodadas, H_Restantes e qualquer campo onde zero seja dado válido.
function cellVal(v) {
  if (v === undefined || v === null || v === '') return '';
  return v;
}

function cellStr(v) {
  const val = cellVal(v);
  return val === '' ? '' : String(val);
}

// ── MATCHING ROBUSTO DE CLIENTES (v1.8) ──────────────────────────────────
// Dois clientes são considerados o mesmo se:
//   (a) clientKey(nome) coincidir, OU
//   (b) CNPJ numérico coincidir (quando ambos têm CNPJ)
function sameClient(nomeA, cnpjA, nomeB, cnpjB) {
  const ckA = clientKey(String(nomeA || ''));
  const ckB = clientKey(String(nomeB || ''));
  if (ckA && ckB && ckA === ckB) return true;
  const cnpjNum = v => String(v || '').replace(/\D/g, '');
  const cA = cnpjNum(cnpjA);
  const cB = cnpjNum(cnpjB);
  if (cA.length >= 11 && cB.length >= 11 && cA === cB) return true;
  return false;
}

// ── MATCHING ROBUSTO DE MÁQUINAS (v1.8) ──────────────────────────────────
// Dois registros são a mesma máquina se:
//   OBRIGATÓRIO: compactKey(série) coincidir (quando ambas têm série), OU
//                compactKey(tag)   coincidir (quando ambas têm tag)
//   MAIS:        compactKey(marca) E compactKey(modelo) coincidem (reforço)
//   NUNCA casa apenas por cliente + série sem marca/modelo confirmados
function sameMachine(rBrand, rModel, rSerial, rTag, inBrand, inModel, inSerial, inTag) {
  const ck = v => compactKey(String(v || ''));
  const hasSer = ck(rSerial) && ck(inSerial);
  const hasTag = ck(rTag)    && ck(inTag);
  if (!hasSer && !hasTag) return false; // sem identificador físico → não casa

  const serMatch = hasSer && ck(rSerial) === ck(inSerial);
  const tagMatch = hasTag && ck(rTag)    === ck(inTag);
  if (!serMatch && !tagMatch) return false;

  // Com série/tag coincidindo, exige também marca+modelo para evitar falso positivo
  const brandOk = !ck(rBrand)  || !ck(inBrand)  || ck(rBrand)  === ck(inBrand);
  const modelOk = !ck(rModel)  || !ck(inModel)  || ck(rModel)  === ck(inModel);
  return brandOk && modelOk;
}

function resolveCanonicalClientName(input) {
  const raw = String(input || '').trim();
  const key = clientKey(raw);
  if (!key) return raw;

  try {
    const clientes = getSheetDataActive('CLIENTES');
    const found = clientes.find(c => clientKey(String(c['Nome'] || '')) === key);
    if (found) {
      Logger.log('resolveCanonicalClientName: "' + raw + '" → "' + found['Nome'] + '" (via CLIENTES)');
      return String(found['Nome']).trim();
    }
  } catch(e) {}

  try {
    const maquinas = getSheetDataActive('MAQUINAS');
    const candidates = maquinas
      .filter(m => clientKey(String(m['Cliente'] || '')) === key)
      .map(m => String(m['Cliente'] || '').trim())
      .filter(Boolean);
    if (candidates.length > 0) {
      candidates.sort((a, b) => b.length - a.length);
      Logger.log('resolveCanonicalClientName: "' + raw + '" → "' + candidates[0] + '" (via MAQUINAS)');
      return candidates[0];
    }
  } catch(e) {}

  return raw;
}
// ─────────────────────────────────────────────────────────────────────────

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
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Atualizado',
    // v1.9 — campos coletados pelo PCF que faltavam
    'Potência','Tipo_Equip','Obs_Op'
  ],
  MODELOS: [
    'ID','Marca','Modelo','Tipo','Potência','Pressão','Observações',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Atualizado',
    // v16 — campos técnicos adicionais (ao final, retrocompatível)
    'Vazao_l_min','Tensao','Corrente_A'
  ],
  CLIENTES: [
    'ID','Nome','CNPJ','Cidade','Telefone','Email','Observações',
    'Antecedencia_Alerta_Dias','Drive_URL',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Atualizado',
    // v1.5 — novas colunas ao final (retrocompatível)
    'Contato','Filial','Prox_Visita','Ult_Visita',
    // v1.10 — endereço e UF
    'Endereco','UF'
  ],
  VISITAS: [
    'ID','Machine_ID','Cliente','Filial','Marca','Modelo','Série','TAG',
    'Hor.Visita','h/Semana','Cenário','Técnico','Data Visita','Tipo','Obs.Gerais',
    'Ativo','Tipo_Registro','Created_At','Deleted_At','Deleted_By','Enviado',
    // v1.5 — novas colunas ao final
    'Prox_Visita','Contato','CNPJ',
    // v1.9 — dados de contato e campos operacionais da visita
    'Telefone','Email','Cidade','Obs_Op','Regime_JSON'
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
        result = { status: 'ok', version: '16', spreadsheet: SS.getName(), ts: new Date().toISOString() };
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
      case 'getPartApplications':
        result = getPartApplications(params.q || '');
        break;
      case 'getCatalogAudit':
        result = getCatalogAudit();
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
      case 'deletePartMaster':
        result = deletePartMaster(body.partId, body.modelId, body.deletedBy || 'admin');
        break;
      case 'deleteSubsystemParts':
        result = deleteSubsystemParts(body.modelId, body.subId, body.deletedBy || 'admin');
        break;
      case 'reconcileModelParts':
        result = reconcileModelParts(body.modelId, body.activePartIds || [], body.deletedBy || 'admin');
        break;
      case 'saveClient':
        result = saveClient(body.client);
        break;
      case 'deleteClient':
        result = softDelete('CLIENTES', body.id, body.deletedBy || 'admin');
        break;
      case 'mergeClients':
        result = mergeClients(body);
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
  const qCompact = compactKey(query);
  const qClient  = clientKey(query);
  const qLower   = normalizeTextKey(query);

  const idxSerie  = headers.indexOf('Série');
  const idxTag    = headers.indexOf('TAG');
  const idxClient = headers.indexOf('Cliente');
  const idxId     = headers.indexOf('ID');
  const idxAtivo  = headers.indexOf('Ativo');

  // Calcular score por candidato — série/TAG vencem sobre cliente
  let best = null;
  let bestScore = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[idxAtivo] || 'SIM').toUpperCase() === 'NÃO') continue;

    const serie  = String(row[idxSerie]  || '');
    const tag    = String(row[idxTag]    || '');
    const client = String(row[idxClient] || '');
    const id     = String(row[idxId]     || '');
    const brand  = String(row[headers.indexOf('Marca')]  || '');
    const model  = String(row[headers.indexOf('Modelo')] || '');

    let score = 0;
    if (id.trim() === query.trim())                                          score = 100;
    else if (qCompact && compactKey(serie) === qCompact)                     score = 90;
    else if (qCompact && compactKey(tag)   === qCompact)                     score = 85;
    else if (qClient  && clientKey(client) === qClient)                      score = 70;
    else if (qLower.length > 3 && normalizeTextKey(brand + ' ' + model).includes(qLower)) score = 50;
    else if (qClient.length > 3 && clientKey(client).includes(qClient))     score = 20;

    if (score > bestScore) {
      bestScore = score;
      const machine = {};
      headers.forEach((h, j) => machine[h] = row[j]);
      best = machine;
    }
  }

  if (!best) return { status: 'ok', machine: null };

  const result = rowToMachine(best);
  result.parts = getMachinePartsById(result.id, result.serial, result.tag);
  result.canonicalClient = resolveCanonicalClientName(result.client);
  result.matchScore = bestScore;
  Logger.log('searchMachine: query="' + query + '" → score=' + bestScore + ' cliente="' + result.client + '"');

  // GAS-3: enriquecer com dados do cliente (CNPJ, telefone, cidade, próxima visita…)
  try {
    const clients = getSheetDataActive('CLIENTES');
    const normCl = v => String(v || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    const clientRow = clients.find(c => normCl(c['Nome'] || c['Razao_Social'] || '') === normCl(result.client || ''));
    if (clientRow) {
      const clientData = {
        nome:      clientRow['Nome']        || '',
        cnpj:      clientRow['CNPJ']        || '',
        contato:   clientRow['Contato']     || '',
        telefone:  clientRow['Telefone']    || '',
        email:     clientRow['Email']       || '',
        cidade:    clientRow['Cidade']      || '',
        filial:    clientRow['Filial']      || '',
        nextVisit: clientRow['Prox_Visita'] || '',
        lastVisit: clientRow['Ult_Visita']  || ''
      };
      return { status: 'ok', machine: enrichMachineWithClientData(result, clientData) };
    }
  } catch (enrichErr) {
    Logger.log('searchMachine enrichClient ERRO: ' + enrichErr.message);
  }

  return { status: 'ok', machine: result };
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

  // v1.7: resolver cliente canônico antes de calcular machineId
  if (body.client && !isSuspectClientTerm(body.client)) {
    const canonical = resolveCanonicalClientName(body.client);
    if (canonical !== body.client) {
      Logger.log('saveVisit: cliente corrigido "' + body.client + '" → "' + canonical + '"');
      body.client = canonical;
    }
  }

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
    body.nextVisit || '', body.contact || '', body.cnpj || '',
    // v1.9: dados de contato e campos operacionais
    body.phone || '', body.email || '', body.city || '',
    body.obsOp || '',
    body.regimeData ? JSON.stringify(body.regimeData) : ''
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
      safeVal(ps.valorMostrado),
      safeVal(ps.contador)
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

  // GAS-4: se machine_id ausente, validar que serial/TAG não são ambíguos
  if (!machineId && (serial || tag)) {
    const matchingMids = new Set();
    for (let i = 1; i < data.length; i++) {
      const rowSer = String(data[i][idxSer] || '').trim();
      const rowTag = String(data[i][idxTag] || '').trim();
      if ((serial && rowSer === serial) || (tag && rowTag === tag)) {
        const rowMid = String(data[i][idxMid] || '').trim();
        if (rowMid) matchingMids.add(rowMid);
      }
    }
    if (matchingMids.size > 1) {
      return { status: 'error', error: 'Ambiguidade: múltiplas máquinas com o mesmo serial/TAG' };
    }
  }

  Object.entries(parts).forEach(([partId, ps]) => {
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      const rowMid = String(data[i][idxMid] || '').trim();
      const rowSer = String(data[i][idxSer] || '').trim();
      const rowTag = String(data[i][idxTag] || '').trim();
      const rowPid = String(data[i][idxPid] || '').trim();

      // GAS-4: machine_id tem prioridade; só usa serial/tag se machine_id ausente
      const machineMatch = machineId
        ? rowMid === machineId
        : (serial && rowSer === serial) || (tag && rowTag === tag);
      if (machineMatch && rowPid === partId) {
        rowIdx = i + 1;
        break;
      }
    }

    if (rowIdx > 0) {
      const existingSerial = String(sheet.getRange(rowIdx, idxSer + 1).getValue() || '');
      const existingTag    = String(sheet.getRange(rowIdx, idxTag + 1).getValue() || '');
      const idxVM  = headers.indexOf('Valor_Mostrado');
      const idxCnt = headers.indexOf('Contador');
      const existingVM  = idxVM  >= 0 ? cellStr(sheet.getRange(rowIdx, idxVM  + 1).getValue()) : '';
      const existingCnt = idxCnt >= 0 ? cellStr(sheet.getRange(rowIdx, idxCnt + 1).getValue()) : '';

      const outSerial = serial || existingSerial;
      const outTag    = tag    || existingTag;
      const outVM  = (ps.valorMostrado !== undefined && ps.valorMostrado !== null && ps.valorMostrado !== '')
        ? ps.valorMostrado : existingVM;
      const outCnt = (ps.contador !== undefined && ps.contador !== null && ps.contador !== '')
        ? ps.contador : existingCnt;

      const existingRef = String(sheet.getRange(rowIdx, idxRef + 1).getValue() || '');
      const incomingRef = String(ps.ref || '');
      const refAnterior = (incomingRef && existingRef && incomingRef !== existingRef)
        ? existingRef
        : String(sheet.getRange(rowIdx, idxRefAnterior + 1).getValue() || '');
      const createdAt = String(sheet.getRange(rowIdx, idxCreatedAt + 1).getValue() || '');

      const rowData = [
        machineId, outSerial, outTag,
        partId,
        ps.name     || partId,
        parseInt(ps.lastChange) || 0,
        parseInt(ps.interval)   || 2000,
        incomingRef,
        ps.na  ? 'SIM' : 'NÃO',
        refAnterior, createdAt,
        'SIM', 'PRODUCAO', '', '', now,
        outVM, outCnt
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
        safeVal(ps.valorMostrado),
        safeVal(ps.contador)
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

    const idxAtivo = headers.indexOf('Ativo');
    const result = {};
    for (let i = 1; i < data.length; i++) {
      const rowMid = cellStr(data[i][idxMid]).trim();
      const rowSer = cellStr(data[i][idxSer]).trim();
      const rowTag = cellStr(data[i][idxTag]).trim();
      const rowPid = cellStr(data[i][idxPid]).trim();

      if (idxAtivo >= 0 && String(data[i][idxAtivo] || '').toUpperCase() === 'NÃO') continue;

      const hasMachineId = String(machineId || '').trim();
      const match = hasMachineId
        ? rowMid === hasMachineId
        : ((serial && rowSer === String(serial).trim()) ||
           (tag    && rowTag === String(tag).trim()));
      if (match && rowPid) {
        result[rowPid] = {
          name:          String(data[i][idxName] || rowPid),
          lastChange:    parseInt(data[i][idxLch]) || 0,
          interval:      parseInt(data[i][idxInt]) || 2000,
          ref:           String(data[i][idxRef] || ''),
          na:            String(data[i][idxNA] || '').toUpperCase() === 'SIM',
          valorMostrado: idxVM  >= 0 ? cellStr(data[i][idxVM])  : '',
          contador:      idxCnt >= 0 ? cellStr(data[i][idxCnt]) : ''
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

  // v1.8: JOIN com PECAS_LOG para retornar parts por visita
  const partsMap = {};
  try {
    const plSheet = SS.getSheetByName('PECAS_LOG');
    if (plSheet) {
      const plData = plSheet.getDataRange().getValues();
      const plH = plData[0];
      const piVid  = plH.indexOf('ID_Visita');
      const piPid  = plH.indexOf('ID_Peça');
      const piName = plH.indexOf('Nome Peça');
      const piRef  = plH.indexOf('Ref_Nova');
      const piLch  = plH.indexOf('Últ.Troca(h)');
      const piInt  = plH.indexOf('Intervalo_H');
      const piNA   = plH.indexOf('N/A');
      const piVM   = plH.indexOf('Valor_Mostrado');
      const piCnt  = plH.indexOf('Tipo_Contador');
      const piHRod = plH.indexOf('H_Rodadas');
      const piHRes = plH.indexOf('H_Restantes');
      const piStat = plH.indexOf('Status');
      const piAcao = plH.indexOf('Acao');
      for (let r = 1; r < plData.length; r++) {
        const vid = cellStr(plData[r][piVid]).trim();
        const pid = cellStr(plData[r][piPid]).trim();
        if (!vid || !pid) continue;
        if (!partsMap[vid]) partsMap[vid] = {};
        partsMap[vid][pid] = {
          name:          piName >= 0 ? String(plData[r][piName] || pid) : pid,
          ref:           piRef  >= 0 ? String(plData[r][piRef]  || '') : '',
          lastChange:    piLch  >= 0 ? (parseInt(plData[r][piLch])  || 0) : 0,
          interval:      piInt  >= 0 ? (parseInt(plData[r][piInt])  || 0) : 0,
          na:            piNA   >= 0 && String(plData[r][piNA] || '').toUpperCase() === 'SIM',
          valorMostrado: piVM   >= 0 ? cellStr(plData[r][piVM])   : '',
          contador:      piCnt  >= 0 ? cellStr(plData[r][piCnt])  : '',
          horasRodadas:  piHRod >= 0 ? cellStr(plData[r][piHRod]) : '',
          horasRestantes:piHRes >= 0 ? cellStr(plData[r][piHRes]) : '',
          status:        piStat >= 0 ? cellStr(plData[r][piStat]) : '',
          acao:          piAcao >= 0 ? cellStr(plData[r][piAcao]) : '',
        };
      }
    }
  } catch(e) { Logger.log('getVisitsNormalized: erro ao ler PECAS_LOG — ' + e.message); }

  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, j) => obj[h] = row[j]);
    const vid = String(obj['ID'] || '').trim();
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
      // v1.9: dados de contato e campos operacionais
      phone:      obj['Telefone']     || '',
      email:      obj['Email']        || '',
      city:       obj['Cidade']       || '',
      obsOp:      obj['Obs_Op']       || '',
      regimeData: (()=>{ try{ const v=obj['Regime_JSON']; return v?JSON.parse(v):null; }catch(e){return null;} })(),
      // v1.8: partes da visita via JOIN com PECAS_LOG
      parts:      partsMap[vid] || null,
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

  // v1.7: resolver cliente canônico antes de calcular machineId
  if (body.client && !isSuspectClientTerm(body.client)) {
    const canonical = resolveCanonicalClientName(body.client);
    if (canonical !== body.client) {
      Logger.log('savePreventiva: cliente corrigido "' + body.client + '" → "' + canonical + '"');
      body.client = canonical;
    }
  }

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
    body.nextVisit || '', body.contact || '', body.cnpj || '',
    // v1.9: dados de contato e campos operacionais
    body.phone || '', body.email || '', body.city || '',
    body.obsOp || '',
    body.regimeData ? JSON.stringify(body.regimeData) : ''
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
      safeVal(ps.valorMostrado),
      safeVal(ps.contador)
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
        lastChange: parseInt(ps.lastChange) || 0,
        valorMostrado: (ps.valorMostrado !== undefined && ps.valorMostrado !== null) ? ps.valorMostrado : '',
        contador: ps.contador || '',
        now
      });
      pecasTrocadas++;
    } else if (acao === 'conferida' || acao === 'na') {
      // Não altera lastChange nem ref, mas persiste valorMostrado e contador
      updateMachinePartValorMostrado(machineId, body.serial || '', body.tag || '', partId, safeVal(ps.valorMostrado), safeVal(ps.contador), now);
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
      horasRodadas:   cellVal(p['H_Rodadas']),
      horasRestantes: cellVal(p['H_Restantes']),
      status:         cellVal(p['Status']),
      valorMostrado:  cellVal(p['Valor_Mostrado']),
      tipoContador:   cellVal(p['Tipo_Contador']),
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

  // v1.7: resolver canônico e usar igualdade por clientKey
  const canonical = resolveCanonicalClientName(q);
  const qKey = clientKey(canonical || q);

  const machines = getSheetDataActive('MAQUINAS')
    .filter(m => clientKey(String(m['Cliente'] || '')) === qKey)
    .map(m => {
      const eq = rowToMachineFromObj(m);
      eq.parts = getMachinePartsById(eq.id, eq.serial, eq.tag);
      return eq;
    });

  // v1.5: buscar dados do cliente para o PCF pré-preencher campos cadastrais
  const clientsData = getSheetDataActive('CLIENTES');
  const matchClient = clientsData.find(c => clientKey(String(c['Nome'] || '')) === qKey);
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
  Logger.log('getMachinesByClient: query="' + q + '" canonical="' + canonical + '" encontradas=' + enrichedMachines.length);
  return { status: 'ok', machines: enrichedMachines, clientData, canonicalClient: canonical };
}

// ═══════════════════════════════════════════════════════════════════════════
// GET CLIENTS FOR FIELD
// ═══════════════════════════════════════════════════════════════════════════
function getClientsForField() {
  const clientsData  = getSheetDataActive('CLIENTES');
  const machinesData = getSheetDataActive('MAQUINAS');

  // v1.7: mapa clientKey → nome canônico (CLIENTES é fonte primária)
  const canonicalMap = {};
  clientsData.forEach(c => {
    const nome = String(c['Nome'] || '').trim();
    if (!nome || isSuspectClientTerm(nome)) return;
    const k = clientKey(nome);
    if (k && !canonicalMap[k]) canonicalMap[k] = nome;
  });

  // Complementar com MAQUINAS apenas para chaves ainda não vistas
  machinesData.forEach(m => {
    const nome = String(m['Cliente'] || '').trim();
    if (!nome || isSuspectClientTerm(nome)) return;
    const k = clientKey(nome);
    if (k && !canonicalMap[k]) canonicalMap[k] = nome;
  });

  const unique = Object.values(canonicalMap)
    .sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));

  // Aliases: mapa de clientKey → variantes (para debug/merge no PCM)
  const aliasMap = {};
  [...clientsData, ...machinesData].forEach(row => {
    const nome = String(row['Nome'] || row['Cliente'] || '').trim();
    if (!nome || isSuspectClientTerm(nome)) return;
    const k = clientKey(nome);
    if (!k) return;
    if (!aliasMap[k]) aliasMap[k] = new Set();
    aliasMap[k].add(nome);
  });
  const aliases = {};
  Object.entries(aliasMap).forEach(([k, s]) => {
    if (s.size > 1) aliases[k] = [...s];
  });

  // clientsFull — somente clientes da aba CLIENTES
  const clientsFull = clientsData
    .filter(c => String(c['Nome'] || '').trim() && !isSuspectClientTerm(c['Nome']))
    .map(c => ({
      nome:      String(c['Nome']        || '').trim(),
      cnpj:      String(c['CNPJ']        || '').trim(),
      cidade:    String(c['Cidade']      || '').trim(),
      telefone:  String(c['Telefone']    || '').trim(),
      email:     String(c['Email']       || '').trim(),
      contato:   String(c['Contato']     || '').trim(),
      filial:    String(c['Filial']      || '').trim(),
      nextVisit: String(c['Prox_Visita'] || '').trim(),
      lastVisit: String(c['Ult_Visita']  || '').trim(),
    }));

  return { status: 'ok', clients: unique, clientsFull, aliases };
}


// ═══════════════════════════════════════════════════════════════════════════
// AUDITORIA — getSystemHealth
// ═══════════════════════════════════════════════════════════════════════════
function getSystemHealth() {
  const tabelas = ['MAQUINAS', 'MODELOS', 'CLIENTES', 'VISITAS', 'PECAS_LOG', 'MACHINE_PARTS', 'PARTS_MASTER', 'PART_SIMILARITIES'];
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
  const clientsData = getSheetData('CLIENTES');
  const clientByKey = {};
  const clientByCnpj = {};
  clientsData.forEach(c => {
    const nome = String(c['Nome'] || '').trim();
    const cnpj = String(c['CNPJ'] || '').replace(/\D/g, '');
    const ativo = String(c['Ativo'] || 'SIM').toUpperCase();
    if (ativo === 'NÃO') return;
    const k = clientKey(nome);
    if (k) { if (!clientByKey[k]) clientByKey[k] = []; clientByKey[k].push(nome); }
    if (cnpj.length >= 11) { if (!clientByCnpj[cnpj]) clientByCnpj[cnpj] = []; clientByCnpj[cnpj].push(nome); }
  });
  const dupClientsByKey = Object.entries(clientByKey).filter(([, v]) => v.length > 1);
  const dupClientsByCnpj = Object.entries(clientByCnpj).filter(([, v]) => v.length > 1);

  return {
    status: 'ok',
    ts: new Date().toISOString(),
    duplicatas: {
      por_machine_id:    porMachineId,
      por_cliente_serie: porClienteSerial,
      por_cliente_tag:   porClienteTag,
      por_cliente_nome_canonico: dupClientsByKey,
      por_cliente_cnpj: dupClientsByCnpj
    },
    total_duplicatas: porMachineId.length + porClienteSerial.length + porClienteTag.length + dupClientsByKey.length + dupClientsByCnpj.length
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
  const tabelas = ['MAQUINAS', 'MODELOS', 'CLIENTES', 'VISITAS', 'PECAS_LOG', 'MACHINE_PARTS', 'PARTS_MASTER', 'PART_SIMILARITIES'];
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
    const mid = String(machineId || '').trim();
    const match = mid
      ? rowMid === mid
      : (serial && rowSer === String(serial).trim()) || (tag && rowTag === String(tag).trim());
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
    // GAS-4: machine_id tem prioridade; só usa serial/tag se machine_id ausente
    const match = machineId
      ? rowMid === machineId
      : (serial && rowSer === serial) || (tag && rowTag === tag);
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
        (payload.lastChange !== undefined && payload.lastChange !== null && payload.lastChange !== '')
          ? (parseInt(payload.lastChange) || 0)
          : (parseInt(payload.hourTotal) || 0),
        parseInt(payload.interval) || 2000,
        String(payload.refNova || '').trim(),
        String(payload.na || '').toUpperCase() === 'SIM' ? 'SIM' : 'NÃO',
        refAnterior,
        createdAt || (payload.now || new Date().toISOString()),
        'SIM', 'PRODUCAO', '', '', payload.now || new Date().toISOString(),
        (payload.valorMostrado !== undefined && payload.valorMostrado !== null) ? payload.valorMostrado : '',
        payload.contador || ''
      ];
      sheet.getRange(i + 1, 1, 1, HEADERS.MACHINE_PARTS.length).setValues([updated]);
      return;
    }
  }

  sheet.appendRow([
    machineId, serial, tag,
    partId,
    payload.name || partId,
    (payload.lastChange !== undefined && payload.lastChange !== null && payload.lastChange !== '')
      ? (parseInt(payload.lastChange) || 0)
      : (parseInt(payload.hourTotal) || 0),
    parseInt(payload.interval) || 2000,
    String(payload.refNova || '').trim(),
    String(payload.na || '').toUpperCase() === 'SIM' ? 'SIM' : 'NÃO',
    String(payload.refAnterior || '').trim(),
    payload.now || new Date().toISOString(),
    'SIM', 'PRODUCAO', '', '', payload.now || new Date().toISOString(),
    (payload.valorMostrado !== undefined && payload.valorMostrado !== null) ? payload.valorMostrado : '',
    payload.contador || ''
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
    if (idxVM < 0 || idxCnt < 0) {
      Logger.log('updateMachinePartValorMostrado: colunas Valor_Mostrado/Contador não encontradas');
      return;
    }
    for (let i = 1; i < data.length; i++) {
      const rowMid = String(data[i][idxMid] || '').trim();
      const rowSer = String(data[i][idxSer] || '').trim();
      const rowTag = String(data[i][idxTag] || '').trim();
      const rowPid = String(data[i][idxPid] || '').trim();
      const mid = String(machineId || '').trim();
      const match = mid
        ? rowMid === mid
        : (serial && rowSer === String(serial).trim()) || (tag && rowTag === String(tag).trim());
      if (match && rowPid === String(partId).trim()) {
        sheet.getRange(i + 1, idxVM  + 1).setValue(safeVal(valorMostrado));
        sheet.getRange(i + 1, idxCnt + 1).setValue(safeVal(contador));
        if (idxAtual >= 0) sheet.getRange(i + 1, idxAtual + 1).setValue(now || new Date().toISOString());
        return;
      }
    }
    Logger.log('updateMachinePartValorMostrado: linha não encontrada para machineId=' + machineId + ' partId=' + partId);
  } catch(e) {
    Logger.log('updateMachinePartValorMostrado ERRO: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SAVE / UPDATE MACHINE
// ═══════════════════════════════════════════════════════════════════════════
function saveMachine(m) {
  if (!m) return { status: 'error', error: 'Dados ausentes' };
  const sheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const now = new Date().toISOString();
  const ad = auditDefaults(m.tipoRegistro);

  // v1.7: resolver cliente canônico antes de qualquer comparação
  if (m.client && !isSuspectClientTerm(m.client)) {
    m.client = resolveCanonicalClientName(m.client) || m.client;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId    = headers.indexOf('ID');
  const idxSerie = headers.indexOf('Série');
  const idxAtivo = headers.indexOf('Ativo');

  for (let i = 1; i < data.length; i++) {
    const idxCli = headers.indexOf('Cliente');
    const idxMar = headers.indexOf('Marca');
    const idxMod = headers.indexOf('Modelo');
    const idxTag = headers.indexOf('TAG');
    // v1.8: match por ID exato OU por marca+modelo+série (sameMachine robusto)
    const sameId = String(data[i][idxId] || '').trim() === String(m.id || '').trim() && m.id;
    const machineMatch = !sameId && sameMachine(
      String(data[i][idxMar] || ''), String(data[i][idxMod] || ''),
      String(data[i][idxSerie] || ''), String(data[i][idxTag] || ''),
      m.brand || '', m.model || '', m.serial || '', m.tag || ''
    );
    if (sameId || machineMatch) {
      const idxCreated = headers.indexOf('Created_At');
      const idxTipo    = headers.indexOf('Tipo_Registro');
      const originalCreated = String(data[i][idxCreated] || now);
      const originalTipo    = String(data[i][idxTipo]    || 'PRODUCAO');
      const originalAtivo   = String(data[i][idxAtivo]   || 'SIM');
      const existingId      = String(data[i][idxId]      || '').trim();
      sheet.getRange(i+1, 1, 1, HEADERS.MAQUINAS.length).setValues([[
        existingId || m.id || '',   // NUNCA sobrescreve ID existente com diferente
        m.client || '', m.branch || '',
        m.brand  || '', m.model  || '',
        m.serial || '', m.year   || '',
        m.tag    || '', m.location || '',
        parseInt(m.hourTotal) || 0,
        parseInt(m.hpw)       || 0,
        m.obs || '',
        originalAtivo, originalTipo, originalCreated, '', '', now,
        // v1.9:
        m.power || m.potencia || '', m.type || m.tipoEquip || '', m.obsOp || ''
      ]]);
      return { status: 'ok', action: 'updated', matchedBy: sameId ? 'id' : 'brand_model_serial' };
    }
  }

  // v1.8: fallback por série única — só aceita se marca+modelo forem compatíveis
  if (m.serial) {
    const sk = compactKey(m.serial);
    let uniqIdx = -1, uniqCount = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxAtivo] || 'SIM').toUpperCase() === 'NÃO') continue;
      if (compactKey(String(data[i][idxSerie] || '')) === sk) {
        uniqIdx = i; uniqCount++;
        if (uniqCount > 1) break;
      }
    }
    if (uniqCount === 1) {
      const i = uniqIdx;
      const idxMar = headers.indexOf('Marca');
      const idxMod = headers.indexOf('Modelo');
      const rowBrand = String(data[i][idxMar] || '');
      const rowModel = String(data[i][idxMod] || '');
      const brandOk  = !compactKey(rowBrand) || !compactKey(m.brand || '') || compactKey(rowBrand) === compactKey(m.brand || '');
      const modelOk  = !compactKey(rowModel) || !compactKey(m.model || '') || compactKey(rowModel) === compactKey(m.model || '');
      if (!brandOk || !modelOk) {
        Logger.log('saveMachine: serial_unique REJEITADO por marca/modelo incompatível. Row:' + rowBrand + '/' + rowModel + ' vs ' + m.brand + '/' + m.model);
        // não aplica — cai para insert
      } else {
        const idxCreated = headers.indexOf('Created_At');
        const idxTipo    = headers.indexOf('Tipo_Registro');
        const existingId     = String(data[i][idxId] || '').trim();
        const existingClient = String(data[i][headers.indexOf('Cliente')] || '').trim();
        Logger.log('saveMachine: serial_unique match — cliente existente="' + existingClient + '" recebido="' + m.client + '"');
        sheet.getRange(i+1, 1, 1, HEADERS.MAQUINAS.length).setValues([[
          existingId || m.id || '',   // preserva ID existente
          existingClient || m.client || '',
          m.branch || '', m.brand || '', m.model || '',
          m.serial || '', m.year || '', m.tag || '', m.location || '',
          parseInt(m.hourTotal) || 0, parseInt(m.hpw) || 0, m.obs || '',
          String(data[i][idxAtivo] || 'SIM'),
          String(data[i][idxTipo]  || 'PRODUCAO'),
          String(data[i][idxCreated] || now),
          '', '', now,
          // v1.9:
          m.power || m.potencia || '', m.type || m.tipoEquip || '', m.obsOp || ''
        ]]);
        return { status: 'ok', action: 'updated', matchedBy: 'serial_unique', clientCorrected: existingClient !== m.client };
      }
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
    ad.ativo, ad.tipoRegistro, ad.createdAt, ad.deletedAt, ad.deletedBy, now,
    // v1.9:
    m.power || m.potencia || '', m.type || m.tipoEquip || '', m.obsOp || ''
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
      const existRow   = {};
      headers.forEach((h, j) => { existRow[h] = data[i][j]; });
      sheet.getRange(i+1, 1, 1, HEADERS.MODELOS.length).setValues([[
        m.id, m.brand || '', m.model || '', m.type || '',
        m.power || '', m.pressure || '', m.obs || '',
        String(data[i][idxAtivo]   || 'SIM'),
        String(data[i][idxTipo]    || 'PRODUCAO'),
        String(data[i][idxCreated] || now),
        '', '', now,
        m.flow    || m.mFlow    || m.vazao    || existRow['Vazao_l_min'] || '',
        m.voltage || m.mVoltage || m.tensao   || existRow['Tensao']      || '',
        m.ampere  || m.mAmpere  || m.corrente || existRow['Corrente_A']  || ''
      ]]);
      return { status: 'ok', action: 'updated' };
    }
  }

  sheet.appendRow([
    m.id || 'MOD-' + new Date().getTime(),
    m.brand || '', m.model || '', m.type || '',
    m.power || '', m.pressure || '', m.obs || '',
    ad.ativo, ad.tipoRegistro, ad.createdAt, ad.deletedAt, ad.deletedBy, now,
    m.flow    || m.mFlow    || m.vazao    || '',
    m.voltage || m.mVoltage || m.tensao   || '',
    m.ampere  || m.mAmpere  || m.corrente || ''
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

  const idxAtivo = headers.indexOf('Ativo');
  let activeRowIdx = -1;
  let deletedRowIdx = -1;

  for (let i = 1; i < data.length; i++) {
    const samePartId  = String(data[i][idxPartId]  || '').trim() === String(p.id       || '').trim();
    const sameModelId = String(data[i][idxModelId] || '').trim() === String(p.modelId  || '').trim();
    if (!samePartId || !sameModelId) continue;
    const isDeleted = idxAtivo >= 0 && String(data[i][idxAtivo] || '').trim().toUpperCase() === 'NÃO';
    if (isDeleted) {
      if (deletedRowIdx < 0) deletedRowIdx = i; // registra primeiro deletado
    } else {
      activeRowIdx = i; // linha ativa encontrada
      break;
    }
  }

  if (activeRowIdx >= 0) {
    // Atualiza a linha ativa existente no lugar
    sheet.getRange(activeRowIdx + 1, 1, 1, HEADERS.PARTS_MASTER.length).setValues([row]);
    return { status: 'ok', action: 'updated' };
  }

  if (deletedRowIdx >= 0) {
    // Existe linha soft-deletada para este Part_ID+Model_ID.
    // NUNCA criar nova linha — isso causaria ressurreição da peça.
    // O reconcileModelParts que roda logo após cobre a limpeza.
    return { status: 'ok', action: 'skipped_deleted' };
  }

  // Nenhuma linha encontrada — inserir nova
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

function deletePartMaster(partId, modelId, deletedBy) {
  if (!partId) return { status: 'error', error: 'partId obrigatório' };
  if (!modelId) return { status: 'error', error: 'modelId obrigatório' };

  const now = new Date().toISOString();
  const actor = deletedBy || 'admin';
  const partsSheet = getOrCreateSheet('PARTS_MASTER', HEADERS.PARTS_MASTER);
  const partsData = partsSheet.getDataRange().getValues();
  const ph = partsData[0] || [];
  const pIdxPartId = ph.indexOf('Part_ID');
  const pIdxModelId = ph.indexOf('Model_ID');
  const pIdxAtivo = ph.indexOf('Ativo');
  const pIdxDeletedAt = ph.indexOf('Deleted_At');
  const pIdxDeletedBy = ph.indexOf('Deleted_By');
  const pIdxUpdatedAt = ph.indexOf('Updated_At');
  let partsInativadas = 0;

  for (let i = 1; i < partsData.length; i++) {
    const samePart = String(partsData[i][pIdxPartId] || '').trim() === String(partId).trim();
    const sameModel = String(partsData[i][pIdxModelId] || '').trim() === String(modelId).trim();
    if (samePart && sameModel) {
      if (pIdxAtivo >= 0) partsSheet.getRange(i + 1, pIdxAtivo + 1).setValue('NÃO');
      if (pIdxDeletedAt >= 0) partsSheet.getRange(i + 1, pIdxDeletedAt + 1).setValue(now);
      if (pIdxDeletedBy >= 0) partsSheet.getRange(i + 1, pIdxDeletedBy + 1).setValue(actor);
      if (pIdxUpdatedAt >= 0) partsSheet.getRange(i + 1, pIdxUpdatedAt + 1).setValue(now);
      partsInativadas++;
    }
  }

  const simResult = deletePartSimilaritiesByPartModel(partId, modelId, actor, now);
  return { status: 'ok', action: 'soft_deleted', partsInativadas: partsInativadas, similaridadesInativadas: simResult.count };
}

function deleteSubsystemParts(modelId, subId, deletedBy) {
  if (!modelId) return { status: 'error', error: 'modelId obrigatório' };
  if (!subId) return { status: 'error', error: 'subId obrigatório' };

  const now = new Date().toISOString();
  const actor = deletedBy || 'admin';
  const sheet = getOrCreateSheet('PARTS_MASTER', HEADERS.PARTS_MASTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxModelId = headers.indexOf('Model_ID');
  const idxScope = headers.indexOf('Part_Scope');
  const idxSubId = headers.indexOf('Sub_ID');
  const idxPartId = headers.indexOf('Part_ID');
  const idxAtivo = headers.indexOf('Ativo');
  const idxDeletedAt = headers.indexOf('Deleted_At');
  const idxDeletedBy = headers.indexOf('Deleted_By');
  const idxUpdatedAt = headers.indexOf('Updated_At');
  const affectedParts = {};
  let partsInativadas = 0;

  for (let i = 1; i < data.length; i++) {
    const sameModel = String(data[i][idxModelId] || '').trim() === String(modelId).trim();
    const isSub = String(data[i][idxScope] || '').trim().toLowerCase() === 'sub';
    const sameSub = String(data[i][idxSubId] || '').trim() === String(subId).trim();
    if (sameModel && isSub && sameSub) {
      if (idxAtivo >= 0) sheet.getRange(i + 1, idxAtivo + 1).setValue('NÃO');
      if (idxDeletedAt >= 0) sheet.getRange(i + 1, idxDeletedAt + 1).setValue(now);
      if (idxDeletedBy >= 0) sheet.getRange(i + 1, idxDeletedBy + 1).setValue(actor);
      if (idxUpdatedAt >= 0) sheet.getRange(i + 1, idxUpdatedAt + 1).setValue(now);
      affectedParts[String(data[i][idxPartId] || '').trim()] = true;
      partsInativadas++;
    }
  }

  let similaritiesInativadas = 0;
  Object.keys(affectedParts).forEach(function(pid) {
    if (pid) similaritiesInativadas += deletePartSimilaritiesByPartModel(pid, modelId, actor, now).count;
  });

  return { status: 'ok', action: 'soft_deleted', partsInativadas: partsInativadas, similaridadesInativadas: similaritiesInativadas };
}

function reconcileModelParts(modelId, activePartIds, deletedBy) {
  if (!modelId) return { status: 'error', error: 'modelId obrigatório' };

  const now = new Date().toISOString();
  const actor = deletedBy || 'admin';
  const activeMap = {};
  (activePartIds || []).forEach(function(pid) { activeMap[String(pid).trim()] = true; });

  const sheet = getOrCreateSheet('PARTS_MASTER', HEADERS.PARTS_MASTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxPartId = headers.indexOf('Part_ID');
  const idxModelId = headers.indexOf('Model_ID');
  const idxAtivo = headers.indexOf('Ativo');
  const idxDeletedAt = headers.indexOf('Deleted_At');
  const idxDeletedBy = headers.indexOf('Deleted_By');
  const idxUpdatedAt = headers.indexOf('Updated_At');
  let inativadas = 0;
  let simsInativadas = 0;

  for (let i = 1; i < data.length; i++) {
    const sameModel = String(data[i][idxModelId] || '').trim() === String(modelId).trim();
    if (!sameModel) continue;
    const pid = String(data[i][idxPartId] || '').trim();
    if (!pid || activeMap[pid]) continue;

    if (idxAtivo >= 0) sheet.getRange(i + 1, idxAtivo + 1).setValue('NÃO');
    if (idxDeletedAt >= 0) sheet.getRange(i + 1, idxDeletedAt + 1).setValue(now);
    if (idxDeletedBy >= 0) sheet.getRange(i + 1, idxDeletedBy + 1).setValue(actor);
    if (idxUpdatedAt >= 0) sheet.getRange(i + 1, idxUpdatedAt + 1).setValue(now);
    simsInativadas += deletePartSimilaritiesByPartModel(pid, modelId, actor, now).count;
    inativadas++;
  }

  return { status: 'ok', action: 'reconciled', partsInativadas: inativadas, similaridadesInativadas: simsInativadas };
}

function deletePartSimilaritiesByPartModel(partId, modelId, deletedBy, nowOpt) {
  const now = nowOpt || new Date().toISOString();
  const actor = deletedBy || 'admin';
  const sheet = getOrCreateSheet('PART_SIMILARITIES', HEADERS.PART_SIMILARITIES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxPid = headers.indexOf('Part_ID');
  const idxMid = headers.indexOf('Model_ID');
  const idxAtivo = headers.indexOf('Ativo');
  const idxDeletedAt = headers.indexOf('Deleted_At');
  const idxDeletedBy = headers.indexOf('Deleted_By');
  const idxUpdatedAt = headers.indexOf('Updated_At');
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const samePart = String(data[i][idxPid] || '').trim() === String(partId).trim();
    const sameModel = String(data[i][idxMid] || '').trim() === String(modelId).trim();
    if (samePart && sameModel) {
      if (idxAtivo >= 0) sheet.getRange(i + 1, idxAtivo + 1).setValue('NÃO');
      if (idxDeletedAt >= 0) sheet.getRange(i + 1, idxDeletedAt + 1).setValue(now);
      if (idxDeletedBy >= 0) sheet.getRange(i + 1, idxDeletedBy + 1).setValue(actor);
      if (idxUpdatedAt >= 0) sheet.getRange(i + 1, idxUpdatedAt + 1).setValue(now);
      count++;
    }
  }
  return { count: count };
}

// ═══════════════════════════════════════════════════════════════════════════
// GET CATALOG FULL
// ═══════════════════════════════════════════════════════════════════════════
function getCatalogFull() {
  ensureSheetHeaders('MODELOS',           HEADERS.MODELOS);
  ensureSheetHeaders('PARTS_MASTER',      HEADERS.PARTS_MASTER);
  ensureSheetHeaders('PART_SIMILARITIES', HEADERS.PART_SIMILARITIES);
  return {
    status:       'ok',
    models:       getSheetDataActive('MODELOS'),
    parts:        getSheetDataActive('PARTS_MASTER'),
    similarities: getSheetDataActive('PART_SIMILARITIES')
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
      // v1.10: endereço e UF
      mergeVal(c.endereco,     ex['Endereco']),
      mergeVal(c.uf,           ex['UF']),
    ];
  };

  const idxNome = headers.indexOf('Nome');
  const idxCnpj = headers.indexOf('CNPJ');
  const normNovo = clientKey(c.nome || '');
  const onlyDigits = v => String(v || '').replace(/\D/g, '');
  const cnpjNovo = onlyDigits(c.cnpj);
  let _foundRow = -1;
  let _matchType = '';

  for (let i = 1; i < data.length; i++) {
    const rowId   = String(data[i][idxId]   || '').trim();
    const rowNome = idxNome >= 0 ? String(data[i][idxNome] || '').trim() : '';
    if (rowId && rowId === String(c.id || '').trim()) {
      _foundRow = i; _matchType = 'id'; break;
    }
    const rowCnpj = idxCnpj >= 0 ? onlyDigits(data[i][idxCnpj]) : '';
    if (!_matchType && cnpjNovo && rowCnpj && cnpjNovo === rowCnpj) {
      _foundRow = i; _matchType = 'cnpj';
    }
    if (normNovo && clientKey(rowNome) === normNovo && _foundRow < 0) {
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

function mergeClients(body) {
  const { sourceId, sourceName, targetId, targetName, client, deletedBy } = body || {};
  const now = new Date().toISOString();
  Logger.log('[GAS mergeClients] source: ' + sourceName + ' (' + sourceId + ') → target: ' + targetName + ' (' + targetId + ')');
  const clientSheet = getOrCreateSheet('CLIENTES', HEADERS.CLIENTES);
  const clientData = clientSheet.getDataRange().getValues();
  const cHeaders = clientData[0];
  const idxCId = cHeaders.indexOf('ID');
  const idxCNome = cHeaders.indexOf('Nome');
  const idxAtivo = cHeaders.indexOf('Ativo');
  const idxDelAt = cHeaders.indexOf('Deleted_At');
  const idxDelBy = cHeaders.indexOf('Deleted_By');
  const ck = v => clientKey(String(v || ''));
  let targetRow = -1, sourceRow = -1;
  for (let i = 1; i < clientData.length; i++) {
    const rowId = String(clientData[i][idxCId] || '').trim();
    const rowNome = String(clientData[i][idxCNome] || '').trim();
    if (rowId === String(targetId || '').trim() || ck(rowNome) === ck(targetName)) if (targetRow < 0) targetRow = i;
    if (rowId === String(sourceId || '').trim() || rowNome === sourceName) if (sourceRow < 0) sourceRow = i;
  }
  if (targetRow >= 0 && client) {
    const existingObj = {};
    cHeaders.forEach((h, i) => { existingObj[h] = clientData[targetRow][i]; });
    const mergeVal = (nv, ev) => (nv && String(nv).trim()) ? String(nv).trim() : String(ev || '');
    const updatedRow = cHeaders.map(h => {
      switch (h) {
        case 'Nome': return mergeVal(client.nome, existingObj['Nome']);
        case 'CNPJ': return mergeVal(client.cnpj, existingObj['CNPJ']);
        case 'Cidade': return mergeVal(client.cidade, existingObj['Cidade']);
        case 'Telefone': return mergeVal(client.telefone, existingObj['Telefone']);
        case 'Email': return mergeVal(client.email, existingObj['Email']);
        case 'Drive_URL': return mergeVal(client.drive_url, existingObj['Drive_URL']);
        case 'Observações': return mergeVal(client.observacoes, existingObj['Observações']);
        case 'Atualizado': return now;
        default: return existingObj[h] !== undefined ? existingObj[h] : '';
      }
    });
    clientSheet.getRange(targetRow + 1, 1, 1, cHeaders.length).setValues([updatedRow]);
  }
  let sourceDeactivated = false;
  if (sourceRow >= 0 && sourceRow !== targetRow) {
    if (idxAtivo >= 0) clientSheet.getRange(sourceRow + 1, idxAtivo + 1).setValue('NÃO');
    if (idxDelAt >= 0) clientSheet.getRange(sourceRow + 1, idxDelAt + 1).setValue(now);
    if (idxDelBy >= 0) clientSheet.getRange(sourceRow + 1, idxDelBy + 1).setValue(deletedBy || 'merge');
    sourceDeactivated = true;
  }
  const machSheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const machData = machSheet.getDataRange().getValues();
  const mHeaders = machData[0];
  const idxMCli = mHeaders.indexOf('Cliente');
  const isSuspect = isSuspectClientTerm(sourceName);
  let machinesUpdated = 0;
  for (let i = 1; i < machData.length; i++) {
    const rowCli = String(machData[i][idxMCli] || '').trim();
    const matches = isSuspect ? (rowCli === sourceName) : (ck(rowCli) === ck(sourceName) || rowCli === sourceName);
    if (matches) { machSheet.getRange(i + 1, idxMCli + 1).setValue(targetName); machinesUpdated++; }
  }
  const visitSheet = getOrCreateSheet('VISITAS', HEADERS.VISITAS);
  const visitData = visitSheet.getDataRange().getValues();
  const vHeaders = visitData[0];
  const idxVCli = vHeaders.indexOf('Cliente');
  let visitsUpdated = 0;
  if (idxVCli >= 0) for (let i = 1; i < visitData.length; i++) {
    const rowCli = String(visitData[i][idxVCli] || '').trim();
    const matches = isSuspect ? (rowCli === sourceName) : (ck(rowCli) === ck(sourceName) || rowCli === sourceName);
    if (matches) { visitSheet.getRange(i + 1, idxVCli + 1).setValue(targetName); visitsUpdated++; }
  }
  Logger.log('[GAS mergeClients] machinesUpdated: ' + machinesUpdated + ', visitsUpdated: ' + visitsUpdated + ', sourceDeactivated: ' + sourceDeactivated);
  return { status: 'ok', merged: true, targetId, targetName, sourceId, sourceName, machinesUpdated, visitsUpdated, sourceDeactivated };
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
  const idxAtivo = headers.indexOf('Ativo');
  // v1.9: novos campos de MAQUINAS
  const idxPot = headers.indexOf('Potência');
  const idxTipoEq = headers.indexOf('Tipo_Equip');
  const idxObsOp  = headers.indexOf('Obs_Op');

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
  const potencia  = body.potencia  || '';
  const tipoEquip = body.tipoEquip || '';  // tipo de equipamento (ex: Compressor Parafuso) — separado do tipo de visita
  const obsOp     = body.obsOp     || '';

  // v1.8: resolver cliente canônico antes de qualquer comparação
  const clientCanonical = (!isSuspectClientTerm(client) ? resolveCanonicalClientName(client) : '') || client;

  // Helper: aplica campos na linha i+1 sem sobrescrever ID existente
  const applyRow = (rowIdx, preserveId) => {
    const existingId = String(data[rowIdx][idxId] || '').trim();
    // NUNCA sobrescreve um ID já existente com um diferente — protege contra corrompimento
    const finalId = preserveId ? existingId : (machineId || existingId);
    sheet.getRange(rowIdx + 1, idxId  + 1).setValue(finalId);
    if (clientCanonical) sheet.getRange(rowIdx + 1, idxCli + 1).setValue(clientCanonical);
    if (branch)   sheet.getRange(rowIdx + 1, idxFil + 1).setValue(branch);
    if (brand)    sheet.getRange(rowIdx + 1, idxMar + 1).setValue(brand);
    if (model)    sheet.getRange(rowIdx + 1, idxMod + 1).setValue(model);
    if (serial)   sheet.getRange(rowIdx + 1, idxSer + 1).setValue(serial);
    if (year)     sheet.getRange(rowIdx + 1, idxAno + 1).setValue(year);
    if (tag)      sheet.getRange(rowIdx + 1, idxTag + 1).setValue(tag);
    if (location) sheet.getRange(rowIdx + 1, idxLoc + 1).setValue(location);
    if (hourTotal > 0) sheet.getRange(rowIdx + 1, idxHor + 1).setValue(hourTotal);
    if (hpw)      sheet.getRange(rowIdx + 1, idxHpw + 1).setValue(hpw);
    if (obs)      sheet.getRange(rowIdx + 1, idxObs + 1).setValue(obs);
    // v1.9: gravar potência, tipo de equipamento e obs operacionais se disponíveis
    if (potencia  && idxPot    >= 0) sheet.getRange(rowIdx + 1, idxPot    + 1).setValue(potencia);
    if (tipoEquip && idxTipoEq >= 0) sheet.getRange(rowIdx + 1, idxTipoEq + 1).setValue(tipoEquip);
    if (obsOp     && idxObsOp  >= 0) sheet.getRange(rowIdx + 1, idxObsOp  + 1).setValue(obsOp);
    sheet.getRange(rowIdx + 1, idxUpd + 1).setValue(now);
    return finalId;
  };

  // ── Passagem 1: match por Machine_ID exato ────────────────────────────
  if (machineId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxId] || '').trim() === machineId) {
        const usedId = applyRow(i, false);
        return { status: 'ok', action: 'updated', matchedBy: 'id', machineId: usedId };
      }
    }
  }

  // ── Passagem 2: match robusto por marca+modelo+série (v1.8) ──────────
  // Requer que série (ou tag), marca e modelo coincidam — evita falso positivo
  for (let i = 1; i < data.length; i++) {
    const rowAtivo = String(data[i][idxAtivo] || 'SIM').toUpperCase();
    if (rowAtivo === 'NÃO') continue;
    const rowBrand  = String(data[i][idxMar] || '');
    const rowModel  = String(data[i][idxMod] || '');
    const rowSerial = String(data[i][idxSer] || '');
    const rowTag    = String(data[i][idxTag] || '');
    if (sameMachine(rowBrand, rowModel, rowSerial, rowTag, brand, model, serial, tag)) {
      Logger.log('ensureMachineFromVisit: match marca+modelo+série linha=' + i);
      const usedId = applyRow(i, true); // preserva o ID existente
      return { status: 'ok', action: 'updated', matchedBy: 'brand_model_serial', machineId: usedId };
    }
  }

  // ── Passagem 3: fallback por série única + validação de marca/modelo ──
  // Só aceita se a série for única na planilha E marca/modelo forem compatíveis
  if (serial) {
    const serialCompact = compactKey(serial);
    let uniqueMatch = null;
    let matchCount = 0;
    for (let j = 1; j < data.length; j++) {
      const rowAtivo = String(data[j][idxAtivo] || 'SIM').toUpperCase();
      if (rowAtivo === 'NÃO') continue;
      if (compactKey(String(data[j][idxSer] || '')) === serialCompact) {
        uniqueMatch = j; matchCount++;
        if (matchCount > 1) break;
      }
    }
    if (matchCount === 1 && uniqueMatch !== null) {
      const j = uniqueMatch;
      const rowBrand = String(data[j][idxMar] || '');
      const rowModel = String(data[j][idxMod] || '');
      const rowCli   = String(data[j][idxCli] || '').trim();
      // v1.8: validar marca+modelo antes de aceitar o fallback
      const brandCompatible = !compactKey(rowBrand) || !compactKey(brand) || compactKey(rowBrand) === compactKey(brand);
      const modelCompatible = !compactKey(rowModel) || !compactKey(model) || compactKey(rowModel) === compactKey(model);
      if (!brandCompatible || !modelCompatible) {
        Logger.log('ensureMachineFromVisit: serial_unique REJEITADO por marca/modelo incompatível. Row:' + rowBrand + '/' + rowModel + ' vs ' + brand + '/' + model);
        // Não aplica — vai inserir nova linha abaixo
      } else {
        Logger.log('ensureMachineFromVisit: serial_unique match linha=' + j + ' cliente="' + rowCli + '" recebido="' + client + '"');
        const usedId = applyRow(j, true); // SEMPRE preserva o ID existente
        return { status: 'ok', action: 'updated', matchedBy: 'serial_unique', machineId: usedId };
      }
    }
  }

  // ── Passagem 4: inserir nova linha ────────────────────────────────────
  const finalId = machineId || machineKey(clientCanonical || client, brand, model, serial) || ('EQ-' + new Date().getTime());
  const newRow = [
    finalId,
    clientCanonical || client, branch, brand, model, serial, year, tag, location,
    hourTotal, hpw, obs,
    'SIM', 'PRODUCAO', now, '', '', now
  ];
  // v1.9: append optional columns only if headers exist
  if (idxPot >= 0 || idxTipoEq >= 0 || idxObsOp >= 0) {
    // pad array to cover new column positions
    while (newRow.length < HEADERS.MAQUINAS.length) newRow.push('');
    if (idxPot    >= 0) newRow[idxPot]    = potencia  || '';
    if (idxTipoEq >= 0) newRow[idxTipoEq] = tipoEquip || '';
    if (idxObsOp  >= 0) newRow[idxObsOp]  = obsOp     || '';
  }
  sheet.appendRow(newRow);
  return { status: 'ok', action: 'inserted', machineId: finalId };
}

function updateMachineHorímetro(machineId, client, serial, tag, hourTotal, visitDate) {
  const sheet = getOrCreateSheet('MAQUINAS', HEADERS.MAQUINAS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId    = headers.indexOf('ID');
  const idxMar   = headers.indexOf('Marca');
  const idxMod   = headers.indexOf('Modelo');
  const idxSerie = headers.indexOf('Série');
  const idxTag   = headers.indexOf('TAG');
  const idxHor   = headers.indexOf('Hor.Total');
  const idxUpd   = headers.indexOf('Atualizado');

  for (let i = 1; i < data.length; i++) {
    const rowId    = String(data[i][idxId]    || '').trim();
    const rowBrand = String(data[i][idxMar]   || '');
    const rowModel = String(data[i][idxMod]   || '');
    const rowSer   = String(data[i][idxSerie] || '');
    const rowTag   = String(data[i][idxTag]   || '');
    // v1.8: match por ID exato OU por marca+modelo+série (robusto)
    const match = (machineId && rowId === String(machineId).trim()) ||
                  sameMachine(rowBrand, rowModel, rowSer, rowTag, '', '', serial, tag);
    if (match) {
      if (hourTotal > 0) sheet.getRange(i+1, idxHor+1).setValue(hourTotal);
      sheet.getRange(i+1, idxUpd+1).setValue(new Date().toISOString());
      return;
    }
  }
  Logger.log('updateMachineHorímetro: nenhuma linha encontrada para machineId=' + machineId + ' serial=' + serial);
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
    obs:       row['Observações']  || '',
    // v1.9:
    power:     row['Potência']     || '',
    type:      row['Tipo_Equip']   || '',
    obsOp:     row['Obs_Op']       || ''
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

// ═══════════════════════════════════════════════════════════════════════════
// GAS-6 — GET PART APPLICATIONS
// ═══════════════════════════════════════════════════════════════════════════
function getPartApplications(query) {
  if (!query) return { status: 'error', error: 'Parâmetro q obrigatório' };

  const norm = ref => String(ref || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  const normQuery = norm(query);
  if (!normQuery) return { status: 'error', error: 'Query inválida' };

  ensureSheetHeaders('PARTS_MASTER',      HEADERS.PARTS_MASTER);
  ensureSheetHeaders('PART_SIMILARITIES', HEADERS.PART_SIMILARITIES);
  ensureSheetHeaders('MACHINE_PARTS',     HEADERS.MACHINE_PARTS);
  ensureSheetHeaders('MODELOS',           HEADERS.MODELOS);
  ensureSheetHeaders('MAQUINAS',          HEADERS.MAQUINAS);

  const parts        = getSheetDataActive('PARTS_MASTER');
  const sims         = getSheetData('PART_SIMILARITIES').filter(r =>
    String(r['Ativo'] || 'SIM').toUpperCase() !== 'NÃO'
  );
  const machineParts = getSheetData('MACHINE_PARTS');
  const models       = getSheetDataActive('MODELOS');
  const machines     = getSheetDataActive('MAQUINAS');

  // Índice de similares por Part_ID+Model_ID
  const simMap = {};
  sims.forEach(s => {
    const key = (s['Model_ID'] || '') + '::' + (s['Part_ID'] || '');
    if (!simMap[key]) simMap[key] = [];
    simMap[key].push({
      ref:   s['Ref_Similar']   || '',
      brand: s['Brand_Similar'] || '',
      notes: s['Obs']           || ''
    });
  });

  const modelById   = {};
  models.forEach(m => { modelById[String(m['ID'] || '')] = m; });
  const machineById = {};
  machines.forEach(m => { machineById[String(m['ID'] || '')] = m; });

  const catalogMatches = [];
  const equivalentRefs = new Set();
  const matchedModelIds = new Set();

  parts.forEach(p => {
    const oemNorm  = norm(p['OEM_Ref'] || '');
    const nameNorm = norm(p['Name']    || '');
    const slotNorm = norm(p['Slot']    || '');
    const simsForPart = simMap[(p['Model_ID'] || '') + '::' + (p['Part_ID'] || '')] || [];
    const simRefs = simsForPart.map(s => norm(s.ref)).filter(Boolean);

    const hit = oemNorm === normQuery
      || nameNorm.includes(normQuery)
      || slotNorm === normQuery
      || simRefs.includes(normQuery);

    if (hit) {
      const model = modelById[String(p['Model_ID'] || '')] || {};
      matchedModelIds.add(String(p['Model_ID'] || ''));
      equivalentRefs.add(p['OEM_Ref'] || '');
      simsForPart.forEach(s => equivalentRefs.add(s.ref));

      catalogMatches.push({
        modelId:         p['Model_ID']        || '',
        brand:           model['Marca']        || '',
        model:           model['Modelo']       || '',
        power:           model['Potência']     || '',
        partId:          p['Part_ID']          || '',
        partName:        p['Name']             || '',
        oemRef:          p['OEM_Ref']          || '',
        matchedRef:      query,
        matchType:       oemNorm === normQuery ? 'oem' : simRefs.includes(normQuery) ? 'similar' : 'name_slot',
        slot:            p['Slot']             || '',
        partBrand:       p['Part_Brand']       || '',
        supplierPrimary: p['Supplier_Primary'] || '',
        subName:         p['Sub_Name']         || '',
        intervalH:       p['Interval_H']       || '',
        similarities:    simsForPart
      });
    }
  });

  // Máquinas via MACHINE_PARTS
  const machineMatches = [];
  machineParts.forEach(mp => {
    const mpRefNorm  = norm(mp['Ref'] || '');
    const mpPrevNorm = norm(mp['Ref_Anterior'] || '');
    if (mpRefNorm === normQuery || mpPrevNorm === normQuery) {
      const machine = machineById[String(mp['Machine_ID'] || '')] || {};
      if (!machine['ID']) return;
      machineMatches.push({
        machineId:  mp['Machine_ID'] || '',
        client:     machine['Cliente'] || '',
        brand:      machine['Marca']   || '',
        model:      machine['Modelo']  || '',
        serial:     machine['Série']   || '',
        tag:        machine['TAG']     || '',
        hourTotal:  machine['Hor.Total'] || 0,
        partId:     mp['Part_ID']      || '',
        partName:   mp['Part_Name']    || '',
        ref:        mp['Ref']          || '',
        matchedRef: query,
        source:     'machine_parts'
      });
    }
  });

  // Máquinas via modelo do catálogo (não cobertas por MACHINE_PARTS)
  const mpMachineIds = new Set(machineMatches.map(m => m.machineId));
  machines.forEach(machine => {
    const mkBrand = norm(machine['Marca']  || '');
    const mkModel = norm(machine['Modelo'] || '');
    const match = catalogMatches.find(c =>
      norm(c.brand) === mkBrand && norm(c.model) === mkModel
    );
    if (match && !mpMachineIds.has(String(machine['ID'] || ''))) {
      machineMatches.push({
        machineId:  String(machine['ID'] || ''),
        client:     machine['Cliente']   || '',
        brand:      machine['Marca']     || '',
        model:      machine['Modelo']    || '',
        serial:     machine['Série']     || '',
        tag:        machine['TAG']       || '',
        hourTotal:  machine['Hor.Total'] || 0,
        partId:     match.partId,
        partName:   match.partName,
        ref:        match.oemRef,
        matchedRef: query,
        source:     'catalog_model'
      });
    }
  });

  return {
    status:        'ok',
    query,
    normalizedRef: normQuery,
    catalogMatches,
    machineMatches,
    equivalentRefs: [...equivalentRefs].filter(Boolean),
    warnings: (catalogMatches.length === 0 && machineMatches.length === 0)
      ? ['Nenhuma correspondência encontrada para "' + query + '"'] : []
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// GAS-7 — GET CATALOG AUDIT
// ═══════════════════════════════════════════════════════════════════════════
function getCatalogAudit() {
  ensureSheetHeaders('PARTS_MASTER',      HEADERS.PARTS_MASTER);
  ensureSheetHeaders('PART_SIMILARITIES', HEADERS.PART_SIMILARITIES);
  ensureSheetHeaders('MACHINE_PARTS',     HEADERS.MACHINE_PARTS);
  ensureSheetHeaders('MODELOS',           HEADERS.MODELOS);

  const parts        = getSheetDataActive('PARTS_MASTER');
  const sims         = getSheetData('PART_SIMILARITIES');
  const models       = getSheetDataActive('MODELOS');
  const machineParts = getSheetData('MACHINE_PARTS');
  const norm = ref => String(ref || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

  const modelIds = new Set(models.map(m => String(m['ID'] || '').trim()).filter(Boolean));
  const activePartKeys = new Set(parts.map(p => (p['Model_ID'] || '') + '::' + (p['Part_ID'] || '')));

  // Refs duplicadas dentro do mesmo modelo (erro real)
  const refsByModel = {};
  parts.forEach(p => {
    const n = norm(p['OEM_Ref'] || '');
    if (!n) return;
    const key = (p['Model_ID'] || '') + '::' + n;
    refsByModel[key] = (refsByModel[key] || 0) + 1;
  });
  const duplicateRefsInsideSameModel = Object.values(refsByModel).filter(c => c > 1).length;

  // Refs compartilhadas entre modelos (informativo)
  const refsByNorm = {};
  parts.forEach(p => {
    const n = norm(p['OEM_Ref'] || '');
    if (!n) return;
    if (!refsByNorm[n]) refsByNorm[n] = new Set();
    refsByNorm[n].add(p['Model_ID'] || '');
  });
  const sharedRefsAcrossModels = Object.values(refsByNorm).filter(s => s.size > 1).length;

  // Part_IDs duplicados no mesmo modelo (erro real)
  const partIdByModel = {};
  parts.forEach(p => {
    const key = (p['Model_ID'] || '') + '::' + (p['Part_ID'] || '');
    partIdByModel[key] = (partIdByModel[key] || 0) + 1;
  });
  const duplicatePartIdSameModel = Object.values(partIdByModel).filter(c => c > 1).length;

  // Índice de refs do catálogo para verificar peças de máquinas
  const catalogOemRefs = new Set(parts.map(p => norm(p['OEM_Ref'] || '')).filter(Boolean));
  const simRefs        = new Set(sims.map(s => norm(s['Ref_Similar'] || '')).filter(Boolean));
  const allCatalogRefs = new Set([...catalogOemRefs, ...simRefs]);

  const issues = {
    partsWithoutRef:              parts.filter(p => !p['OEM_Ref']).length,
    partsWithoutModel:            parts.filter(p => p['Model_ID'] && !modelIds.has(String(p['Model_ID']))).length,
    orphanSimilarities:           sims.filter(s => !activePartKeys.has((s['Model_ID'] || '') + '::' + (s['Part_ID'] || ''))).length,
    duplicateRefsInsideSameModel,
    duplicatePartIdSameModel,
    modelsWithoutParts:           models.filter(m => !parts.find(p => String(p['Model_ID'] || '') === String(m['ID'] || ''))).length,
    machinePartsNotInCatalog:     machineParts.filter(mp => {
      const r = norm(mp['Ref'] || '');
      return r && !allCatalogRefs.has(r);
    }).length,
    sharedRefsAcrossModels
  };

  const severity = {
    partsWithoutRef:              issues.partsWithoutRef > 0 ? 'warning' : 'ok',
    partsWithoutModel:            issues.partsWithoutModel > 0 ? 'error' : 'ok',
    orphanSimilarities:           issues.orphanSimilarities > 0 ? 'warning' : 'ok',
    duplicateRefsInsideSameModel: issues.duplicateRefsInsideSameModel > 0 ? 'error' : 'ok',
    duplicatePartIdSameModel:     issues.duplicatePartIdSameModel > 0 ? 'error' : 'ok',
    modelsWithoutParts:           issues.modelsWithoutParts > 0 ? 'warning' : 'ok',
    machinePartsNotInCatalog:     issues.machinePartsNotInCatalog > 0 ? 'warning' : 'ok',
    sharedRefsAcrossModels:       'info'
  };

  return { status: 'ok', ts: new Date().toISOString(), issues, severity };
}

// ═══════════════════════════════════════════════════════════════════════════
// INICIALIZAÇÃO DO BANCO DE DADOS
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
