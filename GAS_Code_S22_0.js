// GAS S18.0 — Persistência completa: _garantirColunas, FIELD_DRAFTS, HISTORICO_CALIBRACOES,
//   HISTORICO_LACRES, registrarLacre, reabrirOS, fix atualizarDadosAdministrativos,
//   fix registrarCalibracao, fix avancarFaseOS (reverter itens fase 1), draft GAS
// =============================================================================
// ENGETAP — TapControl Field + TapControl Manager
// GAS_Code_S18.0.js
// Correção S12: cabeçalho atualizado, ping/healthCheck adicionado,
//   _contarEquipamentosDaOS localiza colunas pelo cabeçalho,
//   getEquipamentosByOS retorna lista vazia (nunca erro) quando não há equips,
//   salvarEquipamento não bloqueia limpeza de id_os com string vazia.
// =============================================================================
//
// INSTRUÇÕES DE DEPLOY:
// 1. Abra a planilha Engetap_TapControl_DB no Google Sheets
// 2. Extensões > Apps Script > apague o código padrão > cole este arquivo
// 3. Salve (Ctrl+S). Nomeie o projeto: "Engetap API v4"
// 4. Execute setupPlanilha() UMA VEZ para criar abas e dados de teste
// 5. Implantar > Nova implantação > App da Web
//    - Executar como: Você
//    - Quem tem acesso: Qualquer pessoa
// 6. Copie a URL gerada e use nos HTMLs (constante GAS_URL)
//
// =============================================================================
// ESTRUTURA DA PLANILHA
// =============================================================================
//
// TRANSACIONAIS:
//   CLIENTES       : id_cliente | nome | cnpj | endereco | contato | observacoes | ativo
//   OS             : id_os | id_cliente | cliente | numero_os | descricao | data_abertura | status | fase_atual
//   EQUIPAMENTOS   : id_equipamento | id_cliente | id_os | tipo | tag | localizacao | pmta | campo_livre | cadastrado_por | ativo
//   LEVANTAMENTOS  : id_levantamento | id_os | id_cliente | id_inspetor | nome_inspetor | fase | timestamp_envio | total_itens | status_geral
//   ITENS          : id_item | id_levantamento | id_os | id_cliente | id_equipamento | tag_equipamento |
//                    tipo_acessorio | id_item_catalogo | descricao_curta | quantidade |
//                    escala_inicio | escala_fim | escala_unidade | id_kit | componentes_kit |
//                    ajuste_kit | requer_calibracao | requer_lacre | instalado_imediatamente |
//                    status | fornecedor | valor_cotado | prazo_entrega | status_aprovacao_cliente |
//                    data_calibracao | num_certificado | num_lacre | validade_calibracao | lab_calibracao |
//                    data_instalacao | data_entrega | observacao_fechamento |
//                    timestamp_criacao | timestamp_atualizacao | origem
//   HISTORICO_STATUS: id_historico | id_item | status_anterior | status_novo | responsavel | observacao | timestamp
//
// CATÁLOGOS:
//   CAT_INSPETORES : id | nome | pin_hash | ativo
//   CAT_MANOMETROS : id | tipo | fluido | caixa_mm | rosca | tipo_rosca | entrada | escala_padrao | marca | modelo | desc_curta | requer_calibracao | ativo
//   CAT_VALVULAS   : id | bitola_ent | bitola_sai | fluido | pa_kgf | material | marca | modelo | desc_curta | requer_calibracao | ativo
//   CAT_PURGADORES : id | tipo | bitola | fluido | desc_curta | ativo
//   CAT_DCBI       : id | bitola | componentes | desc_curta | ativo
//   CAT_CONEXOES   : id | tipo | bitola | material | desc_curta | ativo
//   CAT_KITS       : id | nome | bitola | tipo_kit | componentes | inclui_rs | ativo
//
// =============================================================================

// =============================================================================
// CONFIGURAÇÃO
// =============================================================================

var DRIVE_LINK_PUBLICO = false; // false = acesso restrito (recomendado para contexto industrial)

// =============================================================================
// SETUP
// =============================================================================

function setupPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Transacionais
  _criarAba(ss, 'CLIENTES',        ['id_cliente','nome','cnpj','endereco','contato','observacoes','ativo']);
  _criarAba(ss, 'OS',              ['id_os','id_cliente','cliente','numero_os','descricao','data_abertura','status','fase_atual','tecnicos_vinculados']);
  _criarAba(ss, 'EQUIPAMENTOS',    ['id_equipamento','id_cliente','id_os','tipo','tag','localizacao','pmta','campo_livre','cadastrado_por','ativo']);
  _criarAba(ss, 'LEVANTAMENTOS',   ['id_levantamento','id_os','id_cliente','id_inspetor','nome_inspetor','fase','timestamp_envio','total_itens','status_geral','drive_folder_id','drive_folder_url','quantidade_fotos_equip','quantidade_fotos_ponto','fotos_equip_json','fotos_ponto_json']);
  _criarAba(ss, 'ITENS', [
    'id_item','id_levantamento','id_os','id_cliente','id_equipamento','tag_equipamento',
    'tipo_acessorio','id_item_catalogo','descricao_curta','quantidade',
    'escala_inicio','escala_fim','escala_unidade','id_kit','componentes_kit',
    'ajuste_kit','requer_calibracao','requer_lacre','instalado_imediatamente',
    'status','fornecedor','valor_cotado','prazo_entrega','status_aprovacao_cliente',
    'data_calibracao','num_certificado','num_lacre','validade_calibracao','lab_calibracao',
    'data_instalacao','data_entrega','observacao_fechamento',
    'timestamp_criacao','timestamp_atualizacao','origem','foto_equipamento_b64','foto_ponto_instalacao_b64'
  ]);
  _criarAba(ss, 'HISTORICO_STATUS', ['id_historico','id_item','status_anterior','status_novo','responsavel','observacao','timestamp']);

  // S18 — Novas abas de histórico
  _criarAba(ss, 'HISTORICO_CALIBRACOES', [
    'id_calibracao','id_item','id_os','id_equipamento','tag_equipamento',
    'tipo_calibracao','data_calibracao','validade_calibracao',
    'num_certificado','num_lacre','num_lacre2',
    'lab_calibracao','responsavel','observacao',
    'timestamp_registro','origem'
  ]);
  _criarAba(ss, 'HISTORICO_LACRES', [
    'id_lacre','id_item','id_os','id_equipamento','tag_equipamento',
    'num_lacre','data_lacre','responsavel_lacre','observacao_lacre',
    'timestamp_registro','origem'
  ]);
  _criarAba(ss, 'FIELD_DRAFTS', [
    'id_draft','id_os','numero_os','id_inspetor','nome_inspetor',
    'fase','timestamp_atualizacao','total_itens','status_draft','payload_json'
  ]);

  // Sprint 6 — Histórico e Log Operacional
  _criarAba(ss, 'HISTORICO_OS', [
    'id_historico_os','id_os','numero_os','id_cliente','cliente',
    'data_abertura','data_encerramento','status_final_os','fase_final',
    'encerrada_por','origem_encerramento','motivo_encerramento',
    'teve_faturamento','motivo_sem_faturamento',
    'valor_total_estimado','valor_total_aprovado',
    'total_itens','total_itens_concluidos','total_itens_reprovados','total_itens_pendentes',
    'observacao_encerramento','timestamp_registro'
  ]);
  _criarAba(ss, 'LOG_OPERACIONAL', [
    'id_evento','timestamp','tipo_evento',
    'id_os','numero_os','id_cliente','cliente',
    'id_item','id_equipamento','tag_equipamento',
    'produto','tipo_acessorio',
    'status_anterior','status_novo','acao_realizada',
    'responsavel_id','responsavel_nome','responsavel_tipo',
    'origem','fornecedor','valor','observacao','payload_json'
  ]);

  // Catálogos
  _criarAba(ss, 'CAT_INSPETORES',  ['id','nome','pin_hash','ativo']);
  _criarAba(ss, 'CAT_MANOMETROS',  ['id','tipo','fluido','caixa_mm','rosca','tipo_rosca','entrada','escala_padrao','marca','modelo','desc_curta','requer_calibracao','ativo']);
  _criarAba(ss, 'CAT_VALVULAS',    ['id','bitola_ent','bitola_sai','fluido','pa_kgf','material','marca','modelo','desc_curta','requer_calibracao','ativo']);
  _criarAba(ss, 'CAT_PURGADORES',  ['id','tipo','bitola','fluido','desc_curta','ativo']);

  // S18 — Garantir novas colunas em abas existentes
  _garantirColunas(ss, 'ITENS', [
    'conexoes_montagem_json','itens_adicionais_json',
    'origem_item_manual','dados_item_manual_json','payload_item_json',
    'tipo_calibracao','num_lacre2','observacao_cotacao',
    'foto_equipamento_b64','foto_ponto_instalacao_b64'
  ]);
  _garantirColunas(ss, 'LEVANTAMENTOS', [
    'drive_folder_id','drive_folder_url',
    'quantidade_fotos_equip','quantidade_fotos_ponto',
    'fotos_equip_json','fotos_ponto_json'
  ]);
  _garantirColunas(ss, 'CAT_VALVULAS', [
    'tipo_valvula','tipo_entrada','tipo_saida','tipo_rosca',
    'pa_alivio','pa_alivio_unidade','pa_vacuo','pa_vacuo_unidade',
    'pmta','pmta_unidade','payload_json'
  ]);
  _garantirColunas(ss, 'CAT_PURGADORES', ['pn','marca','modelo','payload_json']);
  _criarAba(ss, 'CAT_DCBI',        ['id','bitola','componentes','desc_curta','ativo']);
  _criarAba(ss, 'CAT_CONEXOES',    ['id','tipo','bitola','material','desc_curta','ativo']);
  _criarAba(ss, 'CAT_KITS',        ['id','nome','bitola','tipo_kit','componentes','inclui_rs','ativo']);

  _popularDadosDeTeste(ss);
  Logger.log('setupPlanilha() concluído.');
}

function _criarAba(ss, nome, cabecalhos) {
  var aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos])
       .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
    Logger.log('Aba criada: ' + nome);
  } else {
    Logger.log('Aba já existe: ' + nome);
  }
}

// S18 — adiciona colunas que ainda não existem ao final do cabeçalho.
// Nunca apaga, reordena nem sobrescreve colunas ou dados existentes.
function _garantirColunas(ss, nomeAba, colunasObrigatorias) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) { Logger.log('_garantirColunas: aba não encontrada — ' + nomeAba); return; }
  var cabAtual = aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), 1)).getValues()[0].map(String);
  var ausentes = colunasObrigatorias.filter(function(c) { return cabAtual.indexOf(c) < 0; });
  if (ausentes.length === 0) { Logger.log('_garantirColunas: ' + nomeAba + ' já tem todas as colunas.'); return; }
  var colInicio = aba.getLastColumn() + 1;
  aba.getRange(1, colInicio, 1, ausentes.length).setValues([ausentes])
     .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  Logger.log('_garantirColunas: ' + nomeAba + ' — adicionadas: ' + ausentes.join(', '));
}


function norm(v) {
  return String(v || '').trim();
}

function getOrCreateDrivePath(partes) {
  var pasta = DriveApp.getRootFolder();
  partes.forEach(function(nome) {
    var it = pasta.getFoldersByName(nome);
    pasta = it.hasNext() ? it.next() : pasta.createFolder(nome);
  });
  return pasta;
}

function dataUrlToBlob(dataUrl, nome) {
  var arr = String(dataUrl || '').split(',');
  var mimeMatch = arr[0].match(/:(.*?);/);
  var mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
  var ext = mime.split('/')[1] === 'png' ? '.png' : '.jpg';
  var bytes = Utilities.base64Decode(arr[1] || '');
  return Utilities.newBlob(bytes, mime, nome + ext);
}

function sanitizeFileName(s) {
  return String(s || '').replace(/[^\w\-_.]/g, '_').substring(0, 80);
}

function _setColIfExists(aba, rowNum, cab, col, val) {
  var idx = cab.indexOf(col);
  if (idx >= 0) aba.getRange(rowNum, idx + 1).setValue(val);
}

function logErro(contexto, msg, detalhe) {
  try {
    Logger.log('[ERRO] ' + contexto + ': ' + msg + (detalhe ? ' — ' + detalhe : ''));
    registrarEventoOperacional({
      tipo_evento: 'ERRO_SISTEMA',
      acao_realizada: contexto,
      origem: 'GAS',
      observacao: msg || '',
      payload_json: detalhe || ''
    });
  } catch(e) {
    Logger.log('logErro falhou: ' + e.message);
  }
}

function _popularDadosDeTeste(ss) {
  // CAT_INSPETORES
  var abaInsp = ss.getSheetByName('CAT_INSPETORES');
  if (abaInsp.getLastRow() <= 1) {
    abaInsp.appendRow(['INS-001', 'João Silva',     _hashPin('1234'), 'true']);
    abaInsp.appendRow(['INS-002', 'Maria Oliveira', _hashPin('5678'), 'true']);
  }

  // CLIENTES
  var abaCli = ss.getSheetByName('CLIENTES');
  if (abaCli.getLastRow() <= 1) {
    abaCli.appendRow(['CLI-001', 'Gerdau – Divinópolis',     '26.668.182/0001-85', 'Rod. BR-040 km 5, Divinópolis MG', 'Carlos Melo · (37) 99999-0001', '', 'true']);
    abaCli.appendRow(['CLI-002', 'Farmax Ind. Farmacêutica', '12.345.678/0001-90', 'Av. Industrial, 200, Divinópolis MG', 'Ana Paula · (37) 99999-0002', '', 'true']);
  }

  // OS
  var abaOS = ss.getSheetByName('OS');
  if (abaOS.getLastRow() <= 1) {
    abaOS.appendRow(['OS-2026-001', 'CLI-001', 'Gerdau – Divinópolis',     'OS-2026-001', 'Inspeção NR-13 anual', '2026-05-01', 'ativa', '1', '["INS-001"]']);
    abaOS.appendRow(['OS-2026-002', 'CLI-002', 'Farmax Ind. Farmacêutica', 'OS-2026-002', 'Calibração e inspeção',  '2026-05-03', 'ativa', '1', '["INS-002"]']);
  }

  // EQUIPAMENTOS
  var abaEq = ss.getSheetByName('EQUIPAMENTOS');
  if (abaEq.getLastRow() <= 1) {
    abaEq.appendRow(['EQ-001', 'CLI-001', 'OS-2026-001', 'vaso',      'VP-01', 'Sala de utilidades',  '10', 'Compressor de ar centrífugo', 'adm', 'true']);
    abaEq.appendRow(['EQ-002', 'CLI-001', 'OS-2026-001', 'caldeira',  'CL-01', 'Casa de caldeiras',   '8',  '',                            'adm', 'true']);
    abaEq.appendRow(['EQ-003', 'CLI-002', 'OS-2026-002', 'vaso',      'VP-01', 'Área de produção',    '12', 'Reator principal',             'adm', 'true']);
    abaEq.appendRow(['EQ-004', 'CLI-002', 'OS-2026-002', 'tanque',    'TQ-01', 'Pátio externo',       '6',  'Tanque de expansão',           'adm', 'true']);
  }

  // CAT_MANOMETROS
  var abaMano = ss.getSheetByName('CAT_MANOMETROS');
  if (abaMano.getLastRow() <= 1) {
    abaMano.appendRow(['MAN-001','Manômetro','Glicerinado','63','1/4"','NPT','Reta',  '0-14 kgf/cm²','Pressure','','Manômetro Glicerinado 63mm · 1/4" NPT · Reta · 0-14 kgf/cm² · Pressure','false','true']);
    abaMano.appendRow(['MAN-002','Manômetro','Seco',       '100','1/4"','NPT','Reta', '0-10 kgf/cm²','Wika',    '','Manômetro Seco 100mm · 1/4" NPT · Reta · 0-10 kgf/cm² · Wika',          'false','true']);
    abaMano.appendRow(['MAN-003','Manômetro','Glicerinado','63','1/2"','NPT','Reta',  '0-21 kgf/cm²','Pressure','','Manômetro Glicerinado 63mm · 1/2" NPT · Reta · 0-21 kgf/cm² · Pressure','false','true']);
    abaMano.appendRow(['MAN-004','Vacuômetro','Glicerinado','63','1/4"','NPT','Reta', '-1-0 kgf/cm²','Wika',   '','Vacuômetro Glicerinado 63mm · 1/4" NPT · Reta · -1-0 kgf/cm² · Wika',   'false','true']);
    abaMano.appendRow(['MAN-005','Manômetro','Seco',       '100','1/2"','NPT','Angular','0-14 kgf/cm²','Genebre','','Manômetro Seco 100mm · 1/2" NPT · Angular · 0-14 kgf/cm² · Genebre','true','true']);
  }

  // CAT_VALVULAS
  var abaVal = ss.getSheetByName('CAT_VALVULAS');
  if (abaVal.getLastRow() <= 1) {
    abaVal.appendRow(['VAL-001','1/2"','1/2"','Ar','10','Aço inox','Leser','SIL526','Válvula 1/2" Ar PA=10kgf · Aço inox · Leser SIL526','true','true']);
    abaVal.appendRow(['VAL-002','3/4"','3/4"','Vapor','8', 'Aço carbono','Farago','FK-300','Válvula 3/4" Vapor PA=8kgf · Aço carbono · Farago FK-300','true','true']);
  }

  // CAT_PURGADORES
  var abaPurg = ss.getSheetByName('CAT_PURGADORES');
  if (abaPurg.getLastRow() <= 1) {
    abaPurg.appendRow(['PUR-001','Termodinâmico','1/2"','Vapor','Purgador Termodinâmico 1/2" · Vapor','true']);
    abaPurg.appendRow(['PUR-002','Boia',         '3/4"','Vapor','Purgador Boia 3/4" · Vapor',         'true']);
  }

  // CAT_DCBI
  var abaDCBI = ss.getSheetByName('CAT_DCBI');
  if (abaDCBI.getLastRow() <= 1) {
    abaDCBI.appendRow(['DCBI-001','1/4"', JSON.stringify([
      {descricao:'Válvula tripartida 1/4"',qtd:1},
      {descricao:'Niple 1/4" NPT',qtd:2},
      {descricao:'Joelho 1/4" NPT',qtd:1},
      {descricao:'Lacre',qtd:1},
      {descricao:'Plaquinha PVC advertência',qtd:1}
    ]), 'DCBI 1/4" completo', 'true']);
    abaDCBI.appendRow(['DCBI-002','1/2"', JSON.stringify([
      {descricao:'Válvula tripartida 1/2"',qtd:1},
      {descricao:'Niple 1/2" NPT',qtd:2},
      {descricao:'Joelho 1/2" NPT',qtd:1},
      {descricao:'Lacre',qtd:1},
      {descricao:'Plaquinha PVC advertência',qtd:1}
    ]), 'DCBI 1/2" completo', 'true']);
  }

  // CAT_KITS
  var abaKits = ss.getSheetByName('CAT_KITS');
  if (abaKits.getLastRow() <= 1) {
    abaKits.appendRow(['KIT-001','Kit instalação simples 1/4"','1/4"','simples',
      JSON.stringify([{descricao:'Niple 1/4" NPT',qtd:1},{descricao:'Joelho 1/4" NPT',qtd:1}]),
      'false','true']);
    abaKits.appendRow(['KIT-002','Kit DCBI 1/4"','1/4"','dcbi',
      JSON.stringify([{descricao:'Válvula tripartida 1/4"',qtd:1},{descricao:'Niple 1/4" NPT',qtd:2},{descricao:'Joelho 1/4" NPT',qtd:1},{descricao:'Lacre',qtd:1},{descricao:'Plaquinha PVC advertência',qtd:1}]),
      'true','true']);
    abaKits.appendRow(['KIT-003','Kit instalação simples 1/2"','1/2"','simples',
      JSON.stringify([{descricao:'Niple 1/2" NPT',qtd:1},{descricao:'Joelho 1/2" NPT',qtd:1}]),
      'false','true']);
    abaKits.appendRow(['KIT-004','Kit DCBI 1/2"','1/2"','dcbi',
      JSON.stringify([{descricao:'Válvula tripartida 1/2"',qtd:1},{descricao:'Niple 1/2" NPT',qtd:2},{descricao:'Joelho 1/2" NPT',qtd:1},{descricao:'Lacre',qtd:1},{descricao:'Plaquinha PVC advertência',qtd:1}]),
      'true','true']);
  }

  Logger.log('Dados de teste inseridos.');
}


// =============================================================================
// ROTEADOR PRINCIPAL
// =============================================================================

function doGet(e) {
  try {
    var action = e.parameter.action;
    var params = e.parameter;
    var resultado;
    switch (action) {
      case 'validarPIN':                resultado = validarPIN(params);                break;
      case 'getClientes':               resultado = getClientes(params);               break;
      case 'getOS':                     resultado = getOS(params);                     break;
      case 'getOSByCliente':            resultado = getOSByCliente(params);            break;
      case 'getEquipamentosByOS':       resultado = getEquipamentosByOS(params);       break;
      case 'getEquipamentosByCliente':  resultado = getEquipamentosByCliente(params);  break;
      case 'getEquipamentosAll':        resultado = getEquipamentosAll(params);        break;
      case 'getCatalogos':              resultado = getCatalogos(params);              break;
      case 'getCatalogosHash':          resultado = getCatalogosHash(params);          break;
      case 'getCatalogoByTipo':         resultado = getCatalogoByTipo(params);         break;
      case 'getKitsByBitola':           resultado = getKitsByBitola(params);           break;
      case 'getLevantamentosByCliente': resultado = getLevantamentosByCliente(params); break;
      case 'getLevantamentosByOS':      resultado = getLevantamentosByOS(params);      break;
      case 'getItensByLevantamento':    resultado = getItensByLevantamento(params);    break;
      case 'getFotoBase64':              resultado = getFotoBase64(params);              break;
      case 'getLevantamentosComFotos':   resultado = getLevantamentosComFotos(params);   break;
      case 'getItensByOS':              resultado = getItensByOS(params);              break;
      case 'getInspetores':             resultado = getInspetores(params);             break;
      // Sprint 6 — Histórico
      case 'getHistoricoOS':            resultado = getHistoricoOS(params);            break;
      case 'getHistoricoItens':         resultado = getHistoricoItens(params);         break;
      case 'getLogOperacional':         resultado = getLogOperacional(params);         break;
      case 'getHistorico':              resultado = getHistorico(params);              break;
      // Correção S12 — diagnóstico: ping / healthCheck
      case 'ping':
      case 'healthCheck':
        resultado = { status: 'ok', app: 'Engetap TapControl GAS', version: 'S17.0-Beta-RC1', timestamp: new Date().toISOString() };
        break;
      default:
        resultado = _erro('ACAO_INVALIDA', 'action não reconhecida: ' + action);
    }
    return _resposta(resultado);
  } catch (err) {
    return _resposta(_erro('ERRO_INTERNO', err.message));
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var resultado;
    switch (action) {
      case 'salvarCliente':                 resultado = salvarCliente(body);                 break;
      case 'salvarOS':                      resultado = salvarOS(body);                      break;
      case 'salvarEquipamento':             resultado = salvarEquipamento(body);             break;
      case 'enviarLevantamento':            resultado = enviarLevantamento(body);            break;
      case 'atualizarStatusItem':           resultado = atualizarStatusItem(body);           break;
      case 'atualizarDadosAdministrativos': resultado = atualizarDadosAdministrativos(body); break;
      case 'registrarCalibracao':           resultado = registrarCalibracao(body);           break;
      case 'registrarInstalacao':           resultado = registrarInstalacao(body);           break;
      case 'avancarFaseOS':                 resultado = avancarFaseOS(body);                 break;
      case 'salvarInspetor':                resultado = salvarInspetor(body);                break;
      case 'toggleAtivoInspetor':           resultado = toggleAtivoInspetor(body);           break;
      // Sprint 6 — Histórico
      case 'encerrarOSHistorico':           resultado = encerrarOSHistorico(body);           break;
      // Correção S10 — vínculo ativo equipamento/OS: endpoint explícito
      case 'vincularEquipamentoOS':         resultado = vincularEquipamentoOS(body);         break;
      case 'salvarItemCatalogo':            resultado = salvarItemCatalogo(body);            break;
      case 'toggleAtivoItem':               resultado = toggleAtivoItem(body);               break;
      case 'cancelarEnvioLevantamentoFieldTap': resultado = cancelarEnvioLevantamentoFieldTap(body); break;
      case 'cancelarFechamentoFase2FieldTap':   resultado = cancelarFechamentoFase2FieldTap(body);   break;
      case 'reabrirOSParaFieldTap':             resultado = reabrirOSParaFieldTap(body);             break;
      case 'reenviarOSFieldTap':                resultado = reenviarOSFieldTap(body);                break;
      // S18 — novos endpoints
      case 'reabrirOS':                         resultado = reabrirOS(body);                         break;
      case 'registrarLacre':                    resultado = registrarLacre(body);                    break;
      case 'salvarDraftFieldTap':               resultado = salvarDraftFieldTap(body);               break;
      case 'getDraftFieldTap':                  resultado = getDraftFieldTap(body);                  break;
      case 'limparDraftFieldTap':               resultado = limparDraftFieldTap(body);               break;
      default:
        resultado = _erro('ACAO_INVALIDA', 'action não reconhecida: ' + action);
    }
    return _resposta(resultado);
  } catch (err) {
    return _resposta(_erro('ERRO_INTERNO', err.message));
  }
}


// =============================================================================
// ENDPOINTS GET
// =============================================================================

function validarPIN(params) {
  if (!params.pin) return _erro('CAMPO_OBRIGATORIO', 'pin é obrigatório.');
  var hashInformado = _hashPin(params.pin);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('CAT_INSPETORES').getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][2] === hashInformado && String(dados[i][3]).toLowerCase() === 'true') {
      return { status: 'ok', valido: true, id_inspetor: dados[i][0], nome: dados[i][1] };
    }
  }
  return { status: 'ok', valido: false };
}

function getClientes(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('CLIENTES').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.ativo).toLowerCase() === 'true') lista.push(obj);
  }
  return { status: 'ok', clientes: lista };
}


function _isOSVisivelNoFieldTap(obj) {
  var fase = String(obj.fase_atual || '').toLowerCase();
  var status = String(obj.status || '').toLowerCase();
  if (_isOSEncerradaGAS(obj.status, obj.fase_atual)) return false;
  if (fase === 'aguardando_tapparts' || status === 'aguardando tapparts') return false;
  if (fase === 'campo_concluido' || status === 'fase 2 campo concluída') return false;
  var isF1 = fase === '1' || status === 'fase 1';
  var isF2 = fase === '2' || status === 'fase 2';
  return isF1 || isF2;
}

function getOS(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('OS').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  var idInspetor = params.id_inspetor ? String(params.id_inspetor) : null;
  var isFieldTap = !!idInspetor;
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (_isOSEncerradaGAS(obj.status, obj.fase_atual)) continue;
    var tecs = [];
    try { tecs = JSON.parse(obj.tecnicos_vinculados || '[]'); } catch(e) { tecs = []; }
    obj.tecnicos_vinculados = tecs;
    if (idInspetor && tecs.length > 0 && tecs.indexOf(idInspetor) < 0) continue;
    if (isFieldTap && !_isOSVisivelNoFieldTap(obj)) continue;
    obj.total_equipamentos = _contarEquipamentosDaOS(ss, obj.id_os);
    lista.push(obj);
  }
  return { status: 'ok', os: lista };
}

function getOSByCliente(params) {
  if (!params.id_cliente) return _erro('CAMPO_OBRIGATORIO', 'id_cliente é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('OS').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.id_cliente) === String(params.id_cliente)) {
      obj.total_equipamentos = _contarEquipamentosDaOS(ss, obj.id_os);
      lista.push(obj);
    }
  }
  return { status: 'ok', os: lista };
}

function getEquipamentosByOS(params) {
  // Correção S12 — retorna lista vazia com status ok; nunca retorna erro por lista vazia
  if (!params.id_os) return _erro('CAMPO_OBRIGATORIO', 'id_os é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('EQUIPAMENTOS').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.id_os) === String(params.id_os) && String(obj.ativo).toLowerCase() === 'true') {
      lista.push(obj);
    }
  }
  // Sempre retorna ok, mesmo com lista vazia — o FieldTap trata a ausência de equipamentos
  return { status: 'ok', equipamentos: lista, total: lista.length };
}

function getEquipamentosByCliente(params) {
  if (!params.id_cliente) return _erro('CAMPO_OBRIGATORIO', 'id_cliente é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('EQUIPAMENTOS').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.id_cliente) === String(params.id_cliente) && String(obj.ativo).toLowerCase() === 'true') {
      lista.push(obj);
    }
  }
  return { status: 'ok', equipamentos: lista };
}

function getEquipamentosAll(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('EQUIPAMENTOS').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.ativo).toLowerCase() === 'true') {
      lista.push(obj);
    }
  }
  return { status: 'ok', equipamentos: lista };
}

function getCatalogos(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    status: 'ok',
    manometros: _lerCatalogo(ss, 'CAT_MANOMETROS'),
    valvulas:   _lerCatalogo(ss, 'CAT_VALVULAS'),
    purgadores: _lerCatalogo(ss, 'CAT_PURGADORES'),
    dcbi:       _lerCatalogoDCBI(ss),
    conexoes:   _lerCatalogo(ss, 'CAT_CONEXOES'),
    kits:       _lerCatalogo(ss, 'CAT_KITS')
  };
}

function getCatalogosHash(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = ['CAT_MANOMETROS','CAT_VALVULAS','CAT_PURGADORES','CAT_DCBI','CAT_CONEXOES','CAT_KITS'];
  var resultado = { status: 'ok' };
  abas.forEach(function(nome) {
    var aba = ss.getSheetByName(nome);
    if (!aba) { resultado[nome] = 'vazio'; return; }
    var conteudo = JSON.stringify(aba.getDataRange().getValues());
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, conteudo);
    resultado[nome.replace('CAT_','').toLowerCase()] = bytes.map(function(b) {
      return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
    }).join('').substring(0, 8);
  });
  return resultado;
}

function getCatalogoByTipo(params) {
  if (!params.tipo) return _erro('CAMPO_OBRIGATORIO', 'tipo é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = {
    manometros: 'CAT_MANOMETROS',
    valvulas:   'CAT_VALVULAS',
    purgadores: 'CAT_PURGADORES',
    dcbi:       'CAT_DCBI',
    conexoes:   'CAT_CONEXOES',
    kits:       'CAT_KITS'
  };
  var nomeAba = mapa[params.tipo];
  if (!nomeAba) return _erro('CATALOGO_INVALIDO', 'Tipo de catálogo inválido: ' + params.tipo);
  var itens = params.tipo === 'dcbi' ? _lerCatalogoDCBI(ss) : _lerCatalogo(ss, nomeAba);
  return { status: 'ok', itens: itens };
}

function getKitsByBitola(params) {
  if (!params.bitola) return _erro('CAMPO_OBRIGATORIO', 'bitola é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var todos = _lerCatalogo(ss, 'CAT_KITS');
  // Aceita 'todos' para retornar todos os kits sem filtro de bitola
  if (String(params.bitola).trim().toLowerCase() === 'todos') {
    return { status: 'ok', kits: todos };
  }
  var filtrados = todos.filter(function(k) {
    return String(k.bitola || '').trim() === String(params.bitola).trim();
  });
  return { status: 'ok', kits: filtrados };
}

function getLevantamentosByCliente(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('LEVANTAMENTOS').getDataRange().getValues();
  var cab = dados[0];
  // [LEV_ENRICH] S17 — carregar OS uma vez e construir mapa id_os -> {cliente, numero_os}
  var osMap = _carregarMapaOS(ss);
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    obj.resumo_status = _resumoStatusDoLevantamento(ss, obj.id_levantamento);
    var infoOS = osMap[String(obj.id_os)];
    if (infoOS) {
      obj.cliente   = obj.cliente   || infoOS.cliente   || '';
      obj.numero_os = obj.numero_os || infoOS.numero_os || '';
      obj.id_cliente = obj.id_cliente || infoOS.id_cliente || '';
    }
    lista.push(obj);
  }
  lista.sort(function(a, b) { return String(b.timestamp_envio).localeCompare(String(a.timestamp_envio)); });
  return { status: 'ok', levantamentos: lista };
}

function getLevantamentosByOS(params) {
  if (!params.id_os) return _erro('CAMPO_OBRIGATORIO', 'id_os é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('LEVANTAMENTOS').getDataRange().getValues();
  var cab = dados[0];
  // [LEV_ENRICH] S17 — enriquecer com cliente/numero_os
  var osMap = _carregarMapaOS(ss);
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.id_os) === String(params.id_os)) {
      obj.resumo_status = _resumoStatusDoLevantamento(ss, obj.id_levantamento);
      var infoOS = osMap[String(obj.id_os)];
      if (infoOS) {
        obj.cliente   = obj.cliente   || infoOS.cliente   || '';
        obj.numero_os = obj.numero_os || infoOS.numero_os || '';
        obj.id_cliente = obj.id_cliente || infoOS.id_cliente || '';
      }
      lista.push(obj);
    }
  }
  return { status: 'ok', levantamentos: lista };
}


function getLevantamentosComFotos(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _garantirColunas(ss, 'LEVANTAMENTOS', [
    'drive_folder_id','drive_folder_url',
    'quantidade_fotos_equip','quantidade_fotos_ponto',
    'fotos_equip_json','fotos_ponto_json'
  ]);
  var aba = ss.getSheetByName('LEVANTAMENTOS');
  if (!aba || aba.getLastRow() < 2) return { status: 'ok', levantamentos: [] };
  var dados = aba.getDataRange().getValues();
  var cab = dados[0];
  var osMap = _carregarMapaOS(ss);
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (obj.drive_folder_id || parseInt(obj.quantidade_fotos_equip, 10) > 0 || parseInt(obj.quantidade_fotos_ponto, 10) > 0) {
      var infoOS = osMap[String(obj.id_os)];
      if (infoOS) {
        obj.cliente = obj.cliente || infoOS.cliente || '';
        obj.numero_os = obj.numero_os || infoOS.numero_os || '';
        obj.id_cliente = obj.id_cliente || infoOS.id_cliente || '';
      }
      lista.push(obj);
    }
  }
  lista.sort(function(a, b) { return String(b.timestamp_envio).localeCompare(String(a.timestamp_envio)); });
  return { status: 'ok', levantamentos: lista };
}

function getFotoBase64(params) {
  var fileId = norm(params && params.file_id);
  if (!fileId) return { status: 'erro', mensagem: 'file_id obrigatório' };
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return {
      status: 'ok',
      base64: Utilities.base64Encode(blob.getBytes()),
      mimeType: blob.getContentType() || 'image/jpeg'
    };
  } catch(e) {
    logErro('getFotoBase64', e.message, JSON.stringify({ file_id: fileId }));
    return { status: 'erro', mensagem: 'Erro ao buscar foto: ' + e.message };
  }
}

function salvarFotosDriveTapControl(body) {
  var resp = {
    drive_folder_id: '', drive_folder_url: '',
    fotos_salvas_equip: [], fotos_salvas_ponto: [],
    fotos_com_erro: [],
    quantidade_fotos_equip: 0,
    quantidade_fotos_ponto: 0
  };

  var folder;
  if (body.drive_folder_id_existente) {
    try { folder = DriveApp.getFolderById(body.drive_folder_id_existente); } catch(e) { folder = null; }
  }
  if (!folder) {
    var ym = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM');
    folder = getOrCreateDrivePath([
      'TapControl_Fotos', ym,
      '[' + sanitizeFileName(body.id_os || body.numero_os || 'SEM_OS') + ']',
      '[' + sanitizeFileName(body.id_levantamento || 'SEM_LEV') + ']'
    ]);
  }
  if (DRIVE_LINK_PUBLICO) {
    try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  }
  resp.drive_folder_id = folder.getId();
  resp.drive_folder_url = folder.getUrl();

  Object.keys(body.fotos_equipamentos || {}).forEach(function(idEquip) {
    var dataUrl = body.fotos_equipamentos[idEquip];
    if (!dataUrl || typeof dataUrl !== 'string' || dataUrl.indexOf('data:') !== 0) return;
    try {
      var nome = sanitizeFileName('equip_' + idEquip);
      var existing = folder.getFilesByName(nome + '.jpg');
      var arq = existing.hasNext() ? existing.next() : folder.createFile(dataUrlToBlob(dataUrl, nome));
      if (DRIVE_LINK_PUBLICO) { try { arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {} }
      resp.fotos_salvas_equip.push({ id_equipamento: idEquip, file_id: arq.getId(), file_url: arq.getUrl() });
    } catch(err) {
      resp.fotos_com_erro.push({ key: 'equip_' + idEquip, erro: err.message });
      logErro('salvarFotosDriveTapControl/equip', err.message, idEquip);
    }
  });

  Object.keys(body.fotos_pontos_instalacao || {}).forEach(function(tipoKey) {
    var dataUrl = body.fotos_pontos_instalacao[tipoKey];
    if (!dataUrl || typeof dataUrl !== 'string' || dataUrl.indexOf('data:') !== 0) return;
    try {
      var nome = sanitizeFileName('ponto_' + tipoKey);
      var existing = folder.getFilesByName(nome + '.jpg');
      var arq = existing.hasNext() ? existing.next() : folder.createFile(dataUrlToBlob(dataUrl, nome));
      if (DRIVE_LINK_PUBLICO) { try { arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {} }
      resp.fotos_salvas_ponto.push({ tipo_key: tipoKey, file_id: arq.getId(), file_url: arq.getUrl() });
    } catch(err) {
      resp.fotos_com_erro.push({ key: 'ponto_' + tipoKey, erro: err.message });
      logErro('salvarFotosDriveTapControl/ponto', err.message, tipoKey);
    }
  });

  resp.quantidade_fotos_equip = resp.fotos_salvas_equip.length;
  resp.quantidade_fotos_ponto = resp.fotos_salvas_ponto.length;
  return resp;
}

// [LEV_ENRICH] S17 — auxiliar reutilizável (lê OS uma vez, devolve map)
function _carregarMapaOS(ss) {
  var dadosOS = ss.getSheetByName('OS').getDataRange().getValues();
  var cabOS = dadosOS[0];
  var idxId = cabOS.indexOf('id_os');
  var idxCli = cabOS.indexOf('cliente');
  var idxNum = cabOS.indexOf('numero_os');
  var idxIdCli = cabOS.indexOf('id_cliente');
  var map = {};
  for (var j = 1; j < dadosOS.length; j++) {
    var key = String(dadosOS[j][idxId] || '');
    if (!key) continue;
    map[key] = {
      cliente: idxCli >= 0 ? String(dadosOS[j][idxCli] || '') : '',
      numero_os: idxNum >= 0 ? String(dadosOS[j][idxNum] || '') : '',
      id_cliente: idxIdCli >= 0 ? String(dadosOS[j][idxIdCli] || '') : ''
    };
  }
  return map;
}

function getItensByLevantamento(params) {
  if (!params.id_levantamento) return _erro('CAMPO_OBRIGATORIO', 'id_levantamento é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('ITENS').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.id_levantamento) === String(params.id_levantamento)) {
      if (obj.componentes_kit) obj.componentes_kit = _parseJSON(obj.componentes_kit);
      lista.push(obj);
    }
  }
  return { status: 'ok', itens: lista };
}

// S5 — Retorna itens de uma OS, com filtro opcional por status (aceita lista separada por vírgula)
function getItensByOS(params) {
  if (!params.id_os) return _erro('CAMPO_OBRIGATORIO', 'id_os é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('ITENS').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  // Aceita status único ou lista "em_campo_f2,instalado,calibrado,entregue"
  var filtroStatus = params.status ? String(params.status).split(',').map(function(s){ return s.trim(); }) : null;
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.id_os) !== String(params.id_os)) continue;
    if (filtroStatus && filtroStatus.indexOf(String(obj.status)) < 0) continue;
    if (obj.componentes_kit) obj.componentes_kit = _parseJSON(obj.componentes_kit);
    lista.push(obj);
  }
  return { status: 'ok', itens: lista, total: lista.length };
}


// =============================================================================
// ENDPOINTS POST
// =============================================================================

function salvarCliente(body) {
  if (!body.nome) return _erro('CAMPO_OBRIGATORIO', 'nome é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('CLIENTES');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (body.id_cliente) {
      // Atualizar
      var dados = aba.getDataRange().getValues();
      var cab = dados[0];
      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(body.id_cliente)) {
          var row = i + 1;
          aba.getRange(row, 2).setValue(body.nome || dados[i][1]);
          aba.getRange(row, 3).setValue(body.cnpj || dados[i][2]);
          aba.getRange(row, 4).setValue(body.endereco || dados[i][3]);
          aba.getRange(row, 5).setValue(body.contato || dados[i][4]);
          aba.getRange(row, 6).setValue(body.observacoes || dados[i][5]);
          SpreadsheetApp.flush();
          return { status: 'ok', modo: 'atualizado', id_cliente: body.id_cliente };
        }
      }
      return _erro('ITEM_NAO_ENCONTRADO', 'Cliente não encontrado: ' + body.id_cliente);
    } else {
      // [DUPLICADO] S17 — verificar duplicado antes de criar
      var dadosVerif = aba.getDataRange().getValues();
      var cabV = dadosVerif[0];
      var idxNome = cabV.indexOf('nome');
      var idxCnpj = cabV.indexOf('cnpj');
      var idxIdC  = cabV.indexOf('id_cliente');
      var idxAtv  = cabV.indexOf('ativo');
      var cnpjNovo = String(body.cnpj || '').replace(/[^0-9]/g, '');
      var nomeNorm = _normalizarTexto(body.nome);
      for (var k = 1; k < dadosVerif.length; k++) {
        var ativoExist = idxAtv >= 0 ? String(dadosVerif[k][idxAtv]) : 'true';
        if (ativoExist === 'false') continue; // permite renomear cliente após inativar duplicado
        var cnpjExist = String(dadosVerif[k][idxCnpj] || '').replace(/[^0-9]/g, '');
        var nomeExist = _normalizarTexto(dadosVerif[k][idxNome]);
        if (cnpjNovo && cnpjExist && cnpjNovo === cnpjExist) {
          return _erroDuplicado('Cliente já cadastrado com este CNPJ.', { id_cliente: dadosVerif[k][idxIdC], nome: dadosVerif[k][idxNome] });
        }
        if (!cnpjNovo && nomeNorm && nomeNorm === nomeExist) {
          return _erroDuplicado('Cliente já cadastrado com este nome.', { id_cliente: dadosVerif[k][idxIdC], nome: dadosVerif[k][idxNome] });
        }
      }
      // Criar
      var id = _gerarIdCatalogo(ss, 'CLIENTES', 'CLI');
      aba.appendRow([id, body.nome, body.cnpj || '', body.endereco || '', body.contato || '', body.observacoes || '', 'true']);
      SpreadsheetApp.flush();
      registrarEventoOperacional({ tipo_evento: 'CRIACAO_CLIENTE', id_cliente: id, cliente: body.nome, acao_realizada: 'Cliente criado', responsavel_nome: body.responsavel || 'adm', origem: 'TapControl Manager', observacao: '' });
      return { status: 'ok', modo: 'criado', id_cliente: id };
    }
  } finally { lock.releaseLock(); }
}

// Auxiliares de normalização de status/fase (usadas em getOS e demais funções)
function _normalizarStatusOS(status) {
  return String(status || '').trim().toLowerCase();
}
function _isOSEncerradaGAS(status, fase_atual) {
  var s = _normalizarStatusOS(status);
  var f = String(fase_atual || '').trim().toLowerCase();
  return s === 'encerrada' || s === 'encerrado' || s === 'finalizada' || s === 'finalizado' || f === 'encerrada';
}

function salvarOS(body) {
  // Atualização parcial: se id_os fornecido, não exige id_cliente nem descricao
  if (!body.id_os) {
    if (!body.id_cliente) return _erro('CAMPO_OBRIGATORIO', 'id_cliente é obrigatório.');
    if (!body.descricao)  return _erro('CAMPO_OBRIGATORIO', 'descricao é obrigatório.');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('OS');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    // Busca nome do cliente
    var nomeCliente = _buscarNomeCliente(ss, body.id_cliente);
    if (body.id_os) {
      // Atualizar
      var dados = aba.getDataRange().getValues();
      var cab = dados[0];
      var colTecs = cab.indexOf('tecnicos_vinculados');
      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(body.id_os)) {
          var row = i + 1;
          if (body.descricao)  aba.getRange(row, 5).setValue(body.descricao);
          if (body.status)     aba.getRange(row, 7).setValue(body.status);
          if (body.fase_atual) aba.getRange(row, 8).setValue(body.fase_atual);
          if (body.tecnicos_vinculados !== undefined && colTecs >= 0) {
            var tecsVal = Array.isArray(body.tecnicos_vinculados) ? JSON.stringify(body.tecnicos_vinculados) : String(body.tecnicos_vinculados);
            aba.getRange(row, colTecs + 1).setValue(tecsVal);
          }
          SpreadsheetApp.flush();
          return { status: 'ok', modo: 'atualizado', id_os: body.id_os };
        }
      }
      return _erro('ITEM_NAO_ENCONTRADO', 'OS não encontrada: ' + body.id_os);
    } else {
      // [OS_DUPLICADA] BUG3-FIX (S17): verificar numero_os duplicado antes de criar nova OS
      var numeroOSPretendido = body.numero_os ? String(body.numero_os).trim() : null;
      if (numeroOSPretendido) {
        var dadosVerif = aba.getDataRange().getValues();
        var cabVerif   = dadosVerif[0];
        var colNumOS   = cabVerif.indexOf('numero_os');
        var colIdOSV   = cabVerif.indexOf('id_os');
        var colClienteV= cabVerif.indexOf('cliente');
        if (colNumOS >= 0) {
          var numNorm = _normalizarTexto(numeroOSPretendido);
          for (var v = 1; v < dadosVerif.length; v++) {
            var numExist = _normalizarTexto(dadosVerif[v][colNumOS]);
            if (numExist && numExist === numNorm) {
              var idExist = String(dadosVerif[v][colIdOSV] || '');
              var cliExist = colClienteV >= 0 ? String(dadosVerif[v][colClienteV] || '') : '';
              Logger.log('[OS_DUPLICADA] numero_os=' + numeroOSPretendido + ' já existe como id_os=' + idExist + '.');
              return _erroDuplicado('OS com numero_os=' + numeroOSPretendido + ' já existe.', { id_os: idExist, numero_os: numeroOSPretendido, cliente: cliExist });
            }
          }
        }
      }
      var id = _gerarId(ss, 'OS', 'OS');
      var data = body.data_abertura || new Date().toISOString().split('T')[0];
      // Usa numero_os do body se fornecido; caso contrário usa o id gerado
      var numeroOS = numeroOSPretendido || id;
      var tecsJson = body.tecnicos_vinculados ? (Array.isArray(body.tecnicos_vinculados) ? JSON.stringify(body.tecnicos_vinculados) : String(body.tecnicos_vinculados)) : '[]';
      aba.appendRow([id, body.id_cliente, nomeCliente, numeroOS, body.descricao, data, 'Fase 1', '1', tecsJson]);
      SpreadsheetApp.flush();
      registrarEventoOperacional({ tipo_evento: 'CRIACAO_OS', id_os: id, numero_os: numeroOS, id_cliente: body.id_cliente, cliente: nomeCliente, acao_realizada: 'OS criada', responsavel_nome: body.responsavel || 'adm', origem: 'TapControl Manager', observacao: body.descricao || '' });
      return { status: 'ok', modo: 'criado', id_os: id };
    }
  } finally { lock.releaseLock(); }
}

function salvarEquipamento(body) {
  if (!body.id_cliente) return _erro('CAMPO_OBRIGATORIO', 'id_cliente é obrigatório.');
  if (!body.tipo)       return _erro('CAMPO_OBRIGATORIO', 'tipo é obrigatório.');
  if (!body.tag)        return _erro('CAMPO_OBRIGATORIO', 'tag é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('EQUIPAMENTOS');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (body.id_equipamento) {
      // Atualizar
      var dados = aba.getDataRange().getValues();
      // Correção S10 — vínculo ativo equipamento/OS: localizar coluna id_os pelo cabeçalho
      var cab = dados[0];
      var colIdOS = cab.indexOf('id_os');
      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(body.id_equipamento)) {
          var row = i + 1;
          if (body.tipo)        aba.getRange(row, 4).setValue(body.tipo);
          if (body.tag)         aba.getRange(row, 5).setValue(body.tag);
          if (body.localizacao) aba.getRange(row, 6).setValue(body.localizacao);
          if (body.pmta)        aba.getRange(row, 7).setValue(body.pmta);
          if (body.campo_livre) aba.getRange(row, 8).setValue(body.campo_livre);
          // Correção S10 — vínculo ativo: atualiza id_os sempre que vier no body
          // Aceita string vazia "" para desvincular. NÃO usar "if (body.id_os)" pois impediria limpeza.
          if (body.id_os !== undefined && colIdOS >= 0) {
            aba.getRange(row, colIdOS + 1).setValue(body.id_os || '');
          }
          SpreadsheetApp.flush();
          return { status: 'ok', modo: 'atualizado', id_equipamento: body.id_equipamento };
        }
      }
      return _erro('ITEM_NAO_ENCONTRADO', 'Equipamento não encontrado: ' + body.id_equipamento);
    } else {
      // [DUPLICADO] S17 — bloquear equipamento mesmo cliente + mesma TAG normalizada (ignora id_os)
      var dadosVerif = aba.getDataRange().getValues();
      var cabV = dadosVerif[0];
      var idxIdEq = cabV.indexOf('id_equipamento');
      var idxCli  = cabV.indexOf('id_cliente');
      var idxTag  = cabV.indexOf('tag');
      var idxAtv  = cabV.indexOf('ativo');
      var tagNorm = _normalizarTexto(body.tag);
      var cliNovo = String(body.id_cliente);
      for (var k = 1; k < dadosVerif.length; k++) {
        var ativoExist = idxAtv >= 0 ? String(dadosVerif[k][idxAtv]) : 'true';
        if (ativoExist === 'false') continue;
        var cliExist = String(dadosVerif[k][idxCli]);
        var tagExist = _normalizarTexto(dadosVerif[k][idxTag]);
        if (cliExist === cliNovo && tagExist && tagExist === tagNorm) {
          return _erroDuplicado('Equipamento já existe para este cliente com a mesma TAG.', {
            id_equipamento: dadosVerif[k][idxIdEq],
            tag: dadosVerif[k][idxTag]
          });
        }
      }
      var id = _gerarIdCatalogo(ss, 'EQUIPAMENTOS', 'EQ');
      var cadastradoPor = body.cadastrado_por || 'adm'; // 'adm' ou 'inspetor'
      aba.appendRow([id, body.id_cliente, body.id_os || '', body.tipo, body.tag, body.localizacao || '', body.pmta || '', body.campo_livre || '', cadastradoPor, 'true']);
      SpreadsheetApp.flush();
      return { status: 'ok', modo: 'criado', id_equipamento: id };
    }
  } finally { lock.releaseLock(); }
}

function enviarLevantamento(body) {
  if (!body.id_inspetor) return _erro('CAMPO_OBRIGATORIO', 'id_inspetor é obrigatório.');
  if (!body.id_os)       return _erro('CAMPO_OBRIGATORIO', 'id_os é obrigatório.');
  if (!body.itens || !body.itens.length) return _erro('CAMPO_OBRIGATORIO', 'itens não pode ser vazio.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    _garantirColunas(ss, 'LEVANTAMENTOS', [
      'drive_folder_id','drive_folder_url',
      'quantidade_fotos_equip','quantidade_fotos_ponto',
      'fotos_equip_json','fotos_ponto_json'
    ]);

    var nomeInspetor = _buscarNomeInspetor(ss, body.id_inspetor);
    var timestamp = body.timestamp || new Date().toISOString();
    var fase = body.fase || '1';
    var id_cliente = body.id_cliente || _buscarClienteDaOS(ss, body.id_os);

    // [CANCELAMENTO_ENVIO] BUG1-FIX: se id_levantamento foi passado (reenvio após cancelamento),
    // reusar o mesmo. Caso contrário, verificar se já existe levantamento aberto para os+fase.
    var idLevantamento = null;
    var modoEnvio = 'novo';

    if (body.id_levantamento) {
      // Reenvio explícito após cancelamento — verificar se existe
      var abaLev = ss.getSheetByName('LEVANTAMENTOS');
      var dadosLev = abaLev.getDataRange().getValues();
      var cabLev = dadosLev[0];
      var colLevId = cabLev.indexOf('id_levantamento');
      var colLevOS = cabLev.indexOf('id_os');
      var colLevStatus = cabLev.indexOf('status_geral');
      for (var lv = 1; lv < dadosLev.length; lv++) {
        if (String(dadosLev[lv][colLevId] || '') === String(body.id_levantamento)) {
          idLevantamento = body.id_levantamento;
          modoEnvio = 'reenvio';
          Logger.log('[CANCELAMENTO_ENVIO] Reenvio após cancelamento — reusando levantamento=' + idLevantamento);
          // Atualiza status_geral de volta para aguardando_cotacao e timestamp
          if (colLevStatus >= 0) abaLev.getRange(lv + 1, colLevStatus + 1).setValue('aguardando_cotacao');
          abaLev.getRange(lv + 1, cabLev.indexOf('timestamp_envio') + 1).setValue(timestamp);
          abaLev.getRange(lv + 1, cabLev.indexOf('total_itens') + 1).setValue(body.itens.length);
          SpreadsheetApp.flush();
          break;
        }
      }
    }

    if (!idLevantamento) {
      // Verificar se já existe levantamento em_correcao_campo ou aguardando_cotacao para esta OS+fase
      var abaLev2 = ss.getSheetByName('LEVANTAMENTOS');
      var dadosLev2 = abaLev2.getDataRange().getValues();
      var cabLev2 = dadosLev2[0];
      var colL2Id = cabLev2.indexOf('id_levantamento');
      var colL2OS = cabLev2.indexOf('id_os');
      var colL2Fase = cabLev2.indexOf('fase');
      var colL2Status = cabLev2.indexOf('status_geral');
      for (var lv2 = 1; lv2 < dadosLev2.length; lv2++) {
        if (String(dadosLev2[lv2][colL2OS] || '') === String(body.id_os) &&
            String(dadosLev2[lv2][colL2Fase] || '') === String(fase) &&
            ['em_correcao_campo', 'cancelado_para_correcao'].includes(String(dadosLev2[lv2][colL2Status] || ''))) {
          idLevantamento = String(dadosLev2[lv2][colL2Id] || '');
          modoEnvio = 'reenvio_correcao';
          Logger.log('[CANCELAMENTO_ENVIO] Levantamento em correção encontrado — reusando id=' + idLevantamento);
          if (colL2Status >= 0) abaLev2.getRange(lv2 + 1, colL2Status + 1).setValue('aguardando_cotacao');
          abaLev2.getRange(lv2 + 1, cabLev2.indexOf('timestamp_envio') + 1).setValue(timestamp);
          abaLev2.getRange(lv2 + 1, cabLev2.indexOf('total_itens') + 1).setValue(body.itens.length);
          SpreadsheetApp.flush();
          break;
        }
      }
    }

    if (!idLevantamento) {
      // Criar novo levantamento normalmente
      idLevantamento = _gerarId(ss, 'LEVANTAMENTOS', 'LEV');
      ss.getSheetByName('LEVANTAMENTOS').appendRow([
        idLevantamento, body.id_os, id_cliente, body.id_inspetor,
        nomeInspetor, fase, timestamp, body.itens.length, 'aguardando_cotacao'
      ]);
    }

    var abaItens = ss.getSheetByName('ITENS');
    var itensCriados = [];

    // [CANCELAMENTO_ENVIO] BUG1-FIX: se reenvio, apagar itens antigos deste levantamento
    // antes de reinserir os itens corrigidos (evita duplicidade de itens)
    if (modoEnvio === 'reenvio' || modoEnvio === 'reenvio_correcao') {
      var dadosItensAntigos = abaItens.getDataRange().getValues();
      var cabItens = dadosItensAntigos[0];
      var colItenLev = cabItens.indexOf('id_levantamento');
      var rowsParaExcluir = [];
      for (var ri = 1; ri < dadosItensAntigos.length; ri++) {
        if (String(dadosItensAntigos[ri][colItenLev] || '') === String(idLevantamento)) {
          rowsParaExcluir.push(ri + 1);
        }
      }
      // Apagar de baixo para cima para não deslocar índices
      for (var rd = rowsParaExcluir.length - 1; rd >= 0; rd--) {
        abaItens.deleteRow(rowsParaExcluir[rd]);
      }
      SpreadsheetApp.flush();
      Logger.log('[CANCELAMENTO_ENVIO] Itens antigos do levantamento ' + idLevantamento + ' removidos (' + rowsParaExcluir.length + ' linhas).');
    }

    body.itens.forEach(function(item) {
      var idItem = _gerarId(ss, 'ITENS', 'ACE');
      var agora  = new Date().toISOString();
      // Pipeline: instalado_imediatamente = true → status inicial diferente
      var instaladoAgora = String(item.instalado_imediatamente).toLowerCase() === 'true';
      var statusInicial  = instaladoAgora ? 'instalado_campo' : 'aguardando_cotacao';

      // S18 — mapeamento dinâmico por cabeçalho para incluir colunas novas automaticamente
      var cabItensH = abaItens.getRange(1, 1, 1, abaItens.getLastColumn()).getValues()[0];
      var itemBase = {
        id_item:                  idItem,
        id_levantamento:          idLevantamento,
        id_os:                    body.id_os,
        id_cliente:               id_cliente,
        id_equipamento:           item.id_equipamento   || '',
        tag_equipamento:          item.tag_equipamento  || '',
        tipo_acessorio:           item.tipo_acessorio   || '',
        id_item_catalogo:         item.id_item_catalogo || '',
        descricao_curta:          item.descricao_curta  || '',
        quantidade:               item.quantidade       || 1,
        escala_inicio:            item.escala_inicio    || '',
        escala_fim:               item.escala_fim       || '',
        escala_unidade:           item.escala_unidade   || '',
        id_kit:                   item.id_kit           || '',
        componentes_kit:          item.componentes_kit ? JSON.stringify(item.componentes_kit) : '',
        ajuste_kit:               item.ajuste_kit       || '',
        requer_calibracao:        item.requer_calibracao || 'false',
        requer_lacre:             item.requer_lacre     || 'false',
        instalado_imediatamente:  instaladoAgora ? 'true' : 'false',
        status:                   statusInicial,
        fornecedor: '', valor_cotado: '', prazo_entrega: '', status_aprovacao_cliente: 'pendente',
        data_calibracao: '', num_certificado: '', num_lacre: '', validade_calibracao: '', lab_calibracao: '',
        data_instalacao: '', data_entrega: '', observacao_fechamento: '',
        timestamp_criacao:        agora,
        timestamp_atualizacao:    agora,
        origem:                   'tapcontrol_field',
        // S18 — novos campos
        conexoes_montagem_json:   item.conexoes_montagem ? JSON.stringify(item.conexoes_montagem) : (item.conexoes_montagem_json || ''),
        itens_adicionais_json:    item.itens_adicionais  ? JSON.stringify(item.itens_adicionais)  : (item.itens_adicionais_json  || ''),
        origem_item_manual:       item._veio_de_outro ? 'manual/outro' : '',
        dados_item_manual_json:   JSON.stringify({ tipo: item.tipo, tipo_acessorio: item.tipo_acessorio, descricao: item.descricao, desc_curta: item.descricao_curta }),
        payload_item_json:        JSON.stringify(item),  // blindagem — nenhum campo futuro se perde
        foto_equipamento_b64:      '',  // S22 — base64 não gravado; fotos ficam no Drive
        foto_ponto_instalacao_b64: ''   // metadados gravados na aba LEVANTAMENTOS
      };
      var linhaItem = cabItensH.map(function(col) { return itemBase[col] !== undefined ? itemBase[col] : ''; });
      abaItens.appendRow(linhaItem);

      itensCriados.push({ id_item: idItem, descricao_curta: item.descricao_curta || '', instalado_imediatamente: instaladoAgora });
    });

    if (String(fase) === '1') {
      var abaOS = ss.getSheetByName('OS');
      var dadosOS = abaOS.getDataRange().getValues();
      var cabOS = dadosOS[0];
      var colIdOS = cabOS.indexOf('id_os');
      var colFase = cabOS.indexOf('fase_atual');
      var colStatus = cabOS.indexOf('status');
      for (var j = 1; j < dadosOS.length; j++) {
        if (String(dadosOS[j][colIdOS]) === String(body.id_os)) {
          if (colFase >= 0) abaOS.getRange(j + 1, colFase + 1).setValue('aguardando_tapparts');
          if (colStatus >= 0) abaOS.getRange(j + 1, colStatus + 1).setValue('Aguardando TapControl Manager');
          break;
        }
      }
      SpreadsheetApp.flush();
      registrarEventoOperacional({ tipo_evento: 'FASE1_ENVIADA_FIELDTP', id_os: body.id_os, id_cliente: id_cliente, acao_realizada: 'Fase 1 enviada pelo TapControl Field; aguardando tratamento no TapControl Manager', responsavel_id: body.id_inspetor, responsavel_nome: nomeInspetor, responsavel_tipo: 'inspetor', origem: 'TapControl Field' });
    }

    // ── FOTOS NO DRIVE ──
    var fotosInfo = {
      drive_folder_id: '', drive_folder_url: '',
      fotos_salvas_equip: [], fotos_salvas_ponto: [],
      fotos_com_erro: [],
      quantidade_fotos_equip: 0, quantidade_fotos_ponto: 0
    };
    var temFotos = (body.fotos_equipamentos && Object.keys(body.fotos_equipamentos).some(function(k){ return !!body.fotos_equipamentos[k]; })) ||
                   (body.fotos_pontos_instalacao && Object.keys(body.fotos_pontos_instalacao).some(function(k){ return !!body.fotos_pontos_instalacao[k]; }));

    if (temFotos) {
      try {
        var driveIdExistente = '';
        var abaLevCheck = ss.getSheetByName('LEVANTAMENTOS');
        var dadosLevCheck = abaLevCheck.getDataRange().getValues();
        var cabLevCheck = dadosLevCheck[0];
        for (var lc = 1; lc < dadosLevCheck.length; lc++) {
          if (String(dadosLevCheck[lc][cabLevCheck.indexOf('id_levantamento')] || '') === String(idLevantamento)) {
            driveIdExistente = String(dadosLevCheck[lc][cabLevCheck.indexOf('drive_folder_id')] || '');
            break;
          }
        }
        fotosInfo = salvarFotosDriveTapControl({
          id_os: body.id_os,
          id_levantamento: idLevantamento,
          id_inspetor: body.id_inspetor,
          numero_os: body.numero_os || body.id_os,
          drive_folder_id_existente: driveIdExistente,
          fotos_equipamentos: body.fotos_equipamentos || {},
          fotos_pontos_instalacao: body.fotos_pontos_instalacao || {}
        });
      } catch(errFotos) {
        logErro('enviarLevantamento/fotos', errFotos.message, body.id_os);
      }
    }

    if (fotosInfo.drive_folder_id) {
      var abaLevF = ss.getSheetByName('LEVANTAMENTOS');
      var dadosLevF = abaLevF.getDataRange().getValues();
      var cabLevF = dadosLevF[0];
      for (var lf = 1; lf < dadosLevF.length; lf++) {
        if (String(dadosLevF[lf][cabLevF.indexOf('id_levantamento')] || '') === String(idLevantamento)) {
          _setColIfExists(abaLevF, lf + 1, cabLevF, 'drive_folder_id', fotosInfo.drive_folder_id);
          _setColIfExists(abaLevF, lf + 1, cabLevF, 'drive_folder_url', fotosInfo.drive_folder_url);
          _setColIfExists(abaLevF, lf + 1, cabLevF, 'quantidade_fotos_equip', fotosInfo.quantidade_fotos_equip);
          _setColIfExists(abaLevF, lf + 1, cabLevF, 'quantidade_fotos_ponto', fotosInfo.quantidade_fotos_ponto);
          _setColIfExists(abaLevF, lf + 1, cabLevF, 'fotos_equip_json', JSON.stringify(fotosInfo.fotos_salvas_equip));
          _setColIfExists(abaLevF, lf + 1, cabLevF, 'fotos_ponto_json', JSON.stringify(fotosInfo.fotos_salvas_ponto));
          SpreadsheetApp.flush();
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    registrarEventoOperacional({ tipo_evento: 'ENVIO_LEVANTAMENTO', id_os: body.id_os, id_cliente: id_cliente, acao_realizada: 'Levantamento enviado pelo TapControl Field', responsavel_id: body.id_inspetor, responsavel_nome: nomeInspetor, responsavel_tipo: 'inspetor', origem: 'TapControl Field', observacao: 'Total itens: ' + body.itens.length, payload_json: JSON.stringify({ id_levantamento: idLevantamento }) });
    return {
      status: 'ok',
      id_levantamento: idLevantamento,
      itens_criados: itensCriados,
      fotos: {
        drive_folder_url: fotosInfo.drive_folder_url || '',
        quantidade_fotos_equip: fotosInfo.quantidade_fotos_equip || 0,
        quantidade_fotos_ponto: fotosInfo.quantidade_fotos_ponto || 0,
        erros: (fotosInfo.fotos_com_erro || []).length
      }
    };

  } finally { lock.releaseLock(); }
}

function atualizarStatusItem(body) {
  if (!body.id_item)     return _erro('CAMPO_OBRIGATORIO', 'id_item é obrigatório.');
  if (!body.novo_status) return _erro('CAMPO_OBRIGATORIO', 'novo_status é obrigatório.');
  if (String(body.novo_status) === 'em_campo_f2' && String(body.origem) !== 'ENVIO_FORMAL_FASE2' && body.autorizacao_fase2 !== true) {
    return _erro('TRANSICAO_FASE2_REQUER_ENVIO_FORMAL', 'Status em_campo_f2 só pode ser aplicado pelo envio formal da Fase 2.');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var abaItens = ss.getSheetByName('ITENS');
    var dados = abaItens.getDataRange().getValues();
    var cab = dados[0];
    var colStatus = cab.indexOf('status');
    var colTs     = cab.indexOf('timestamp_atualizacao');
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id_item)) {
        var statusAnterior = dados[i][colStatus];
        var agora = new Date().toISOString();
        abaItens.getRange(i + 1, colStatus + 1).setValue(body.novo_status);
        abaItens.getRange(i + 1, colTs + 1).setValue(agora);
        ss.getSheetByName('HISTORICO_STATUS').appendRow([
          _gerarIdHistorico(ss), body.id_item, statusAnterior, body.novo_status,
          body.responsavel || '', body.observacao || '', agora
        ]);
        SpreadsheetApp.flush();
        var tipoEvento = String(body.novo_status) === 'em_campo_f2' ? 'ENVIO_FORMAL_FASE2' : 'ALTERACAO_STATUS_ITEM';
        registrarEventoOperacional({ tipo_evento: tipoEvento, id_os: dados[i][cab.indexOf('id_os')] || '', numero_os: dados[i][cab.indexOf('id_os')] || '', id_item: body.id_item, produto: dados[i][cab.indexOf('descricao_curta')] || '', tipo_acessorio: dados[i][cab.indexOf('tipo_acessorio')] || '', status_anterior: statusAnterior, status_novo: body.novo_status, acao_realizada: 'Status do item alterado', responsavel_nome: body.responsavel || '', responsavel_tipo: 'adm', origem: body.origem || 'TapControl Manager', observacao: body.observacao || '' });
        return { status: 'ok', id_item: body.id_item, novo_status: body.novo_status, timestamp: agora };
      }
    }
    return _erro('ITEM_NAO_ENCONTRADO', 'Item não encontrado: ' + body.id_item);
  } finally { lock.releaseLock(); }
}

function atualizarDadosAdministrativos(body) {
  if (!body.id_item) return _erro('CAMPO_OBRIGATORIO', 'id_item é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName('ITENS');
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var cols = {
      fornecedor:               cab.indexOf('fornecedor'),
      valor_cotado:             cab.indexOf('valor_cotado'),
      prazo_entrega:            cab.indexOf('prazo_entrega'),
      status_aprovacao_cliente: cab.indexOf('status_aprovacao_cliente'),
      status:                   cab.indexOf('status'),
      observacao_cotacao:       cab.indexOf('observacao_cotacao'),  // S18
      ts:                       cab.indexOf('timestamp_atualizacao')
    };
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id_item)) {
        var agora = new Date().toISOString();
        var statusAnterior = String(dados[i][cols.status] || '');
        if (body.fornecedor !== undefined)               aba.getRange(i+1, cols.fornecedor+1).setValue(body.fornecedor);
        if (body.valor_cotado !== undefined)             aba.getRange(i+1, cols.valor_cotado+1).setValue(body.valor_cotado);
        if (body.prazo_entrega !== undefined)            aba.getRange(i+1, cols.prazo_entrega+1).setValue(body.prazo_entrega);
        if (body.status_aprovacao_cliente !== undefined) aba.getRange(i+1, cols.status_aprovacao_cliente+1).setValue(body.status_aprovacao_cliente);
        // S18 — gravar status_novo e observacao_cotacao
        if (body.status_novo && cols.status >= 0)            aba.getRange(i+1, cols.status+1).setValue(body.status_novo);
        if (body.observacao_cotacao !== undefined && cols.observacao_cotacao >= 0) aba.getRange(i+1, cols.observacao_cotacao+1).setValue(body.observacao_cotacao);
        aba.getRange(i+1, cols.ts+1).setValue(agora);
        var statusNovoDef = body.status_novo || statusAnterior;
        ss.getSheetByName('HISTORICO_STATUS').appendRow([
          _gerarIdHistorico(ss), body.id_item, statusAnterior,
          body.status_novo ? body.status_novo : 'dados_adm_atualizados',
          body.responsavel || '', body.observacao || body.observacao_cotacao || '', agora
        ]);
        SpreadsheetApp.flush();
        registrarEventoOperacional({ tipo_evento: 'ATUALIZACAO_DADOS_ADM', id_os: dados[i][cab.indexOf('id_os')] || '', id_item: body.id_item, produto: dados[i][cab.indexOf('descricao_curta')] || '', tipo_acessorio: dados[i][cab.indexOf('tipo_acessorio')] || '', status_anterior: statusAnterior, status_novo: statusNovoDef, acao_realizada: 'Dados administrativos atualizados', responsavel_nome: body.responsavel || '', fornecedor: body.fornecedor || '', valor: body.valor_cotado || '', origem: 'TapControl Manager', observacao: body.observacao_cotacao || body.observacao || '' });
        return { status: 'ok', id_item: body.id_item, timestamp: agora };
      }
    }
    return _erro('ITEM_NAO_ENCONTRADO', 'Item não encontrado: ' + body.id_item);
  } finally { lock.releaseLock(); }
}

function registrarCalibracao(body) {
  if (!body.id_item) return _erro('CAMPO_OBRIGATORIO', 'id_item é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName('ITENS');
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var cols = {
      data_calibracao:    cab.indexOf('data_calibracao'),
      num_certificado:    cab.indexOf('num_certificado'),
      num_lacre:          cab.indexOf('num_lacre'),
      num_lacre2:         cab.indexOf('num_lacre2'),         // S18
      validade_calibracao:cab.indexOf('validade_calibracao'),
      lab_calibracao:     cab.indexOf('lab_calibracao'),
      tipo_calibracao:    cab.indexOf('tipo_calibracao'),    // S18
      status:             cab.indexOf('status'),
      ts:                 cab.indexOf('timestamp_atualizacao')
    };
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id_item)) {
        var agora = new Date().toISOString();
        var statusAnterior = String(dados[i][cols.status] || '');
        var tipoCalib = body.tipo_calibracao || 'Interna';

        if (body.data_calibracao && cols.data_calibracao >= 0)      aba.getRange(i+1, cols.data_calibracao+1).setValue(body.data_calibracao);
        if (body.num_certificado && cols.num_certificado >= 0)       aba.getRange(i+1, cols.num_certificado+1).setValue(body.num_certificado);
        if (body.num_lacre && cols.num_lacre >= 0)                   aba.getRange(i+1, cols.num_lacre+1).setValue(body.num_lacre);
        if (body.num_lacre2 && cols.num_lacre2 >= 0)                 aba.getRange(i+1, cols.num_lacre2+1).setValue(body.num_lacre2);  // S18
        if (body.validade_calibracao && cols.validade_calibracao>=0) aba.getRange(i+1, cols.validade_calibracao+1).setValue(body.validade_calibracao);
        if (body.lab_calibracao && cols.lab_calibracao >= 0)         aba.getRange(i+1, cols.lab_calibracao+1).setValue(body.lab_calibracao);
        if (cols.tipo_calibracao >= 0)                               aba.getRange(i+1, cols.tipo_calibracao+1).setValue(tipoCalib);  // S18

        // S18 — só avança para 'calibrado' se não for lacre avulso
        var novoStatus = statusAnterior;
        if (tipoCalib !== 'Lacre avulso') {
          novoStatus = 'calibrado';
          aba.getRange(i+1, cols.status+1).setValue(novoStatus);
        }
        aba.getRange(i+1, cols.ts+1).setValue(agora);

        ss.getSheetByName('HISTORICO_STATUS').appendRow([
          _gerarIdHistorico(ss), body.id_item, statusAnterior, novoStatus,
          body.responsavel || '', body.observacao || '', agora
        ]);

        // S18 — inserir em HISTORICO_CALIBRACOES
        var idCalib = 'CAL-' + agora.replace(/[^0-9]/g,'').substring(0,14);
        var abaHCalib = ss.getSheetByName('HISTORICO_CALIBRACOES');
        if (abaHCalib) {
          abaHCalib.appendRow([
            idCalib,
            body.id_item,
            dados[i][cab.indexOf('id_os')] || '',
            dados[i][cab.indexOf('id_equipamento')] || '',
            dados[i][cab.indexOf('tag_equipamento')] || '',
            tipoCalib,
            body.data_calibracao || '',
            body.validade_calibracao || '',
            body.num_certificado || '',
            body.num_lacre || '',
            body.num_lacre2 || '',
            body.lab_calibracao || '',
            body.responsavel || '',
            body.observacao || '',
            agora,
            body.origem || 'TapControl Manager'
          ]);
        }

        SpreadsheetApp.flush();
        registrarEventoOperacional({ tipo_evento: 'REGISTRO_CALIBRACAO', id_os: dados[i][cab.indexOf('id_os')] || '', id_item: body.id_item, produto: dados[i][cab.indexOf('descricao_curta')] || '', tipo_acessorio: dados[i][cab.indexOf('tipo_acessorio')] || '', status_anterior: statusAnterior, status_novo: novoStatus, acao_realizada: 'Calibração registrada — tipo: ' + tipoCalib, responsavel_nome: body.responsavel || '', origem: body.origem || 'TapControl Manager', observacao: 'Cert: ' + (body.num_certificado || '') + ' Lacre: ' + (body.num_lacre || '') });
        return { status: 'ok', id_item: body.id_item, novo_status: novoStatus, timestamp: agora };
      }
    }
    return _erro('ITEM_NAO_ENCONTRADO', 'Item não encontrado: ' + body.id_item);
  } finally { lock.releaseLock(); }
}

function registrarInstalacao(body) {
  if (!body.id_item) return _erro('CAMPO_OBRIGATORIO', 'id_item é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName('ITENS');
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var cols = {
      data_instalacao:      cab.indexOf('data_instalacao'),
      data_entrega:         cab.indexOf('data_entrega'),
      data_calibracao:      cab.indexOf('data_calibracao'),
      num_certificado:      cab.indexOf('num_certificado'),
      num_lacre:            cab.indexOf('num_lacre'),
      observacao_fechamento:cab.indexOf('observacao_fechamento'),
      status:               cab.indexOf('status'),
      ts:                   cab.indexOf('timestamp_atualizacao')
    };
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id_item)) {
        var agora = new Date().toISOString();
        // S5: status vem do body (instalado / entregue / calibrado) ou inferido
        var statusAnterior = String(dados[i][cols.status]);
        var novoStatus = body.status ||
          (body.tipo === 'entrega' ? 'entregue' : 'instalado');
        if (body.data_instalacao && cols.data_instalacao >= 0)
          aba.getRange(i+1, cols.data_instalacao+1).setValue(body.data_instalacao);
        if (body.data_entrega && cols.data_entrega >= 0)
          aba.getRange(i+1, cols.data_entrega+1).setValue(body.data_entrega);
        if (body.data_calibracao && cols.data_calibracao >= 0)
          aba.getRange(i+1, cols.data_calibracao+1).setValue(body.data_calibracao);
        if (body.num_certificado && cols.num_certificado >= 0)
          aba.getRange(i+1, cols.num_certificado+1).setValue(body.num_certificado);
        if (body.num_lacre && cols.num_lacre >= 0)
          aba.getRange(i+1, cols.num_lacre+1).setValue(body.num_lacre);
        if (body.observacao_fechamento && cols.observacao_fechamento >= 0)
          aba.getRange(i+1, cols.observacao_fechamento+1).setValue(body.observacao_fechamento);
        if (body.observacao && cols.observacao_fechamento >= 0)
          aba.getRange(i+1, cols.observacao_fechamento+1).setValue(body.observacao);
        aba.getRange(i+1, cols.status+1).setValue(novoStatus);
        aba.getRange(i+1, cols.ts+1).setValue(agora);
        ss.getSheetByName('HISTORICO_STATUS').appendRow([
          _gerarIdHistorico(ss), body.id_item, statusAnterior, novoStatus,
          body.responsavel || '', body.observacao_fechamento || body.observacao || '', agora
        ]);
        SpreadsheetApp.flush();
        registrarEventoOperacional({ tipo_evento: novoStatus === 'entregue' ? 'REGISTRO_ENTREGA' : 'REGISTRO_INSTALACAO', id_os: dados[i][cab.indexOf('id_os')] || '', id_item: body.id_item, produto: dados[i][cab.indexOf('descricao_curta')] || '', tipo_acessorio: dados[i][cab.indexOf('tipo_acessorio')] || '', status_anterior: statusAnterior, status_novo: novoStatus, acao_realizada: 'Instalação/entrega registrada', responsavel_nome: body.responsavel || '', origem: 'TapControl Manager', observacao: body.observacao_fechamento || body.observacao || '' });
        return { status: 'ok', id_item: body.id_item, novo_status: novoStatus, timestamp: agora };
      }
    }
    return _erro('ITEM_NAO_ENCONTRADO', 'Item não encontrado: ' + body.id_item);
  } finally { lock.releaseLock(); }
}

function avancarFaseOS(body) {
  if (!body.id_os)    return _erro('CAMPO_OBRIGATORIO', 'id_os é obrigatório.');
  if (!body.nova_fase) return _erro('CAMPO_OBRIGATORIO', 'nova_fase é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName('OS');
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var colFase   = cab.indexOf('fase_atual');
    var colStatus = cab.indexOf('status');
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id_os)) {
        var novaFaseStr = String(body.nova_fase);
        aba.getRange(i+1, colFase+1).setValue(novaFaseStr);
        // Atualiza coluna status para refletir o estado visual
        if (colStatus >= 0) {
          var novoStatus = novaFaseStr === 'encerrada'       ? 'Encerrada'
                       : novaFaseStr === 'campo_concluido'  ? 'Fase 2 Campo Concluída'
                       : novaFaseStr === '2'               ? 'Fase 2'
                       : 'Fase ' + novaFaseStr;
          aba.getRange(i+1, colStatus+1).setValue(novoStatus);
        }
        // Registra observação se enviada
        if (body.observacao) {
          var colDesc = cab.indexOf('descricao');
          // não sobrescreve descricao; apenas loga via flush
        }
        // S18 — reverter itens em_campo_f2 → pronto_campo quando retroce para fase 1
        if (novaFaseStr === '1') {
          var abaItensF1 = ss.getSheetByName('ITENS');
          var dadosItensF1 = abaItensF1.getDataRange().getValues();
          var cabItensF1 = dadosItensF1[0];
          var colF1OS     = cabItensF1.indexOf('id_os');
          var colF1Status = cabItensF1.indexOf('status');
          var colF1Ts     = cabItensF1.indexOf('timestamp_atualizacao');
          var agoraF1 = new Date().toISOString();
          for (var ri = 1; ri < dadosItensF1.length; ri++) {
            if (String(dadosItensF1[ri][colF1OS]) === String(body.id_os) &&
                String(dadosItensF1[ri][colF1Status]) === 'em_campo_f2') {
              abaItensF1.getRange(ri+1, colF1Status+1).setValue('pronto_campo');
              if (colF1Ts >= 0) abaItensF1.getRange(ri+1, colF1Ts+1).setValue(agoraF1);
            }
          }
          Logger.log('[avancarFaseOS] Itens em_campo_f2 revertidos para pronto_campo — OS: ' + body.id_os);
        }

        SpreadsheetApp.flush();
        registrarEventoOperacional({ tipo_evento: 'AVANCO_FASE_OS', id_os: body.id_os, numero_os: body.id_os, acao_realizada: 'Fase da OS avançada para ' + novaFaseStr, responsavel_nome: body.responsavel || 'adm', origem: body.origem || 'TapControl Manager', observacao: body.observacao || '' });
        return { status: 'ok', id_os: body.id_os, fase_atual: novaFaseStr };
      }
    }
    return _erro('ITEM_NAO_ENCONTRADO', 'OS não encontrada: ' + body.id_os);
  } finally { lock.releaseLock(); }
}

function salvarItemCatalogo(body) {
  if (!body.tipo)       return _erro('CAMPO_OBRIGATORIO', 'tipo é obrigatório.');
  // Compatibilidade: aceita body.dados_item (legado) ou campos diretos no body (TapParts)
  var d = body.dados_item ? body.dados_item : body;
  var mapa = {
    manometro:  { aba: 'CAT_MANOMETROS', prefixo: 'MAN' },
    manometros: { aba: 'CAT_MANOMETROS', prefixo: 'MAN' },
    valvula:    { aba: 'CAT_VALVULAS',   prefixo: 'VAL' },
    valvulas:   { aba: 'CAT_VALVULAS',   prefixo: 'VAL' },
    purgador:   { aba: 'CAT_PURGADORES', prefixo: 'PUR' },
    purgadores: { aba: 'CAT_PURGADORES', prefixo: 'PUR' },
    dcbi:       { aba: 'CAT_DCBI',       prefixo: 'DCBI' },
    conexao:    { aba: 'CAT_CONEXOES',   prefixo: 'CON' },
    conexoes:   { aba: 'CAT_CONEXOES',   prefixo: 'CON' },
    kits:       { aba: 'CAT_KITS',       prefixo: 'KIT' }
  };
  var cfg = mapa[body.tipo];
  if (!cfg) return _erro('CATALOGO_INVALIDO', 'Tipo inválido: ' + body.tipo);
  // [SALVAR_CAT_FIX] S17 — desacoplar roteador de coluna "tipo".
  // Se o client mandou `_tipo_dado`, ele é o valor real da coluna `tipo` (ex.: "Termodinâmico", "Niple", "Manômetro").
  // Sem isso, `d.tipo` ficaria com o roteador ("purgador"/"conexao") e o banco gravaria errado.
  if (d._tipo_dado !== undefined && d._tipo_dado !== null && d._tipo_dado !== '') {
    d.tipo = d._tipo_dado;
  } else if (body.tipo === 'purgador' || body.tipo === 'purgadores' ||
             body.tipo === 'conexao'  || body.tipo === 'conexoes') {
    // Defensivo: se chegou aqui sem _tipo_dado, NÃO grava o roteador como tipo (evita poluir banco).
    delete d.tipo;
  }
  delete d._tipo_dado;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName(cfg.aba);
    // Gera desc_curta para manômetros automaticamente
    if (body.tipo === 'manometros' || body.tipo === 'manometro') d.desc_curta = d.desc_curta || _gerarDescCurtaMano(d);
    if ((body.tipo === 'kits') && Array.isArray(d.componentes)) d.componentes = JSON.stringify(d.componentes);
    if (body.tipo === 'dcbi' && Array.isArray(d.componentes)) d.componentes = JSON.stringify(d.componentes);

    if (d.id) {
      // Atualizar
      var dados = aba.getDataRange().getValues();
      var cab = dados[0];
      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(d.id)) {
          cab.forEach(function(col, ci) {
            if (d[col] !== undefined) aba.getRange(i+1, ci+1).setValue(d[col]);
          });
          SpreadsheetApp.flush();
          return { status: 'ok', modo: 'atualizado', id: d.id, desc_curta: d.desc_curta || '' };
        }
      }
    }
    // Criar
    var novoId = _gerarIdCatalogo(ss, cfg.aba, cfg.prefixo);
    d.id = novoId;
    var cab2 = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    var linha = cab2.map(function(col) { return d[col] !== undefined ? d[col] : ''; });
    aba.appendRow(linha);
    SpreadsheetApp.flush();
    return { status: 'ok', modo: 'criado', id: novoId, desc_curta: d.desc_curta || '' };
  } finally { lock.releaseLock(); }
}

function toggleAtivoItem(body) {
  if (!body.tipo) return _erro('CAMPO_OBRIGATORIO', 'tipo é obrigatório.');
  if (!body.id)   return _erro('CAMPO_OBRIGATORIO', 'id é obrigatório.');
  var mapa = {
    manometros: 'CAT_MANOMETROS', valvulas: 'CAT_VALVULAS',
    purgadores: 'CAT_PURGADORES', dcbi: 'CAT_DCBI',
    conexoes: 'CAT_CONEXOES', kits: 'CAT_KITS',
    clientes: 'CLIENTES', equipamentos: 'EQUIPAMENTOS'
  };
  var nomeAba = mapa[body.tipo];
  if (!nomeAba) return _erro('CATALOGO_INVALIDO', 'Tipo inválido: ' + body.tipo);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName(nomeAba);
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var colAtivo = cab.indexOf('ativo');
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id)) {
        var novoAtivo = String(body.ativo) === 'true' ? 'true' : 'false';
        aba.getRange(i+1, colAtivo+1).setValue(novoAtivo);
        SpreadsheetApp.flush();
        return { status: 'ok', id: body.id, ativo: novoAtivo };
      }
    }
    return _erro('ITEM_NAO_ENCONTRADO', 'Item não encontrado: ' + body.id);
  } finally { lock.releaseLock(); }
}


// =============================================================================
// SPRINT 6 — ENDPOINTS HISTÓRICO
// =============================================================================

/**
 * Endpoint unificado. Aceita params:
 *   tipo: 'os' | 'itens' | 'log' (obrigatório)
 *   Filtros opcionais: id_cliente, cliente, numero_os, id_os, fornecedor,
 *     produto, tipo_acessorio, id_item_catalogo, data_inicio, data_fim,
 *     responsavel_nome, responsavel_id, status_final, teve_faturamento
 */
function getHistorico(params) {
  var tipo = params.tipo || 'os';
  if (tipo === 'os')    return getHistoricoOS(params);
  if (tipo === 'itens') return getHistoricoItens(params);
  if (tipo === 'log')   return getLogOperacional(params);
  return _erro('TIPO_INVALIDO', 'tipo deve ser os, itens ou log');
}

function getHistoricoOS(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Lê aba HISTORICO_OS + aba OS (inclui ativas)
  var resultado = [];

  // Parte 1: OS encerradas na aba HISTORICO_OS
  var abaH = ss.getSheetByName('HISTORICO_OS');
  if (abaH && abaH.getLastRow() > 1) {
    var dadosH = abaH.getDataRange().getValues();
    var cabH = dadosH[0];
    for (var i = 1; i < dadosH.length; i++) {
      var obj = _linhaParaObjeto(cabH, dadosH[i]);
      if (_filtrarOS(obj, params)) resultado.push(obj);
    }
  }

  // Parte 2: OS ativas que ainda não têm registro no HISTORICO_OS
  var idsJaRegistrados = resultado.map(function(r) { return String(r.id_os); });
  var abaOS = ss.getSheetByName('OS');
  var dadosOS = abaOS.getDataRange().getValues();
  var cabOS = dadosOS[0];
  for (var j = 1; j < dadosOS.length; j++) {
    var o = _linhaParaObjeto(cabOS, dadosOS[j]);
    if (idsJaRegistrados.indexOf(String(o.id_os)) >= 0) continue;
    // Montar objeto compatível
    var oH = {
      id_historico_os: '',
      id_os: o.id_os,
      numero_os: o.numero_os || o.id_os,
      id_cliente: o.id_cliente,
      cliente: o.cliente,
      data_abertura: o.data_abertura,
      data_encerramento: '',
      status_final_os: o.status,
      fase_final: o.fase_atual,
      encerrada_por: '',
      origem_encerramento: '',
      motivo_encerramento: '',
      teve_faturamento: 'pendente',
      motivo_sem_faturamento: '',
      valor_total_estimado: '',
      valor_total_aprovado: '',
      total_itens: '',
      total_itens_concluidos: '',
      total_itens_reprovados: '',
      total_itens_pendentes: '',
      observacao_encerramento: '',
      timestamp_registro: ''
    };
    if (_filtrarOS(oH, params)) resultado.push(oH);
  }

  resultado.sort(function(a, b) {
    return String(b.data_abertura || '').localeCompare(String(a.data_abertura || ''));
  });
  return { status: 'ok', historico_os: resultado, total: resultado.length };
}

function _filtrarOS(obj, params) {
  if (params.id_os       && String(obj.id_os)     !== String(params.id_os))       return false;
  if (params.numero_os   && String(obj.numero_os) !== String(params.numero_os))   return false;
  if (params.id_cliente  && String(obj.id_cliente)!== String(params.id_cliente))  return false;
  if (params.cliente     && String(obj.cliente  || '').toLowerCase().indexOf(String(params.cliente).toLowerCase()) < 0) return false;
  if (params.status_final && String(obj.status_final_os).toLowerCase().indexOf(String(params.status_final).toLowerCase()) < 0) return false;
  if (params.teve_faturamento && String(obj.teve_faturamento) !== String(params.teve_faturamento)) return false;
  if (params.data_inicio && obj.data_abertura && String(obj.data_abertura) < String(params.data_inicio)) return false;
  if (params.data_fim    && obj.data_abertura && String(obj.data_abertura) > String(params.data_fim))    return false;
  return true;
}

function getHistoricoItens(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('ITENS').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (!_filtrarItem(obj, params)) continue;
    if (obj.componentes_kit) obj.componentes_kit = _parseJSON(obj.componentes_kit);
    lista.push(obj);
  }
  lista.sort(function(a, b) {
    return String(b.timestamp_atualizacao || '').localeCompare(String(a.timestamp_atualizacao || ''));
  });
  return { status: 'ok', itens: lista, total: lista.length };
}

function _filtrarItem(obj, params) {
  if (params.id_os          && String(obj.id_os)        !== String(params.id_os))       return false;
  if (params.numero_os      && String(obj.id_os)        !== String(params.numero_os))    return false;
  if (params.id_cliente     && String(obj.id_cliente)   !== String(params.id_cliente))   return false;
  if (params.cliente        && String(obj.cliente || '').toLowerCase().indexOf(String(params.cliente).toLowerCase()) < 0) return false;
  if (params.fornecedor     && String(obj.fornecedor || '').toLowerCase().indexOf(String(params.fornecedor).toLowerCase()) < 0) return false;
  if (params.produto        && (String(obj.descricao_curta || '') + ' ' + String(obj.tipo_acessorio || '')).toLowerCase().indexOf(String(params.produto).toLowerCase()) < 0) return false;
  if (params.tipo_acessorio && String(obj.tipo_acessorio || '').toLowerCase().indexOf(String(params.tipo_acessorio).toLowerCase()) < 0) return false;
  if (params.id_item_catalogo && String(obj.id_item_catalogo) !== String(params.id_item_catalogo)) return false;
  if (params.status_final   && String(obj.status).toLowerCase().indexOf(String(params.status_final).toLowerCase()) < 0) return false;
  if (params.responsavel_nome && String(obj.responsavel_ultimo_lancamento || '').toLowerCase().indexOf(String(params.responsavel_nome).toLowerCase()) < 0) return false;
  if (params.data_inicio    && obj.timestamp_criacao && String(obj.timestamp_criacao).substring(0,10) < String(params.data_inicio)) return false;
  if (params.data_fim       && obj.timestamp_criacao && String(obj.timestamp_criacao).substring(0,10) > String(params.data_fim))    return false;
  return true;
}

function getLogOperacional(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('LOG_OPERACIONAL');
  if (!aba || aba.getLastRow() <= 1) return { status: 'ok', eventos: [], total: 0 };
  var dados = aba.getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (params.id_os        && String(obj.id_os)        !== String(params.id_os))        continue;
    if (params.id_item      && String(obj.id_item)      !== String(params.id_item))       continue;
    if (params.id_cliente   && String(obj.id_cliente)   !== String(params.id_cliente))    continue;
    if (params.tipo_evento  && String(obj.tipo_evento)  !== String(params.tipo_evento))   continue;
    if (params.data_inicio  && String(obj.timestamp).substring(0,10) < String(params.data_inicio)) continue;
    if (params.data_fim     && String(obj.timestamp).substring(0,10) > String(params.data_fim))    continue;
    lista.push(obj);
  }
  lista.sort(function(a, b) { return String(b.timestamp).localeCompare(String(a.timestamp)); });
  return { status: 'ok', eventos: lista, total: lista.length };
}

/**
 * encerrarOSHistorico: encerra OS, grava HISTORICO_OS, grava LOG_OPERACIONAL.
 * Não apaga nada — apenas muda status.
 */
function encerrarOSHistorico(body) {
  if (!body.id_os) return _erro('CAMPO_OBRIGATORIO', 'id_os é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    // 1. Atualiza aba OS
    var abaOS = ss.getSheetByName('OS');
    var dadosOS = abaOS.getDataRange().getValues();
    var cabOS = dadosOS[0];
    var colFase   = cabOS.indexOf('fase_atual');
    var colStatus = cabOS.indexOf('status');
    var osObj = null;
    for (var i = 1; i < dadosOS.length; i++) {
      if (String(dadosOS[i][0]) === String(body.id_os)) {
        abaOS.getRange(i+1, colFase+1).setValue('encerrada');
        abaOS.getRange(i+1, colStatus+1).setValue('Encerrada');
        osObj = _linhaParaObjeto(cabOS, dadosOS[i]);
        break;
      }
    }
    if (!osObj) return _erro('ITEM_NAO_ENCONTRADO', 'OS não encontrada: ' + body.id_os);

    // 2. Calcula totais de itens
    var dadosItens = ss.getSheetByName('ITENS').getDataRange().getValues();
    var cabItens = dadosItens[0];
    var colIdOS  = cabItens.indexOf('id_os');
    var colStatusI = cabItens.indexOf('status');
    var totalItens = 0, concluidos = 0, reprovados = 0, pendentes = 0;
    var valorEstimado = 0, valorAprovado = 0;
    var colValor = cabItens.indexOf('valor_cotado');
    var colAprov = cabItens.indexOf('status_aprovacao_cliente');
    var CONCLUIDOS = ['instalado', 'instalado_campo', 'entregue', 'calibrado'];
    var REPROVADOS = ['reprovado_cliente'];
    for (var j = 1; j < dadosItens.length; j++) {
      if (String(dadosItens[j][colIdOS]) !== String(body.id_os)) continue;
      totalItens++;
      var st = String(dadosItens[j][colStatusI]);
      if (CONCLUIDOS.indexOf(st) >= 0) concluidos++;
      else if (REPROVADOS.indexOf(st) >= 0) reprovados++;
      else pendentes++;
      var v = parseFloat(String(dadosItens[j][colValor]).replace(',','.')) || 0;
      valorEstimado += v;
      if (String(dadosItens[j][colAprov]) === 'aprovado') valorAprovado += v;
    }

    // 3. Grava em HISTORICO_OS
    var agora = new Date().toISOString();
    var idHist = _gerarIdHistoricoOS(ss);
    ss.getSheetByName('HISTORICO_OS').appendRow([
      idHist,
      body.id_os,
      osObj.numero_os || body.id_os,
      osObj.id_cliente,
      osObj.cliente,
      osObj.data_abertura || '',
      agora.substring(0,10),
      'Encerrada',
      'encerrada',
      body.encerrada_por || 'adm',
      body.origem_encerramento || 'TapControl Manager',
      body.motivo_encerramento || '',
      body.teve_faturamento || 'pendente',
      body.motivo_sem_faturamento || '',
      valorEstimado.toFixed(2),
      valorAprovado.toFixed(2),
      totalItens, concluidos, reprovados, pendentes,
      body.observacao_encerramento || '',
      agora
    ]);

    SpreadsheetApp.flush();

    // Correção S10 — limpar OS Atual dos equipamentos ao encerrar OS (preserva histórico)
    _limparVinculosEquipamentosDaOS(ss, body.id_os, body.encerrada_por || 'adm');

    // 4. Loga evento
    registrarEventoOperacional({
      tipo_evento: 'ENCERRAMENTO_OS',
      id_os: body.id_os,
      numero_os: osObj.numero_os || body.id_os,
      id_cliente: osObj.id_cliente,
      cliente: osObj.cliente,
      acao_realizada: 'OS encerrada',
      responsavel_nome: body.encerrada_por || 'adm',
      origem: body.origem_encerramento || 'TapControl Manager',
      observacao: body.observacao_encerramento || '',
      payload_json: JSON.stringify({ teve_faturamento: body.teve_faturamento, total_itens: totalItens, concluidos: concluidos, reprovados: reprovados })
    });

    return { status: 'ok', id_os: body.id_os, id_historico_os: idHist, total_itens: totalItens };
  } finally { lock.releaseLock(); }
}


// =============================================================================
// SPRINT S10 — VÍNCULO ATIVO EQUIPAMENTO / OS
// =============================================================================

/**
 * vincularEquipamentoOS — endpoint explícito para vincular ou desvincular.
 * POST body: { id_equipamento (obrig), id_os (string ou ""), responsavel?, observacao? }
 * Atualiza APENAS EQUIPAMENTOS.id_os. Registra LOG_OPERACIONAL.
 */
function vincularEquipamentoOS(body) {
  if (!body.id_equipamento) return _erro('CAMPO_OBRIGATORIO', 'id_equipamento é obrigatório.');
  // id_os pode ser "" para desvincular — aceitar undefined como erro para forçar envio explícito
  if (body.id_os === undefined) return _erro('CAMPO_OBRIGATORIO', 'id_os é obrigatório (use "" para desvincular).');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName('EQUIPAMENTOS');
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var colIdOS = cab.indexOf('id_os');
    if (colIdOS < 0) return _erro('COLUNA_NAO_ENCONTRADA', 'Coluna id_os não encontrada em EQUIPAMENTOS.');

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id_equipamento)) {
        var idOSAnterior = String(dados[i][colIdOS] || '');
        var idOSNova = body.id_os || '';

        // Valida se OS existe (quando vinculando)
        if (idOSNova) {
          var dadosOS = ss.getSheetByName('OS').getDataRange().getValues();
          var osEncontrada = false;
          for (var j = 1; j < dadosOS.length; j++) {
            if (String(dadosOS[j][0]) === String(idOSNova)) { osEncontrada = true; break; }
          }
          if (!osEncontrada) return _erro('ITEM_NAO_ENCONTRADO', 'OS não encontrada: ' + idOSNova);
        }

        aba.getRange(i + 1, colIdOS + 1).setValue(idOSNova);
        SpreadsheetApp.flush();

        var tipoEvento = idOSNova ? 'VINCULO_EQUIPAMENTO_OS' : 'DESVINCULO_EQUIPAMENTO_OS';
        registrarEventoOperacional({
          tipo_evento: tipoEvento,
          id_os: idOSNova || idOSAnterior,
          id_equipamento: body.id_equipamento,
          tag_equipamento: String(dados[i][cab.indexOf('tag')] || ''),
          acao_realizada: idOSNova ? 'Equipamento vinculado à OS' : 'Equipamento desvinculado da OS',
          responsavel_nome: body.responsavel || 'adm',
          origem: body.origem || 'TapControl Manager',
          observacao: body.observacao || '',
          payload_json: JSON.stringify({ id_os_anterior: idOSAnterior, id_os_nova: idOSNova })
        });

        return { status: 'ok', id_equipamento: body.id_equipamento, id_os: idOSNova, modo: idOSNova ? 'vinculado' : 'desvinculado' };
      }
    }
    return _erro('ITEM_NAO_ENCONTRADO', 'Equipamento não encontrado: ' + body.id_equipamento);
  } finally { lock.releaseLock(); }
}

/**
 * _limparVinculosEquipamentosDaOS — chamada no encerramento de OS.
 * Limpa EQUIPAMENTOS.id_os de todos os equipamentos vinculados à OS encerrada.
 * Preserva todo o histórico (HISTORICO_OS, LOG_OPERACIONAL, LEVANTAMENTOS, ITENS).
 */
function _limparVinculosEquipamentosDaOS(ss, id_os, responsavel) {
  try {
    var aba = ss.getSheetByName('EQUIPAMENTOS');
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var colIdOS = cab.indexOf('id_os');
    if (colIdOS < 0) return;
    var limpos = [];
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][colIdOS] || '') === String(id_os)) {
        aba.getRange(i + 1, colIdOS + 1).setValue('');
        limpos.push(String(dados[i][0]));
        registrarEventoOperacional({
          tipo_evento: 'DESVINCULO_EQUIPAMENTO_OS',
          id_os: id_os,
          id_equipamento: String(dados[i][0]),
          tag_equipamento: String(dados[i][cab.indexOf('tag')] || ''),
          acao_realizada: 'Vínculo limpo no encerramento da OS',
          responsavel_nome: responsavel || 'adm',
          origem: 'TapControl Manager',
          payload_json: JSON.stringify({ motivo: 'encerramento_os' })
        });
      }
    }
    if (limpos.length) SpreadsheetApp.flush();
    Logger.log('_limparVinculosEquipamentosDaOS: ' + limpos.length + ' equipamento(s) desvinculados da OS ' + id_os);
  } catch(e) {
    Logger.log('_limparVinculosEquipamentosDaOS ERRO: ' + e.message);
  }
}
// =============================================================================

/**
 * registrarEventoOperacional(evento)
 * Chamado por todos os endpoints de escrita para garantir rastreabilidade.
 * Campos do objeto evento (todos opcionais exceto tipo_evento):
 *   tipo_evento, id_os, numero_os, id_cliente, cliente,
 *   id_item, id_equipamento, tag_equipamento,
 *   produto, tipo_acessorio,
 *   status_anterior, status_novo, acao_realizada,
 *   responsavel_id, responsavel_nome, responsavel_tipo,
 *   origem, fornecedor, valor, observacao, payload_json
 */
function registrarEventoOperacional(evento) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName('LOG_OPERACIONAL');
    if (!aba) return; // segurança: não quebra fluxo se aba não existe
    var agora = new Date().toISOString();
    var idEvento = _gerarIdEvento(ss);
    aba.appendRow([
      idEvento,
      agora,
      evento.tipo_evento     || '',
      evento.id_os           || '',
      evento.numero_os       || '',
      evento.id_cliente      || '',
      evento.cliente         || '',
      evento.id_item         || '',
      evento.id_equipamento  || '',
      evento.tag_equipamento || '',
      evento.produto         || '',
      evento.tipo_acessorio  || '',
      evento.status_anterior || '',
      evento.status_novo     || '',
      evento.acao_realizada  || '',
      evento.responsavel_id  || '',
      evento.responsavel_nome|| '',
      evento.responsavel_tipo|| '',
      evento.origem          || '',
      evento.fornecedor      || '',
      evento.valor           || '',
      evento.observacao      || '',
      evento.payload_json    || ''
    ]);
  } catch(e) {
    // Nunca quebra o fluxo principal
    Logger.log('registrarEventoOperacional ERRO: ' + e.message);
  }
}

function _gerarIdEvento(ss) {
  var aba = ss.getSheetByName('LOG_OPERACIONAL');
  var total = Math.max((aba ? aba.getLastRow() - 1 : 0), 0);
  return 'EVT-' + String(total + 1).padStart(6, '0');
}

function _gerarIdHistoricoOS(ss) {
  var aba = ss.getSheetByName('HISTORICO_OS');
  var total = Math.max((aba ? aba.getLastRow() - 1 : 0), 0);
  return 'HOS-' + String(total + 1).padStart(5, '0');
}


// =============================================================================
// TÉCNICOS DE CAMPO (INSPETORES)
// =============================================================================

function getInspetores(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('CAT_INSPETORES').getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    // Nunca expõe o pin_hash
    delete obj.pin_hash;
    lista.push(obj);
  }
  return { status: 'ok', inspetores: lista };
}

function salvarInspetor(body) {
  if (!body.nome) return _erro('CAMPO_OBRIGATORIO', 'nome é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('CAT_INSPETORES');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (body.id) {
      // Atualizar
      var dados = aba.getDataRange().getValues();
      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(body.id)) {
          aba.getRange(i+1, 2).setValue(body.nome);
          if (body.pin) aba.getRange(i+1, 3).setValue(_hashPin(body.pin));
          SpreadsheetApp.flush();
          return { status: 'ok', modo: 'atualizado', id: body.id };
        }
      }
      return _erro('ITEM_NAO_ENCONTRADO', 'Inspetor não encontrado: ' + body.id);
    } else {
      if (!body.pin) return _erro('CAMPO_OBRIGATORIO', 'pin é obrigatório para novo inspetor.');
      var id = _gerarIdCatalogo(ss, 'CAT_INSPETORES', 'INS');
      aba.appendRow([id, body.nome, _hashPin(String(body.pin)), 'true']);
      SpreadsheetApp.flush();
      return { status: 'ok', modo: 'criado', id: id };
    }
  } finally { lock.releaseLock(); }
}

function toggleAtivoInspetor(body) {
  if (!body.id) return _erro('CAMPO_OBRIGATORIO', 'id é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('CAT_INSPETORES');
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var dados = aba.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id)) {
        var novoAtivo = String(body.ativo) === 'true' ? 'true' : 'false';
        aba.getRange(i+1, 4).setValue(novoAtivo);
        SpreadsheetApp.flush();
        return { status: 'ok', id: body.id, ativo: novoAtivo };
      }
    }
    return _erro('ITEM_NAO_ENCONTRADO', 'Inspetor não encontrado: ' + body.id);
  } finally { lock.releaseLock(); }
}


// =============================================================================
// FUNÇÕES AUXILIARES
// =============================================================================

function _gerarId(ss, nomeAba, prefixo) {
  var ano = new Date().getFullYear();
  var aba = ss.getSheetByName(nomeAba);
  var ultimo = 0;
  var padrao = new RegExp('^' + prefixo + '-' + ano + '-(\\d+)$');
  if (aba.getLastRow() > 1) {
    var ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
    ids.forEach(function(row) {
      var m = String(row[0]).match(padrao);
      if (m) { var n = parseInt(m[1], 10); if (n > ultimo) ultimo = n; }
    });
  }
  return prefixo + '-' + ano + '-' + String(ultimo + 1).padStart(3, '0');
}

function _gerarIdCatalogo(ss, nomeAba, prefixo) {
  var aba = ss.getSheetByName(nomeAba);
  var ultimo = 0;
  var padrao = new RegExp('^' + prefixo + '-(\\d+)$');
  if (aba.getLastRow() > 1) {
    var ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
    ids.forEach(function(row) {
      var m = String(row[0]).match(padrao);
      if (m) { var n = parseInt(m[1], 10); if (n > ultimo) ultimo = n; }
    });
  }
  return prefixo + '-' + String(ultimo + 1).padStart(3, '0');
}

function _gerarIdHistorico(ss) {
  var aba   = ss.getSheetByName('HISTORICO_STATUS');
  var total = Math.max(aba.getLastRow() - 1, 0);
  return 'HST-' + String(total + 1).padStart(5, '0');
}

function _hashPin(pin) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin));
  return bytes.map(function(b) { return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2); }).join('');
}

function _lerCatalogo(ss, nomeAba) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) return [];
  var dados = aba.getDataRange().getValues();
  var cab = dados[0];
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var obj = _linhaParaObjeto(cab, dados[i]);
    if (String(obj.ativo).toLowerCase() === 'true') {
      if (obj.componentes) obj.componentes = _parseJSON(obj.componentes);
      lista.push(obj);
    }
  }
  return lista;
}

function _lerCatalogoDCBI(ss) {
  // DCBI usa campo 'componentes' como JSON
  return _lerCatalogo(ss, 'CAT_DCBI');
}

function _linhaParaObjeto(cab, linha) {
  var obj = {};
  for (var i = 0; i < cab.length; i++) obj[cab[i]] = linha[i];
  return obj;
}

function _contarEquipamentosDaOS(ss, idOS) {
  // Correção S12 — localiza colunas id_os e ativo pelo cabeçalho, não por índice fixo
  var aba = ss.getSheetByName('EQUIPAMENTOS');
  var dados = aba.getDataRange().getValues();
  var cab = dados[0];
  var colIdOS  = cab.indexOf('id_os');
  var colAtivo = cab.indexOf('ativo');
  if (colIdOS < 0 || colAtivo < 0) return 0; // segurança: retorna 0 se estrutura inválida
  var count = 0;
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][colIdOS]) === String(idOS) && String(dados[i][colAtivo]).toLowerCase() === 'true') count++;
  }
  return count;
}

function _buscarNomeInspetor(ss, id) {
  var dados = ss.getSheetByName('CAT_INSPETORES').getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) { if (dados[i][0] === id) return dados[i][1]; }
  return 'Inspetor não identificado';
}

function _buscarNomeCliente(ss, id) {
  var dados = ss.getSheetByName('CLIENTES').getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) { if (String(dados[i][0]) === String(id)) return dados[i][1]; }
  return 'Cliente não identificado';
}

function _buscarClienteDaOS(ss, idOS) {
  var dados = ss.getSheetByName('OS').getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) { if (String(dados[i][0]) === String(idOS)) return dados[i][1]; }
  return '';
}

function _resumoStatusDoLevantamento(ss, idLev) {
  var dados = ss.getSheetByName('ITENS').getDataRange().getValues();
  var cab = dados[0];
  var colLev    = cab.indexOf('id_levantamento');
  var colStatus = cab.indexOf('status');
  var resumo = {};
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][colLev]) === String(idLev)) {
      var s = dados[i][colStatus] || 'sem_status';
      resumo[s] = (resumo[s] || 0) + 1;
    }
  }
  return resumo;
}

function _gerarDescCurtaMano(d) {
  var partes = [];
  var p1 = [d.tipo, d.fluido, d.caixa_mm ? d.caixa_mm + 'mm' : ''].filter(Boolean).join(' ');
  if (p1.trim()) partes.push(p1.trim());
  var p2 = [d.rosca, d.tipo_rosca].filter(Boolean).join(' ');
  if (p2.trim()) partes.push(p2.trim());
  if (d.entrada)      partes.push(String(d.entrada).trim());
  if (d.escala_padrao) partes.push(String(d.escala_padrao).trim());
  if (d.marca)        partes.push(String(d.marca).trim());
  return partes.join(' · ');
}

function _parseJSON(v) {
  if (!v || typeof v !== 'string') return v;
  try { return JSON.parse(v); } catch(e) { return v; }
}

function _resposta(dados) {
  return ContentService.createTextOutput(JSON.stringify(dados)).setMimeType(ContentService.MimeType.JSON);
}

function _erro(codigo, mensagem) {
  return { status: 'erro', codigo: codigo, mensagem: mensagem };
}

// [DUPLICADO] S17 — helpers para bloqueio de cadastros duplicados
function _normalizarTexto(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')  // remove acentos
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')                       // remove pontuação
    .replace(/\s+/g, ' ')                               // espaços únicos
    .trim();
}

function _erroDuplicado(mensagem, existente) {
  return { status: 'erro', codigo: 'DUPLICADO', mensagem: mensagem, existente: existente || null };
}


function _atualizarStatusFaseOS(ss, idOS, statusNovo, faseNova, idTecnico){
  var abaOS = ss.getSheetByName('OS');
  var dados = abaOS.getDataRange().getValues();
  var cab = dados[0], cId=cab.indexOf('id_os'), cStatus=cab.indexOf('status'), cFase=cab.indexOf('fase_atual'), cTec=cab.indexOf('tecnicos_vinculados');
  for (var i=1;i<dados.length;i++){
    if (String(dados[i][cId])===String(idOS)){
      if(cStatus>=0) abaOS.getRange(i+1,cStatus+1).setValue(statusNovo);
      if(cFase>=0) abaOS.getRange(i+1,cFase+1).setValue(faseNova);
      if(idTecnico && cTec>=0) abaOS.getRange(i+1,cTec+1).setValue(JSON.stringify([String(idTecnico)]));
      return { numero_os: dados[i][cab.indexOf('numero_os')] || idOS };
    }
  }
  throw new Error('OS não encontrada: '+idOS);
}
function cancelarEnvioLevantamentoFieldTap(body){
  if(!body.id_os||!body.id_inspetor) return _erro('CAMPO_OBRIGATORIO','id_os e id_inspetor são obrigatórios.');
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var nome=_buscarNomeInspetor(ss, body.id_inspetor);
  var os=_atualizarStatusFaseOS(ss, body.id_os, 'Fase 1', '1');

  // [CANCELAMENTO_ENVIO] BUG1-FIX: marcar levantamento mais recente desta OS/Fase1
  // como em_correcao_campo para que reenvio possa reutilizá-lo
  var abaLev = ss.getSheetByName('LEVANTAMENTOS');
  var dadosLev = abaLev.getDataRange().getValues();
  var cabLev = dadosLev[0];
  var colLevOS     = cabLev.indexOf('id_os');
  var colLevFase   = cabLev.indexOf('fase');
  var colLevStatus = cabLev.indexOf('status_geral');
  var colLevTs     = cabLev.indexOf('timestamp_envio');
  var melhorLinha = -1;
  var melhorTs = '';
  for (var lv = 1; lv < dadosLev.length; lv++) {
    if (String(dadosLev[lv][colLevOS] || '') === String(body.id_os) &&
        String(dadosLev[lv][colLevFase] || '') === '1') {
      var ts = String(dadosLev[lv][colLevTs] || '');
      if (ts > melhorTs || melhorLinha === -1) { melhorTs = ts; melhorLinha = lv; }
    }
  }
  var levId = '';
  if (melhorLinha >= 0 && colLevStatus >= 0) {
    abaLev.getRange(melhorLinha + 1, colLevStatus + 1).setValue('em_correcao_campo');
    levId = String(dadosLev[melhorLinha][cabLev.indexOf('id_levantamento')] || '');
    SpreadsheetApp.flush();
    Logger.log('[CANCELAMENTO_ENVIO] Levantamento ' + levId + ' marcado como em_correcao_campo.');
  }

  registrarEventoOperacional({tipo_evento:'CANCELAMENTO_ENVIO_FASE1_FIELDTP',id_os:body.id_os,numero_os:os.numero_os,responsavel_id:body.id_inspetor,responsavel_nome:nome,responsavel_tipo:'inspetor',origem:body.origem||'TapControl Field',observacao:body.motivo||'Cancelamento para correção em campo'});
  return {status:'ok', id_levantamento: levId};
}
function cancelarFechamentoFase2FieldTap(body){
  if(!body.id_os||!body.id_inspetor) return _erro('CAMPO_OBRIGATORIO','id_os e id_inspetor são obrigatórios.');
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var nome=_buscarNomeInspetor(ss, body.id_inspetor);
  var os=_atualizarStatusFaseOS(ss, body.id_os, 'Fase 2', '2');
  registrarEventoOperacional({tipo_evento:'CANCELAMENTO_FECHAMENTO_FASE2_FIELDTP',id_os:body.id_os,numero_os:os.numero_os,responsavel_id:body.id_inspetor,responsavel_nome:nome,responsavel_tipo:'inspetor',origem:body.origem||'TapControl Field',observacao:body.motivo||'Cancelamento de fechamento para correção'});
  return {status:'ok'};
}
function reabrirOSParaFieldTap(body){
  if(!body.id_os||!body.fase_destino||!body.motivo) return _erro('CAMPO_OBRIGATORIO','id_os, fase_destino e motivo são obrigatórios.');
  if(String(body.fase_destino)==='2' && !body.id_tecnico) return _erro('CAMPO_OBRIGATORIO','id_tecnico obrigatório para Fase 2.');
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var st=String(body.fase_destino)==='2'?'Fase 2':'Fase 1';
  var os=_atualizarStatusFaseOS(ss, body.id_os, st, String(body.fase_destino), body.id_tecnico);
  registrarEventoOperacional({tipo_evento:'REABERTURA_OS_FIELDTP_TAPPARTS',id_os:body.id_os,numero_os:os.numero_os,responsavel_nome:body.responsavel||'TapControl Manager',responsavel_tipo:'adm',origem:'TapControl Manager',observacao:body.motivo});
  return {status:'ok'};
}
function reenviarOSFieldTap(body){
  if(!body.id_os||!body.fase||!body.motivo) return _erro('CAMPO_OBRIGATORIO','id_os, fase e motivo são obrigatórios.');
  if(String(body.fase)==='2' && !body.id_tecnico) return _erro('CAMPO_OBRIGATORIO','id_tecnico obrigatório para Fase 2.');
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var st=String(body.fase)==='2'?'Fase 2':'Fase 1';
  var os=_atualizarStatusFaseOS(ss, body.id_os, st, String(body.fase), body.id_tecnico);
  registrarEventoOperacional({tipo_evento:'REENVIO_OS_FIELDTP_TAPPARTS',id_os:body.id_os,numero_os:os.numero_os,responsavel_nome:body.responsavel||'TapControl Manager',responsavel_tipo:'adm',origem:'TapControl Manager',observacao:body.motivo});
  return {status:'ok'};
}

// =============================================================================
// S18 — NOVOS ENDPOINTS
// =============================================================================

// S18 Correção 11 — reabrirOS (compatível com TCM que chama 'reabrirOS' simples)
function reabrirOS(body) {
  if (!body.id_os) return _erro('CAMPO_OBRIGATORIO', 'id_os é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var os = _atualizarStatusFaseOS(ss, body.id_os, 'Fase 2 Campo Concluída', 'campo_concluido');
    SpreadsheetApp.flush();
    registrarEventoOperacional({ tipo_evento: 'REABERTURA_OS', id_os: body.id_os, numero_os: os.numero_os, acao_realizada: 'OS reaberta para campo_concluido', responsavel_nome: body.responsavel || 'adm', origem: body.origem || 'TapControl Manager' });
    return { status: 'ok', id_os: body.id_os, fase_atual: 'campo_concluido' };
  } catch(e) {
    return _erro('ERRO_INTERNO', e.message);
  } finally { lock.releaseLock(); }
}

// S18 Correção 8 — registrarLacre separado de calibração
function registrarLacre(body) {
  if (!body.id_item) return _erro('CAMPO_OBRIGATORIO', 'id_item é obrigatório.');
  if (!body.num_lacre) return _erro('CAMPO_OBRIGATORIO', 'num_lacre é obrigatório.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName('ITENS');
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var colNumLacre = cab.indexOf('num_lacre');
    var colTs       = cab.indexOf('timestamp_atualizacao');
    var found = false;
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(body.id_item)) {
        var agora = new Date().toISOString();
        if (colNumLacre >= 0) aba.getRange(i+1, colNumLacre+1).setValue(body.num_lacre);
        if (colTs >= 0)       aba.getRange(i+1, colTs+1).setValue(agora);
        // Não altera status, data_calibracao nem validade_calibracao

        // Grava em HISTORICO_LACRES
        var idLacre = 'LAC-' + agora.replace(/[^0-9]/g,'').substring(0,14);
        var abaHL = ss.getSheetByName('HISTORICO_LACRES');
        if (abaHL) {
          abaHL.appendRow([
            idLacre, body.id_item,
            dados[i][cab.indexOf('id_os')] || '',
            dados[i][cab.indexOf('id_equipamento')] || '',
            dados[i][cab.indexOf('tag_equipamento')] || '',
            body.num_lacre,
            body.data_lacre || agora.substring(0,10),
            body.responsavel || '',
            body.observacao || '',
            agora,
            body.origem || 'TapControl Manager'
          ]);
        }
        SpreadsheetApp.flush();
        registrarEventoOperacional({ tipo_evento: 'REGISTRO_LACRE', id_os: dados[i][cab.indexOf('id_os')] || '', id_item: body.id_item, produto: dados[i][cab.indexOf('descricao_curta')] || '', tipo_acessorio: dados[i][cab.indexOf('tipo_acessorio')] || '', acao_realizada: 'Lacre avulso registrado', responsavel_nome: body.responsavel || '', origem: body.origem || 'TapControl Manager', observacao: 'Lacre: ' + body.num_lacre });
        found = true;
        return { status: 'ok', id_item: body.id_item, id_lacre: idLacre, timestamp: agora };
      }
    }
    if (!found) return _erro('ITEM_NAO_ENCONTRADO', 'Item não encontrado: ' + body.id_item);
  } finally { lock.releaseLock(); }
}

// S18 Correção 3 — rascunho permanente no GAS (FIELD_DRAFTS)
function salvarDraftFieldTap(body) {
  if (!body.id_os || !body.id_inspetor) return _erro('CAMPO_OBRIGATORIO', 'id_os e id_inspetor são obrigatórios.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var aba = ss.getSheetByName('FIELD_DRAFTS');
    if (!aba) return _erro('ABA_NAO_ENCONTRADA', 'FIELD_DRAFTS não existe. Execute setupPlanilha().');
    var dados = aba.getDataRange().getValues();
    var cab = dados[0];
    var agora = new Date().toISOString();
    var fase = String(body.fase || '1');
    var totalItens = body.total_itens || 0;

    // Upsert: critério id_os + id_inspetor + fase com status_draft = 'aberto'
    var colOS   = cab.indexOf('id_os');
    var colInsp = cab.indexOf('id_inspetor');
    var colFase = cab.indexOf('fase');
    var colSt   = cab.indexOf('status_draft');
    var colTs   = cab.indexOf('timestamp_atualizacao');
    var colPayl = cab.indexOf('payload_json');
    var colTot  = cab.indexOf('total_itens');
    var colId   = cab.indexOf('id_draft');

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][colOS] || '') === String(body.id_os) &&
          String(dados[i][colInsp] || '') === String(body.id_inspetor) &&
          String(dados[i][colFase] || '') === fase &&
          String(dados[i][colSt] || '') === 'aberto') {
        // Atualizar linha existente
        var idDraft = String(dados[i][colId] || '');
        if (colTs >= 0)   aba.getRange(i+1, colTs+1).setValue(agora);
        if (colPayl >= 0) aba.getRange(i+1, colPayl+1).setValue(body.payload_json || '');
        if (colTot >= 0)  aba.getRange(i+1, colTot+1).setValue(totalItens);
        SpreadsheetApp.flush();
        return { status: 'ok', id_draft: idDraft, modo: 'atualizado' };
      }
    }

    // Inserir nova linha
    var novoId = 'DFT-' + agora.replace(/[^0-9]/g,'').substring(0,14);
    var linhaMap = {
      id_draft: novoId, id_os: body.id_os, numero_os: body.numero_os || '',
      id_inspetor: body.id_inspetor, nome_inspetor: body.nome_inspetor || '',
      fase: fase, timestamp_atualizacao: agora, total_itens: totalItens,
      status_draft: 'aberto', payload_json: body.payload_json || ''
    };
    aba.appendRow(cab.map(function(c) { return linhaMap[c] !== undefined ? linhaMap[c] : ''; }));
    SpreadsheetApp.flush();
    return { status: 'ok', id_draft: novoId, modo: 'criado' };
  } finally { lock.releaseLock(); }
}

function getDraftFieldTap(body) {
  if (!body.id_os || !body.id_inspetor) return _erro('CAMPO_OBRIGATORIO', 'id_os e id_inspetor são obrigatórios.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('FIELD_DRAFTS');
  if (!aba || aba.getLastRow() <= 1) return { status: 'ok', draft: null };
  var dados = aba.getDataRange().getValues();
  var cab = dados[0];
  var fase = String(body.fase || '1');
  var colOS   = cab.indexOf('id_os');
  var colInsp = cab.indexOf('id_inspetor');
  var colFase = cab.indexOf('fase');
  var colSt   = cab.indexOf('status_draft');
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][colOS] || '') === String(body.id_os) &&
        String(dados[i][colInsp] || '') === String(body.id_inspetor) &&
        String(dados[i][colFase] || '') === fase &&
        String(dados[i][colSt] || '') === 'aberto') {
      return { status: 'ok', draft: _linhaParaObjeto(cab, dados[i]) };
    }
  }
  return { status: 'ok', draft: null };
}

function limparDraftFieldTap(body) {
  if (!body.id_os || !body.id_inspetor) return _erro('CAMPO_OBRIGATORIO', 'id_os e id_inspetor são obrigatórios.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('FIELD_DRAFTS');
  if (!aba || aba.getLastRow() <= 1) return { status: 'ok' };
  var dados = aba.getDataRange().getValues();
  var cab = dados[0];
  var fase = String(body.fase || '1');
  var novoStatus = body.motivo === 'cancelado' ? 'cancelado' : 'enviado';
  var colOS   = cab.indexOf('id_os');
  var colInsp = cab.indexOf('id_inspetor');
  var colFase = cab.indexOf('fase');
  var colSt   = cab.indexOf('status_draft');
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][colOS] || '') === String(body.id_os) &&
        String(dados[i][colInsp] || '') === String(body.id_inspetor) &&
        String(dados[i][colFase] || '') === fase &&
        String(dados[i][colSt] || '') === 'aberto') {
      aba.getRange(i+1, colSt+1).setValue(novoStatus);
      SpreadsheetApp.flush();
      return { status: 'ok', status_draft: novoStatus };
    }
  }
  return { status: 'ok', status_draft: 'nao_encontrado' };
}
