# PROMPT — Codex: Ordens de Serviço · FASE 1 COMPLETA (alinhado ao Planejamento Jun/2026)

**Pipeline:** Codex CONSTRÓI → Claude Code AUDITA e CORRIGE contra os arquivos reais. Este prompt é para o **Codex** e será lido a partir do **repositório** `brunoengetap/ProperCare130526` (branch `main`, layout **plano**: arquivos na raiz).

**Escopo desta entrega: FASE 1 COMPLETA** — inclui login PIN, OS numeradas, filtro por técnico, cronômetro, seleção de máquinas, PDF por máquina, **assinatura digital** e **salvamento de fotos/PDF no Google Drive** (itens que o roadmap listava como "v14", trazidos para esta fase a pedido). Fase 2/3 (e-mail, notificação, Calendar) seguem fora — ver "Fora de escopo".

---

## 0. PREMISSAS INVIOLÁVEIS

1. **O sistema atual está em produção e funcionando.** Toda mudança desta fase é **aditiva e paralela**: abas novas no Sheets, ações novas no GAS, aba nova no PCM, telas novas no PCF. **Só cruzar com o existente quando for inevitável** (pontos descritos em 2.5 e 4.x).
2. **O PCF de inspeção/comercial NÃO MUDA.** O fluxo de OS — incl. assinatura e envio de fotos/PDF ao Drive — se acopla **apenas ao modo preventiva** (`_pgpTipoVisita==='preventiva'`). O modo `'inspecao'` e seu fluxo de fotos/PDF continuam exatamente como hoje. Login por PIN é **porta opcional** para o modo OS — nunca portão sobre os fluxos atuais.
3. **A Sheet (`Proper_PGP_DB`) é a única fonte de verdade.** O binário (foto/PDF/assinatura) vive no **Drive**; a planilha guarda **sempre** a referência (`File_ID`+`File_URL`). Nenhuma foto/PDF/assinatura pode existir **apenas** no PDF ou no dispositivo. Numeração reconciliável pela planilha; Script Property só acelera. localStorage/sessionStorage = cache/sessão, nunca verdade.
4. **Schema append-only.** Abas novas são livres; em abas existentes, só acrescentar coluna **ao final**, nunca reordenar.

---

## ⚠️ PASSO ZERO — RECONCILIAR VERSÕES E LER OS ARQUIVOS REAIS

Os nomes/headers citados vêm de inspeção das versões antigas (v17/v22/v10) e **podem ter mudado**. Antes de qualquer edição:

```bash
ls -1 *.html *.js                     # ver versões REAIS no repo
grep -n "const HEADERS" GAS_properCare_*.js
```

- **Alvos esperados pelo plano:** GAS `v19→v20`, PCM `v28→v29`, PCF `v12→v13`.
- **ATENÇÃO:** o repo pode estar com **PCM v27** (não v28). **Não hardcode o número** — descubra o maior número real de cada arquivo no repo e gere `atual+1`. Se a base do PCM no repo (`v27`) divergir do plano (`v28`), **pare e reporte** antes de editar: pode faltar um push.
- Para cada ponto de edição: abra o arquivo real → localize por `grep` → edite seguindo o padrão ao redor. Landmark não encontrado → **pare e reporte**, não invente.

---

## RESTRIÇÕES — NÃO TOCAR

- `machineKey()`, `clientKey()` — nunca modificar.
- `proper_admin_v2` (localStorage) — nunca renomear/reestruturar.
- `pgpFlushPending`, `pgpPostJson`, `pgpBuildVisitRequest`, `pgpBuildPartsUpdateRequest` — não alterar assinaturas; só **chamar** ou **acrescentar**.
- Funções de PDF/foto/inspeção do PCF — estender só por **composição** (parâmetro novo com default que preserva o comportamento atual). Nunca quebrar o fluxo de inspeção/comercial.
- Soft delete apenas (`Ativo = NÃO`).
- Deploy: **GAS → PCM → PCF**. Mudança cirúrgica.

---

## 1. NOVAS ABAS NO SHEETS (append-only; criar via `getOrCreateSheet`)

> Padrão estável confirmado no GAS: objeto `const HEADERS = {...}`; helpers `getOrCreateSheet(name, headers)`, `ensureSheetHeaders(name, expected)`, `getSheetData(name)`, `jsonResponse(data)`; `doPost(e)` faz `body=JSON.parse(e.postData.contents)`, `action=body.action`, `checkKey(body)`, `switch(action)`→`result`→`jsonResponse(result)`; `doGet` tem switch próprio para leituras (`?action=...&key=...`).

Adicionar ao objeto `HEADERS`, sem tocar nas chaves existentes:

```
CAT_TECNICOS: [ 'id','nome','pin_hash','ativo','perfil' ]   // perfil: campo | comercial

ORDENS_SERVICO: [
  'id_os','numero_os','id_cliente','cliente','descricao',
  'data_abertura','data_prevista','status',
  'tecnicos_vinculados','maquinas_vinculadas','id_visita_resultado',
  'drive_folder_id','drive_folder_url',
  'inicio_atendimento','fim_atendimento','pdf_url'
]

TECH_ATTACHMENTS: [
  'Attachment_ID','Entity_Type','Entity_ID','OS_Numero',
  'File_ID','File_URL','File_Name','Mime','Attachment_Type','Caption',
  'Created_At','Created_By','Ativo'
]
```
- `status`: `aberta` | `em_andamento` | `concluida`.
- `tecnicos_vinculados` / `maquinas_vinculadas`: **JSON array** (string).
- `drive_folder_id`/`drive_folder_url`/`pdf_url`: preenchidas pelo GAS ao receber fotos/PDF.
- `inicio_atendimento`/`fim_atendimento`: ISO 8601 string (gravadas pelo PCF).
- `TECH_ATTACHMENTS`: **uma linha por arquivo** (foto/assinatura/PDF) — referência canônica. `Entity_Type`: `os`|`maquina`|`peca`. `Attachment_Type`: `foto`|`foto_peca`|`foto_horimetro`|`assinatura_tecnico`|`assinatura_cliente`|`pdf_os`.

Em **VISITAS** (única coluna nova, o elo inevitável OS↔visita): acrescentar **ao final** do array a coluna `OS_Numero`.

---

## 2. GAS (`v19` → `v20`) — primeiro

**Localizar:**
```bash
grep -n "const HEADERS\|function doPost\|function doGet\|switch *(action)\|case 'saveVisit'\|case 'getVisits'\|function getOrCreateSheet\|function jsonResponse\|function checkKey\|function getSheetData" GAS_properCare_v19.js
```

### 2.1 `loginTecnico` (doPost — recebe hash, nunca PIN em texto)
Novo `case 'loginTecnico'`. Busca `pin_hash` em `CAT_TECNICOS` (`ativo` = SIM), compara com `body.pin_hash`. Retorna `{ status:'ok', tecnico:{id,nome,perfil}, os:[...] }`, onde `os` = OS de `ORDENS_SERVICO` cujo `tecnicos_vinculados` (JSON) contém o `id` e `status != concluida`. Falha → `{ status:'error', error:'PIN inválido' }`. **Não logar o hash.**

### 2.2 `saveOS` (doPost — upsert)
Novo `case 'saveOS'`. Se `body.id_os` vazio → cria: gera `id_os` (`OS-`+timestamp), gera `numero_os` via `generateOsNumber_()`, `data_abertura=now`, `status='aberta'`, grava a linha. Se `id_os` presente → atualiza a linha existente **localizando coluna por NOME** (índice em `HEADERS.ORDENS_SERVICO`), nunca por posição. Campos atualizáveis: `status`, `maquinas_vinculadas`, `tecnicos_vinculados`, `descricao`, `data_prevista`, `id_visita_resultado`, `inicio_atendimento`, `fim_atendimento`. Retorna a OS gravada (incl. `numero_os`).

### 2.3 `getOS` (doGet — leitura)
Novo `case 'getOS'` no switch do `doGet`. Params opcionais: `tecnico` (filtra por id em `tecnicos_vinculados` JSON), `status`, `cliente`, `de`/`ate` (período sobre `data_abertura`). Retorna linhas ativas de `ORDENS_SERVICO`.

### 2.4 Numeração reconciliável pela Sheet
```js
function generateOsNumber_() {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ano = new Date().getFullYear();
    const props = PropertiesService.getScriptProperties();
    const key = 'OS_SEQ_' + ano;
    let maxSheet = 0;                                  // autoridade = a planilha
    const rows = getSheetData('ORDENS_SERVICO');       // adaptar ao formato real (objetos x matriz)
    const re = new RegExp('^PGP-' + ano + '-(\\d+)$');
    (rows || []).forEach(function(r){
      const n = (r && (r['numero_os'] || r.numero_os)) || '';
      const m = re.exec(String(n)); if (m) maxSheet = Math.max(maxSheet, parseInt(m[1],10));
    });
    const next = Math.max(maxSheet, parseInt(props.getProperty(key)||'0',10)) + 1;
    props.setProperty(key, String(next));
    return 'PGP-' + ano + '-' + String(next).padStart(4,'0');
  } finally { lock.releaseLock(); }
}
```

### 2.5 Elo OS↔visita (único cruzamento com o existente)
Em `saveVisit`/`savePreventiva`: **se** `body.os_numero` vier preenchido, gravar na nova coluna `OS_Numero` de `VISITAS` (índice por nome). Se não vier, comportamento atual inalterado (coluna fica vazia). Não remover nada do que já é gravado. (Opcional: `saveOS` grava `id_visita_resultado` quando o PCF informar.)

### 2.6 Drive: helpers + `salvarFotosDrivePGP` (uploadAttachment) + `getAttachments`
Pasta-raiz via Script Property `ROOT_FOLDER_ID` (config; ver Setup). Nunca hardcodar ID.
```js
function getOrCreateDriveFolder(parent, name){
  const safe=String(name||'sem_nome').replace(/[\\/:*?"<>|]/g,'-').trim()||'sem_nome';
  const it=parent.getFoldersByName(safe); return it.hasNext()?it.next():parent.createFolder(safe);
}
function getRootFolder_(){
  const id=PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
  if(!id) throw new Error('ROOT_FOLDER_ID não configurado'); return DriveApp.getFolderById(id);
}
```
Novo `case 'salvarFotosDrivePGP'` (doPost) — recebe UM arquivo por chamada (foto, assinatura ou PDF). Binário → Drive na pasta `ROOT/{cliente}/{numero_os}/`; **linha canônica → `TECH_ATTACHMENTS`**:
```js
function salvarFotosDrivePGP(body){
  // body: { osNumero, clientName, entityType, entityId, fileName, mime, dataBase64,
  //         attachmentType, caption, createdBy }
  const root=getRootFolder_();
  const fCli=getOrCreateDriveFolder(root, body.clientName||'SemCliente');
  const fOS =getOrCreateDriveFolder(fCli, body.osNumero||'SemOS');
  const blob=Utilities.newBlob(Utilities.base64Decode(body.dataBase64), body.mime||'image/jpeg', body.fileName||'arquivo');
  const file=fOS.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // ver Setup: avaliar DOMAIN_WITH_LINK
  const sheet=getOrCreateSheet('TECH_ATTACHMENTS', HEADERS.TECH_ATTACHMENTS);
  const id='ATT-'+Date.now()+'-'+Math.floor(Math.random()*1000);
  const url='https://drive.google.com/file/d/'+file.getId()+'/view';
  sheet.appendRow([ id, body.entityType||'os', body.entityId||body.osNumero||'', body.osNumero||'',
    file.getId(), url, file.getName(), body.mime||'', body.attachmentType||'foto',
    body.caption||'', new Date(), body.createdBy||'', 'SIM' ]);
  // atualizar ORDENS_SERVICO: drive_folder_id/url (uma vez) e, se pdf_os, pdf_url — localizar coluna por NOME
  return { status:'ok', attachment_id:id, file_id:file.getId(), url:url,
           folder_id:fOS.getId(), folder_url:fOS.getUrl() };
}
```
Novo `case 'getAttachments'` (doGet) — dado `osNumero` (ou `entityType`+`entityId`), retorna linhas ativas de `TECH_ATTACHMENTS` (para a galeria do PCM).

### 2.7 `version` → `'20'`.

---

## 3. PCM (`v28`→`v29`, ou `atual+1`) — depois do GAS

> Nav estável: aba principal via `tab-btn` + `switchTab(tab)`. Adicionar **aba nova** sem tocar nas existentes.

1. Nova aba **"Ordens de Serviço"** no nav principal (`switchTab('os')`), painel próprio.
2. **Criar OS:** form com cliente, **máquinas (sugestão — ver 5)**, técnicos responsáveis (de `CAT_TECNICOS`), descrição, `data_prevista` → `saveOS`.
3. **Listar OS** com filtros status/técnico/cliente/período → `getOS`. Tabela reusando `.table-wrap`/`.badge` existentes; badge por `status`.
4. **Abrir OS:** modal/painel com detalhes, técnicos e máquinas vinculadas, `status`; se houver `id_visita_resultado`, link para a visita no histórico. **Galeria de anexos** lendo `getAttachments?osNumero=...` (miniaturas → links Drive). Botões **"Ver pasta no Drive"** (`drive_folder_url`) e **"Ver relatório PDF"** (`pdf_url`) quando preenchidos. Sem base64, sem store local.
5. Tudo lê/grava via GAS; **nenhum dado de OS em store local** além de cache transitório.
6. `version` → próximo número.

---

## 4. PCF (`v12`→`v13`) — por último

> Estável: `_pgpTipoVisita` (`'inspecao'`|`'preventiva'`), `pgpSetTipoVisita(tipo)`, `pgpEnviarPreventiva(midx)`, `gerarPDF()`, `pgpPostJson(action,body)`, `_AUTH_HASH`. **Localizar:**
```bash
grep -n "_pgpTipoVisita\|function pgpSetTipoVisita\|function pgpEnviarPreventiva\|function gerarPDF\|function pgpPostJson" Proper_Field_grupotap_v12.html
```

### 4.1 Porta opcional "Ordens de Serviço" (NÃO altera inspeção/comercial)
Adicionar uma entrada nova (botão/tela) "Ordens de Serviço" que abre a tela de PIN. Os fluxos atuais (inspeção, preventiva avulsa) continuam acessíveis como hoje, **sem exigir PIN**.

### 4.2 Login por PIN
Tela com teclado numérico de 4 dígitos. Gerar hash **no cliente** com SHA-256 via `crypto.subtle.digest` (sem lib externa) e enviar via `pgpPostJson('loginTecnico', { pin_hash })`. **Nunca enviar PIN em texto.** Em sucesso: guardar `SES.tecnico` em memória (+ `sessionStorage` para sobreviver a refresh da sessão; **não** usar `proper_admin_v2`). Exibir a lista de OS retornada.

### 4.3 Lista de OS do técnico
Mostrar apenas OS atribuídas ao técnico logado (já vêm filtradas do `loginTecnico`; opcionalmente atualizar via `getOS?tecnico=<id>`). Tocar numa OS → abre a OS.

### 4.4 Abrir OS · cronômetro · **vínculo de máquinas (CORREÇÃO)**
- Ao abrir: registrar `inicio_atendimento` (timestamp local ISO) e persistir via `saveOS`.
- **Máquinas:** carregar `maquinas_vinculadas` da OS como **sugestão pré-marcada**, mas **o técnico pode adicionar/remover** máquinas a partir da lista completa do cliente. O conjunto **final ajustado pelo técnico** é o que vale e é gravado de volta via `saveOS` (`maquinas_vinculadas`). PCM sugere; **PCF decide.**
- A coleta em si **reusa o fluxo preventiva existente** (`_pgpTipoVisita='preventiva'`), agora carimbado com o `os_numero` da OS. **Não duplicar** a UI de coleta.
- Ao concluir: registrar `fim_atendimento`; calcular duração local (min) para exibir no resumo/PDF; `saveOS` com `status='concluida'`, `id_visita_resultado` (id da visita gerada) e os timestamps.
- No envio da preventiva (`pgpEnviarPreventiva`), incluir `os_numero` no payload para o GAS gravar em `VISITAS.OS_Numero` (2.5).

### 4.5 PDF por máquina individual (refator por composição)
Refatorar `gerarPDF()` → `gerarPDF(maquinaIdx = null)`: `null` = **todas (comportamento atual preservado)**; índice = só aquela máquina. Antes de gerar (no fluxo OS), exibir modal "Gerar relatório de qual máquina?" com opção "Todas". Isso corrige o timeout do jsPDF em OS com várias máquinas. **O fluxo de inspeção continua chamando `gerarPDF()` sem argumento — comportamento idêntico.**

### 4.6 Assinatura digital (canvas) — só no fluxo OS
Ao fechar a OS, exibir dois blocos `<canvas>` com suporte a touch e mouse: **assinatura do técnico** e **assinatura do responsável do cliente** (com campo de texto p/ nome do responsável acima). Capturar cada uma como PNG via `canvas.toDataURL('image/png')`. Embutir no PDF (já há suporte a imagem no `gerarPDF`) **e** enfileirar para upload ao Drive (4.7) como `attachmentType:'assinatura_tecnico'`/`'assinatura_cliente'`. Não exigir assinatura no fluxo de inspeção.

### 4.7 Fila durável de anexos + upload ao Drive (corrige `photoStore` volátil)
- Criar fila persistida em localStorage **`proper_pending_att`** (trânsito, **não** fonte de verdade). Cada item = `{ payload:{osNumero,clientName,entityType,entityId,fileName,mime,dataBase64,attachmentType,caption,createdBy}, ts }`. Toda foto/assinatura/PDF do fluxo OS entra na fila **assim que coletada**, sobrevivendo a refresh/queda.
- `pgpUploadAnexo(payload)` → `pgpPostJson('salvarFotosDrivePGP', payload)`; em sucesso remove o item da fila; em falha mantém. No load do app, tentar flush.
- **Nenhuma foto/assinatura do fluxo OS pode ir só para o PDF** — tudo que entra no PDF tem item correspondente na fila/`TECH_ATTACHMENTS`.
- **PDF da OS:** após `gerarPDF(...)`, obter `doc.output('datauristring')` (tirar prefixo), enfileirar como `attachmentType:'pdf_os'`, `mime:'application/pdf'`. Com a URL/folder retornados, chamar `saveOS` para gravar `pdf_url`/`drive_folder_id`/`drive_folder_url` na OS. Manter o download local atual.
- **Escopo:** fotos do modo **inspeção** continuam como hoje (sem Drive). Só o fluxo OS/preventiva usa a fila e o Drive.

### 4.8 `version` → `'13'`.

---

## 5. CORREÇÃO DE PROJETO REGISTRADA (vínculo de máquinas)
O planejamento previa "PCM seleciona, PCF só mostra as pré-selecionadas". **Alterado por decisão do produto:** `maquinas_vinculadas` definida no PCM é **sugestão**; o técnico em campo **ajusta livremente** (add/remove) e o conjunto final é gravado pelo PCF. Aplicado em 3.2, 4.4.

---

## SETUP MANUAL (necessário por causa do Drive — documentar no topo do GAS)
1. Apps Script → Project Settings → **Script Properties** → criar `ROOT_FOLDER_ID` = ID da pasta-mãe no Drive da Proper.
2. Como o GAS passa a usar `DriveApp` pela 1ª vez: **reautorizar** o script (rodar uma função no editor e aceitar o escopo Drive) e **publicar nova versão** (Gerenciar implantações → Nova versão).
3. Decidir compartilhamento: `ANYONE_WITH_LINK` (simples; IDs não-adivinháveis; cômodo p/ portal do cliente) vs `DOMAIN_WITH_LINK` (mais fechado). Registrar a escolha; o código usa `ANYONE_WITH_LINK` por padrão.

---

## FORA DE ESCOPO DESTA FASE (rounds seguintes)
- **PCM v30:** aba dedicada de histórico de OS (além do que já aparece na aba OS).
- Fase 3: envio por e-mail (botão Gmail), notificação automática (MailApp), evento no Calendar.
- Reuso de fotos de visitas anteriores (thumbs); PWA instalável.

---

## VERSIONAMENTO (repo PLANO — sem subpastas)
- GAS: salvar `GAS_properCare_v20.js`; mover `GAS_properCare_v19.js` para uma pasta de versões anteriores **se ela existir no repo**; caso o repo seja plano sem pasta de arquivo, apenas manter ambos e reportar (não criar estrutura nova sem confirmação).
- PCM: salvar `PCM_v{atual+1}.html`. **`index.html`:** o repo atual **não tem** `index.html` na raiz — **não** criar/copiar `index.html` por conta própria; confirmar com o mantenedor como o GitHub Pages serve antes de adotar esse passo.
- PCF: salvar `Proper_Field_grupotap_v13.html`, mesma regra de index.
> Não inventar pastas/`index.html` que não existem no repo. Em dúvida sobre layout de deploy → **reportar**.

---

## HANDOFF — AUDITORIA DO CLAUDE CODE (lendo os arquivos reais)
1. Versões alvo conferem com o repo real (PCM pode ser v27 — reconciliar antes).
2. Inspeção/comercial do PCF **intactos**; `gerarPDF()` sem argumento mantém comportamento; login PIN não gateia fluxos antigos.
3. Abas novas criadas via `getOrCreateSheet`; `VISITAS` recebeu só `OS_Numero` **ao final**; nenhuma coluna reordenada.
4. `loginTecnico` nunca recebe/loga PIN em texto; hash SHA-256 no cliente.
5. Numeração `PGP-AAAA-NNNN` reconciliada pela `ORDENS_SERVICO`; Script Property não é autoridade única.
6. `saveOS`/`getOS`/`updateOS` localizam coluna **por nome**; `maquinas_vinculadas` final = ajuste do PCF.
7. **Mídia → Drive + Sheet:** nenhuma foto/assinatura do fluxo OS entra no PDF sem item correspondente em `TECH_ATTACHMENTS`; fila `proper_pending_att` é persistida e só esvazia após confirmação do GAS; fotos de **inspeção** seguem como hoje (sem Drive). `ROOT_FOLDER_ID` lido de Script Property, não hardcodado.
8. Funções protegidas e `proper_admin_v2` inalterados. Rodar gates de grep.

---

## GATES DE GREP
```bash
# GAS
grep -c "CAT_TECNICOS\|ORDENS_SERVICO\|TECH_ATTACHMENTS" GAS_properCare_v20.js     # >= 3
grep -c "case 'loginTecnico'\|case 'saveOS'\|case 'getOS'\|case 'salvarFotosDrivePGP'\|case 'getAttachments'" GAS_properCare_v20.js   # 5
grep -c "function generateOsNumber_\|PGP-" GAS_properCare_v20.js                   # >= 1
grep -c "getSheetData('ORDENS_SERVICO')\|ORDENS_SERVICO" GAS_properCare_v20.js     # numeração lê a Sheet
grep -c "DriveApp\|ROOT_FOLDER_ID" GAS_properCare_v20.js                           # > 0 (Drive agora em escopo)
grep -c "function machineKey\|function clientKey" GAS_properCare_v20.js            # igual à v19

# PCF
grep -c "crypto.subtle\|loginTecnico" Proper_Field_grupotap_v13.html              # >= 2
grep -c "proper_pending_att\|pgpUploadAnexo\|salvarFotosDrivePGP" Proper_Field_grupotap_v13.html  # >= 2
grep -c "canvas\|toDataURL" Proper_Field_grupotap_v13.html                         # >= 1 (assinatura)
grep -c "function gerarPDF" Proper_Field_grupotap_v13.html                         # 1 (refatorada, não duplicada)
grep -c "_pgpTipoVisita\|pgpEnviarPreventiva" Proper_Field_grupotap_v13.html       # inalterados (inspeção viva)
grep -c "proper_admin_v2\|machineKey\|clientKey" Proper_Field_grupotap_v13.html    # inalterado vs v12

# PCM
grep -c "switchTab('os')\|Ordens de Servi\|saveOS\|getOS\|getAttachments" PCM_v*.html  # >= 1 no arquivo novo
grep -c "proper_admin_v2" PCM_v*.html                                              # inalterado
```
Qualquer gate fora do esperado → reportar, não forçar.
