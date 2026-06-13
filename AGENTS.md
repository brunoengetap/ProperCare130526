# AGENTS.md — ProperCare PGP

> Lido por Codex/Claude Code ao iniciar a sessão. **Reescrever a cada nova sessão** para refletir o escopo atual. Escopo antigo aqui = execução errada.

## Sessão atual: FASE 1 — Ordens de Serviço (completa)

**Tarefa única desta sessão:** implementar a Fase 1 conforme `PROMPT_Codex_OS_Fase1_v2.md` (na raiz do repo). Esse prompt é a especificação autoritativa — siga-o à risca.

**Inclui:** login por PIN, OS numeradas, filtro de OS por técnico, cronômetro início/fim, seleção de máquinas por OS (ajustável em campo), PDF por máquina individual, assinatura digital e salvamento de fotos/PDF no Google Drive com referência na Sheet.

**Versões alvo:** `GAS_properCare_v19.js → v20`, `PCM_v28.html → v29`, `Proper_Field_grupotap_v12.html → v13`.

## Antes de editar (obrigatório)
1. `ls -1 *.html *.js` — confirmar a versão REAL de cada arquivo no repo; alvo = `atual+1`. Se divergir do esperado acima, **parar e reportar**.
2. Abrir o arquivo real, localizar pontos por `grep`, editar seguindo o padrão ao redor. **Não confiar em números de linha; não inventar nomes.** Landmark não encontrado → **parar e reportar**.

## Restrições invioláveis
- Nunca modificar: `machineKey()`, `clientKey()`, `proper_admin_v2` (schema/nome), nem as assinaturas de `pgpFlushPending`, `pgpPostJson`, `pgpBuildVisitRequest`, `pgpBuildPartsUpdateRequest`.
- Funções de PDF/foto/inspeção do PCF: estender **só por composição** (parâmetro novo com default que preserva o comportamento atual). **Não tocar no fluxo de inspeção/comercial** (`_pgpTipoVisita==='inspecao'`).
- Schema **append-only**: abas novas livres; em abas existentes, só acrescentar coluna **ao final**, nunca reordenar.
- Soft delete apenas (`Ativo = NÃO`). Sem exclusão física.
- **Sheet (`Proper_PGP_DB`) = única fonte de verdade.** Binário (foto/PDF/assinatura) no Drive; referência **sempre** na Sheet (`TECH_ATTACHMENTS`). localStorage/sessionStorage = cache/sessão.

## Ordem de deploy (nunca inverter)
**GAS → PCM → PCF.** O GAS (schema) deve estar publicado antes de qualquer frontend gravar dados novos.

## Versionamento (repo PLANO — sem subpastas)
- Salvar o novo arquivo numerado (`*_v{N+1}`). Manter o anterior no repo.
- **Não** criar pastas (`PCM/`, `Versões anteriores/`) nem `index.html` que não existem no repo. Em dúvida sobre o deploy → reportar.

## Setup manual necessário nesta sessão (Drive)
- Criar Script Property `ROOT_FOLDER_ID` (pasta-mãe no Drive).
- 1º uso de `DriveApp`: reautorizar o script e publicar nova versão (Gerenciar implantações → Nova versão).

## Ao final
- Rodar os **gates de grep** do prompt. Qualquer gate fora do esperado → reportar, não forçar.
- Confirmar: inspeção/comercial intactos; funções protegidas inalteradas; nenhuma foto do fluxo OS só no PDF (toda evidência tem linha em `TECH_ATTACHMENTS`).
