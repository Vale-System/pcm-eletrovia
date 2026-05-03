# Central de Controle Eletrovia

Aplicacao operacional estatica em HTML, CSS e JavaScript para controle de carteira PCM Eletrovia com JSON SAP BO como fonte bruta de leitura e Supabase como camada mestre/operacional.

## Arquitetura

- `base/base_ordens.json`: ordens SAP BO, somente leitura.
- `base/base_futuras.json`: demandas futuras/sistematicas SAP BO, somente leitura.
- `supabase-adapter.js`: acesso REST ao Supabase com chave anon publica.
- `app.js`: normalizacao, enriquecimento, filtros, fluxos operacionais, administracao, lote, login e indicadores.
- `index.html` e `styles.css`: aplicacao estatica responsiva.

O sistema nao copia toda a base SAP para `controle_demandas_eletrovia`. Essa tabela guarda apenas o estado operacional alterado pelo sistema: planejamento, replanejamento, realizado/perda, comentarios, responsavel, status operacional, origem e data/hora da ultima atualizacao.

## Tabelas Supabase

Tabelas validadas via REST em 03/05/2026:

- `cadastro_centros_trabalho`: 592 registros. Fonte mestre para centro, gerencia, supervisao, planejador OM e programador.
- `controle_demandas_eletrovia`: camada operacional por `id_demanda_controle`.
- `usuarios_central_eletrovia`: login, perfil e permissoes.
- `configuracoes_motivos` e `configuracoes_justificativas`: motivos e justificativas de replanejamento.
- `configuracoes_perfil_perdas` e `configuracoes_justificativas_perdas`: perfis e justificativas de perda.
- `historico_planejamento`, `historico_replanejamento`, `historico_realizado_perdas`: auditoria operacional.
- `logs_sistema`: logs padronizados por acao, usuario, lista, referencia, detalhe, nivel, modulo, status e data/hora.
- `cargas_lote` e `cargas_lote_itens`: auditoria de importacoes.

`parametros_sistema` ainda nao existe no Supabase atual. Por isso a aba Parametros fica informativa/desativada ate a tabela ser criada.

## Centro De Trabalho Mestre

`cadastro_centros_trabalho` deve conter:

`centro_trabalho`, `centro_trabalho_chave`, `gerencia`, `supervisao`, `planejador_om`, `planejador_om_email`, `planejador_om_matricula`, `programador`, `programador_email`, `programador_matricula`, `area`, `observacao`, `ativo`, `created_at`, `updated_at`, `created_by`, `updated_by`.

`centro_trabalho_chave` e padronizado em maiusculo, sem acentos e sem espacos extras. A carteira compara o centro vindo de `base_ordens.json` e `base_futuras.json` contra essa chave e sobrescreve gerencia/supervisao/responsaveis com a fonte oficial do Supabase.

A Administracao mostra tambem centros encontrados nos JSON que ainda nao possuem cadastro mestre ativo.

## Normalizacao

Modelo interno principal:

`id`, `ordem`, `tipoDemanda`, `descricao`, `gerencia`, `supervisao`, `centroTrabalho`, `localInstalacao`, `vencimento`, `competencia`, `tipoOM`, `prioridade`, `statusSistema`, `statusUsuario`, `toleranciaMin`, `toleranciaMax`, `dataPlanejada`, `dataReplanejadaAtual`, `dataRealizada`, `perda`, `motivoPerda`, `justificativaPerda`, `comentario`, `usuarioResponsavel`, `origem`, `planejadorOM`, `programador`.

Competencia e convertida para `YYYY-MM` a partir de formatos como `2026-05`, `202605`, `2026-05-01`, `01/05/2026`, texto com mes e serial Excel. Prioridade e convertida para `Alto`, `Medio`, `Baixo` ou `Nao informado`.

## Login E Permissoes

O login usa e-mail Vale + matricula cadastrada em `usuarios_central_eletrovia`. A matricula e digitada como senha, nao aparece na tela e nao e salva no `localStorage`; a sessao local guarda somente o e-mail.

Perfis:

- `Administrador`: configura usuarios, motivos, perdas, centros, parametros e acessa todos os fluxos.
- `Editor`: planeja, replaneja, realiza, registra perdas, exporta e faz carga em lote.
- `Planejador`: planeja, replaneja, cria/vincula demandas futuras, exporta e faz carga em lote.
- `Gestor`: consulta carteira, indicadores e historico.
- `Visualizador`: consulta e exporta.

As permissoes finas gravadas no usuario controlam botoes e menus: planejar, replanejar, realizar, configurar, exportar e carga em lote.

## Filtros

A Carteira usa filtros de multipla selecao com busca digitavel e logica cruzada. Gerencia, supervisao e centro de trabalho se restringem mutuamente com base no recorte atual. O mesmo componente atende campos grandes como origem, tipo OM, competencia, prioridade, planejador OM, programador, status e cadastro do centro.

Os Indicadores usam o mesmo recorte filtrado da Carteira.

## Fluxos Operacionais

- Planejamento: exige data planejada, atualiza/insere `controle_demandas_eletrovia`, grava `historico_planejamento` e `logs_sistema`.
- Replanejamento: exige nova data, motivo e justificativa dependente do motivo; incrementa quantidade de replanejamentos, grava historico e log.
- Realizado/perda: grava data realizada, perda, perfil/motivo, justificativa, comentario, evidencia e log.
- Demandas futuras: lista `base_futuras.json` e demandas criadas pelo sistema; demandas sem OM continuam rastreadas por `ID_Demanda_Controle`.
- Carga em lote: importa CSV/XLSX, valida OM ou `ID_Demanda_Controle` contra a carteira JSON/Supabase, grava linhas validas na camada operacional e audita validos/alertas/erros em `cargas_lote` e `cargas_lote_itens`.

## Power Automate, GitHub Actions E Power BI

- Power Automate deve publicar/atualizar `base_ordens.json` e `base_futuras.json` a partir do SAP BO sem editar registros operacionais do Supabase.
- GitHub Actions pode versionar/publicar os arquivos estaticos e validar sintaxe (`node --check app.js` e `node --check supabase-adapter.js`).
- Power BI deve consumir a carteira final enriquecida ou as tabelas Supabase operacionais, respeitando a tabela mestre de centros para evitar divergencia de gerencia/supervisao.

## RLS

As politicas RLS precisam permitir leitura anonima controlada para as tabelas usadas pelo frontend e escrita somente nas operacoes esperadas: controle, historicos, logs, carga em lote, usuarios/configuracoes/centros para administradores. Para producao, evoluir o login para Supabase Auth ou Azure AD/Microsoft Entra ID.

## Execucao

Use um servidor estatico para permitir `fetch()` dos JSON:

```bash
python -m http.server 4173 --bind 127.0.0.1
```

Abra `http://127.0.0.1:4173/index.html`.
