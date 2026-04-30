# Especificação das Listas SharePoint

Use nomes internos sem acento quando possível. As colunas abaixo estão alinhadas ao protótipo HTML/CSS/JS e à arquitetura de camadas.

## Base_Ordens_SAP

Lista ou Excel bruto vindo do SAP BO. Somente leitura para a aplicação.

| Coluna | Tipo sugerido | Observação |
| --- | --- | --- |
| OrdemSAP | Texto | Pode ser vazio para bases futuras, mas a base SAP real tende a preencher. |
| Descricao | Texto múltiplas linhas | Descrição da OM/demanda. |
| TipoOM | Escolha | Preventiva, Corretiva, Inspeção, Calibração, Sistemática. |
| Gerencia | Escolha | Recorte organizacional. |
| CentroTrabalho | Texto | Indexar. |
| LocalInstalacao | Texto | Indexar se houver alto volume. |
| Vencimento | Data | Indexar. |
| Prioridade | Escolha | Alta, Média, Baixa. |
| StatusSistema | Texto | Status SAP. |
| Competencia | Texto | Formato `AAAA-MM`, indexar. |
| ToleranciaMin | Data | Limite mínimo. |
| ToleranciaMax | Data | Limite máximo. |
| OrigemArquivo | Texto | Nome do arquivo SAP BO. |
| DataImportacao | Data/hora | Auditoria da carga. |

## Controle_Demandas_Eletrovia

Lista principal tratada pela Central de Controle.

| Coluna | Tipo sugerido | Observação |
| --- | --- | --- |
| ID_Demanda_Controle | Texto | Chave interna única, indexar e exigir valor único. |
| OrdemSAP | Texto | Indexar, pode ficar vazio. |
| TipoDemanda | Escolha | SAP BO, Sistemática, Antecipada, Extraordinária. |
| Descricao | Texto múltiplas linhas | Campo visível na carteira. |
| Gerencia | Escolha | Filtro superior. |
| CentroTrabalho | Texto | Indexar. |
| LocalInstalacao | Texto | Indexar. |
| Vencimento | Data | Indexar. |
| Competencia | Texto | Indexar. |
| TipoOM | Escolha | Tipo técnico da OM. |
| Prioridade | Escolha | Alta, Média, Baixa. |
| StatusSistema | Texto | Status vindo do SAP. |
| ToleranciaMin | Data | Usado no status automático. |
| ToleranciaMax | Data | Usado no status automático. |
| DataPlanejada | Data | Planejamento atual inicial. |
| DataReplanejadaAtual | Data | Última data replanejada. |
| DataRealizada | Data | Baixa operacional. |
| StatusOperacional | Escolha | A Planejar, Planejado, Replanejado, Realizado. |
| SubstatusOperacional | Texto | Perda, No Prazo, Fora do Prazo, Pendente. Pode ser calculado pela UI ou persistido para Power BI. |
| Perda | Sim/Não | Indicador operacional. |
| MotivoPerda | Escolha | Vem de `Configuracoes_Perfil_Perdas`. |
| JustificativaPerda | Escolha | Vem de `Configuracoes_Justificativas_Perdas`. |
| Comentario | Texto múltiplas linhas | Comentário operacional. |
| UsuarioResponsavel | Pessoa ou Grupo | Responsável pela última ação. |
| DataUltimaAtualizacao | Data/hora | Auditoria funcional. |
| OrigemInformacao | Escolha | SAP BO, Carga em Lote, Demanda Antecipada, Vínculo SAP. |
| QuantidadeReplanejamentos | Número | Apoia ranking e auditoria. |
| Frequencia | Escolha | Para sistemáticas: mensal, trimestral, etc. |
| Observacao | Texto múltiplas linhas | Campo livre para demanda futura. |
| VinculadaEm | Data/hora | Quando uma demanda futura recebeu ordem SAP. |

## Historico_Planejamento

| Coluna | Tipo sugerido |
| --- | --- |
| ID_Demanda_Controle | Pesquisa ou Texto |
| DataAnterior | Data |
| NovaData | Data |
| Usuario | Pessoa ou Texto |
| DataHoraAlteracao | Data/hora |
| Comentario | Texto múltiplas linhas |

## Historico_Replanejamento

| Coluna | Tipo sugerido |
| --- | --- |
| ID_Demanda_Controle | Pesquisa ou Texto |
| Motivo | Escolha |
| Justificativa | Escolha |
| DataAnterior | Data |
| NovaData | Data |
| Usuario | Pessoa ou Texto |
| DataHoraAlteracao | Data/hora |
| QuantidadeReplanejamentos | Número |

## Historico_Realizado_Perdas

| Coluna | Tipo sugerido |
| --- | --- |
| ID_Demanda_Controle | Pesquisa ou Texto |
| DataRealizada | Data |
| Perda | Sim/Não |
| MotivoPerda | Escolha |
| JustificativaPerda | Escolha |
| Comentario | Texto múltiplas linhas |
| Evidencia | Hiperlink ou Anexo |
| Usuario | Pessoa ou Texto |
| DataHoraRegistro | Data/hora |

## Usuarios_Central_Eletrovia

| Coluna | Tipo sugerido |
| --- | --- |
| Nome | Texto |
| Email | Texto, valor único |
| Matricula | Texto |
| Area | Texto |
| PerfilAcesso | Escolha: Visualizador, Editor, Administrador, Gestor |
| Ativo | Sim/Não |

## Configurações

As configurações trabalham com grupo e subgrupo.

### Configuracoes_Motivos

| Coluna | Tipo sugerido |
| --- | --- |
| Title | Texto |
| Chave | Texto, valor único |
| Ativo | Sim/Não |

### Configuracoes_Justificativas

| Coluna | Tipo sugerido |
| --- | --- |
| Title | Texto |
| Chave | Texto, valor único |
| MotivoChave | Texto |
| Ativo | Sim/Não |

`MotivoChave` deve apontar para `Configuracoes_Motivos.Chave`.

### Configuracoes_Perfil_Perdas

| Coluna | Tipo sugerido |
| --- | --- |
| Title | Texto |
| Chave | Texto, valor único |
| Ativo | Sim/Não |

### Configuracoes_Justificativas_Perdas

| Coluna | Tipo sugerido |
| --- | --- |
| Title | Texto |
| Chave | Texto, valor único |
| PerfilChave | Texto |
| Ativo | Sim/Não |

`PerfilChave` deve apontar para `Configuracoes_Perfil_Perdas.Chave`.

## Base de Realizados SAP BO

Pode ser Excel em biblioteca SharePoint ou lista intermediária. Campos mínimos:

| Coluna | Tipo sugerido | Observação |
| --- | --- | --- |
| OrdemSAP | Texto | Chave para cruzar com a carteira. |
| DataRealizada | Data | Baixa automática da ordem. |
| Perda | Sim/Não | Opcional. |
| PerfilPerda | Texto | Opcional. |
| JustificativaPerda | Texto | Opcional. |
| Comentario | Texto múltiplas linhas | Opcional. |

Quando uma ordem da base de realizados coincidir com `Controle_Demandas_Eletrovia.OrdemSAP`, o sistema atualiza `DataRealizada`, registra histórico e cria log técnico. Essa leitura deve ocorrer automaticamente a partir do arquivo configurado em `realizedExcelFileName`, sem upload manual na tela de carga em lote.

## Logs_Sistema

| Coluna | Tipo sugerido |
| --- | --- |
| DataHora | Data/hora |
| Usuario | Pessoa ou Texto |
| Acao | Texto |
| Lista | Texto |
| Referencia | Texto |
| Detalhe | Texto múltiplas linhas |

## Índices recomendados

Crie índices em:

- `ID_Demanda_Controle`
- `OrdemSAP`
- `CentroTrabalho`
- `Competencia`
- `Vencimento`
- `StatusOperacional`
- `SubstatusOperacional`
- `Perda`
- `DataUltimaAtualizacao`

Para volumes próximos ou acima de 70 mil registros, mantenha consultas sempre filtradas por competência, centro de trabalho, status operacional ou busca direta por ordem.
