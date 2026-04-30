# Central de Controle Eletrovia

Aplicação operacional em HTML, CSS e JavaScript para controlar carteira, planejamento, replanejamento, realizado, perdas, demandas futuras, carga em lote, permissões, administração e indicadores.

## O que foi entregue

- `index.html`: tela principal pronta para publicação como arquivo estático.
- `styles.css`: design system corporativo responsivo, com densidade de dashboard operacional.
- `app.js`: fluxos de carteira, filtros, paginação, ações, lote, futuras, indicadores, administração e logs.
- `sharepoint-adapter.js`: repositório local e esqueleto REST para SharePoint Lists.
- `data/sample-data.js`: massa de demonstração com ordens SAP, demandas futuras, usuários, configurações, históricos e logs.
- `docs/sharepoint-schema.md`: especificação das listas SharePoint.
- `assets/concept-dashboard.png`: conceito visual usado como referência de design.

## Como abrir agora

Abra `index.html` no navegador. A versão atual usa `localStorage` para simular as listas SharePoint, então os inputs são preservados no navegador enquanto o protótipo é testado.

No SharePoint, o usuário é identificado automaticamente pelo contexto da página Microsoft 365. O e-mail logado é cruzado com a lista `Usuarios_Central_Eletrovia`, que define nome e perfil. Em execução local, o protótipo usa o primeiro usuário ativo da massa de demonstração.

Perfis disponíveis:

- Administrador: configura cadastros, parâmetros, usuários e recarrega dados de demonstração.
- Editor: planeja, replaneja, registra realizado/perda, cria demanda futura e salva lote.
- Gestor: consulta indicadores, carteira e logs.
- Visualizador: consulta, filtra e exporta.

## Publicação no SharePoint

1. Suba os arquivos para uma biblioteca do site do Teams/SharePoint.
2. Publique `index.html` em uma página com Web Part de arquivo, Embed ou política corporativa equivalente.
3. Crie as listas conforme `docs/sharepoint-schema.md`.
4. Mantenha a base SAP BO em biblioteca separada e trate essa base como somente leitura.
5. Mantenha também a base de realizados exportada do SAP BO, para atualização automática de `DataRealizada`, perda e histórico.
6. Para ligar em SharePoint REST, abra a página com `?data=sharepoint` e ajuste o mapeamento de campos em `sharepoint-adapter.js`.

## Arquitetura

A UI trabalha em camadas:

- Base SAP/BO: origem bruta somente leitura.
- Controle operacional: lista tratada `Controle_Demandas_Eletrovia`, com `ID_Demanda_Controle`.
- Histórico e auditoria: listas específicas para planejamento, replanejamento, realizado/perda e logs.
- Administração: usuários, motivos, justificativas, perfis de perda e parâmetros.

O app não renderiza todos os registros de uma vez. Ele aplica filtros, busca e paginação para manter a tela leve. Na integração real, a mesma fronteira deve ser usada para empurrar filtros para SharePoint REST/Graph com `$filter`, `$select`, `$top` e paginação.

## Integração com Excel

A tela de Carga em Lote aceita CSV nativamente. Para `.xlsx`, o HTML carrega SheetJS via CDN:

```html
https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
```

Se o ambiente corporativo bloquear CDN, baixe a biblioteca aprovada para a própria biblioteca SharePoint e ajuste o `src` em `index.html`.

Existem dois fluxos de arquivo, mas só um deles fica exposto na tela:

- Carga em lote operacional: aparece na aba `Carga em Lote` e aplica planejamento, replanejamento, perda e comentários com validação de válidos, alertas e erros.
- Atualização de realizados: roda automaticamente no SharePoint lendo `realizedExcelFileName` na mesma biblioteca configurada para SAP BO; cruza ordem e data realizada com a carteira e baixa as ordens executadas.

## Status e travas operacionais

A carteira separa `Status` e `Substatus`:

- `Status`: A Planejar, Planejado, Replanejado ou Realizado.
- `Substatus`: Perda, No Prazo, Fora do Prazo ou Pendente.

Quando houver pendência, o detalhe lateral mostra exatamente o que falta, por exemplo motivo de perda, justificativa de perda, motivo de replanejamento ou justificativa de replanejamento.

As ações são travadas por estágio:

- A Planejar: libera Planejar e Realizado/Perda.
- Planejado ou Replanejado: libera Replanejar e Realizado/Perda.
- Realizado: libera apenas complemento de perda.
- Histórico fica sempre disponível.

## Configuração por grupo/subgrupo

- Motivo é grupo; Justificativa é subgrupo do Motivo.
- Perfil de perda é grupo; Justificativa de perda é subgrupo do Perfil de perda.

Isso evita justificativas soltas e melhora a governança dos cadastros.

## Próximos passos técnicos

1. Mapear nomes internos das colunas SharePoint e preencher os métodos REST de `SharePointRestRepository`.
2. Criar índices nas colunas usadas em filtro: ordem, centro de trabalho, competência, vencimento, status operacional e `ID_Demanda_Controle`.
3. Validar os nomes internos das listas de configuração com grupo/subgrupo.
4. Conectar Power BI às listas tratadas para camada executiva.
