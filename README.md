# Insight Hub

Especificação Funcional – Portal Corporativo de BI (Power BI)

Objetivo

Desenvolver um Portal Corporativo de Business Intelligence (BI) que será o ambiente oficial para acesso aos dashboards e indicadores da empresa.

O portal não será responsável por criar dashboards.

Todos os dashboards continuarão sendo desenvolvidos exclusivamente no Power BI Desktop e publicados no Power BI Service.

O portal será responsável apenas por:

 Autenticação

 Controle de acesso

 Organização dos dashboards

 Navegação

 Experiência do usuário

 Administração

 Exibição dos dashboards publicados no Power BI

A ideia é que os usuários nunca precisem acessar diretamente o portal do Power BI.

Arquitetura

A arquitetura deverá seguir exatamente este conceito:

Power BI Desktop
        │
        │ Publicar
        ▼
Power BI Service
        │
        ▼
Portal Corporativo BI
        │
        ▼
Usuários

O portal deverá consumir os dashboards publicados no Power BI.

Nenhum gráfico será recriado dentro do portal.

Requisitos Obrigatórios

Os dashboards NÃO serão desenvolvidos no portal.

Toda a construção dos dashboards continuará sendo feita no Power BI Desktop.

Sempre que um dashboard for alterado e publicado novamente no Power BI Service, o portal deverá exibir automaticamente a nova versão.

Não deverá existir necessidade de alterar código do portal para atualizar dashboards.

Integração com Power BI

O portal deverá ser preparado para integração com o Power BI Service.

Não utilizar uma implementação limitada baseada apenas em iframe.

A integração deverá ser preparada utilizando o SDK oficial do Power BI (powerbi-client), permitindo evolução futura.

Cada dashboard deverá possuir:

 Nome

 Descrição

 Categoria

 URL do relatório

 Workspace

 Ícone

 Permissões

O portal deverá carregar dinamicamente o dashboard correspondente.

Estrutura do Portal

Criar um portal corporativo moderno.

Menu lateral fixo.

Layout semelhante aos grandes portais corporativos.

Itens iniciais:

 Início

 Dashboard Executivo

 Comercial

 Financeiro

 Produção

 PCP

 Logística

 RH

 Indicadores

 Relatórios

 Alertas

 Administração

 Meu Perfil

O menu deverá ser dinâmico conforme as permissões do usuário.

Tela Inicial

Criar um cockpit executivo.

Exibir:

 Saudação personalizada

 Data

 Hora

 Última atualização dos dashboards

 Dashboards favoritos

 Dashboards recentes

 Indicadores rápidos

 Alertas

 Notificações

Interface moderna.

Dashboard Center

Criar uma central de dashboards.

Cada dashboard será apresentado como um card.

Cada card deverá possuir:

 Nome

 Descrição

 Categoria

 Ícone

 Última atualização

 Botão Abrir

Exemplo:

Dashboard Executivo

Descrição

Abrir Dashboard

Dashboard Comercial

Descrição

Abrir Dashboard

Dashboard Produção

Descrição

Abrir Dashboard

Página do Dashboard

Ao abrir um dashboard:

Exibir:

Título

Descrição

Filtros

Última atualização

Área principal destinada ao Power BI.

O dashboard deverá ocupar praticamente toda a largura da tela.

Layout limpo.

Sem elementos desnecessários.

Administração

Criar área administrativa.

Permitir gerenciar:

Usuários

Perfis

Permissões

Dashboards

Categorias

Logs

Configurações

Cadastro de Dashboards

Criar tela administrativa.

Cadastrar:

Nome

Descrição

Categoria

Ícone

URL Power BI

Workspace

Ordem de exibição

Status

Permissões

O cadastro deverá permitir adicionar novos dashboards futuramente sem alterar código.

Perfis

Criar perfis.

Administrador

Diretor

Gerente

Supervisor

Analista

Cada perfil visualizará apenas os dashboards autorizados.

Controle de Permissões

Cada dashboard poderá ser associado a um ou mais perfis.

Exemplo:

Dashboard Executivo

 Diretor

 Gerente

Dashboard Produção

 Gerente Produção

 Supervisor Produção

Dashboard RH

 RH

 Diretoria

Login

Preparar autenticação.

Estrutura preparada para integração futura com Microsoft Entra ID.

Inicialmente poderá utilizar login interno.

Responsividade

Desktop

Notebook

Tablet

Monitores ultrawide

Interface

Criar interface premium.

Inspirada em produtos como:

Microsoft Fabric

Power BI

Azure

Notion

Monday

ClickUp

SAP Fiori

Interface extremamente limpa.

Pouco texto.

Muito espaço em branco.

Cards modernos.

Ícones minimalistas.

Animações suaves.

Identidade Visual

Utilizar como referência a identidade visual da empresa Vem.

Não utilizar aparência genérica.

Criar uma aplicação elegante.

Objetivo Principal

O usuário nunca deverá precisar acessar o Power BI Service diretamente.

Todo acesso será realizado através do Portal Corporativo.

Atualização dos Dashboards

Este é o requisito mais importante do projeto.

Os dashboards continuarão sendo desenvolvidos exclusivamente no Power BI Desktop.

Fluxo esperado:

 Abrir o arquivo .pbix.

 Alterar gráficos, medidas, KPIs ou páginas.

 Publicar novamente no Power BI Service.

 O portal deverá refletir automaticamente todas as alterações, sem necessidade de qualquer modificação em seu código.

O portal não armazenará dashboards nem copiará visualizações; ele apenas exibirá os dashboards publicados.

Escalabilidade

A aplicação deverá ser preparada para crescimento.

No futuro deverão existir novos módulos como:

 IA Corporativa

 Metas e OKRs

 Aprovações

 Indicadores Estratégicos

 Documentos

 Comunicados

 Calendário Corporativo

 Central de KPIs

 Alertas Inteligentes

 Favoritos

 Pesquisa Global

A arquitetura deve permitir essa evolução sem necessidade de reestruturação.

Observações Importantes

Não recriar os dashboards no portal.

Todos os dashboards serão mantidos exclusivamente no Power BI.

O portal deve ser apenas a camada de experiência do usuário, autenticação, organização e controle de acesso.

A integração deve ser preparada utilizando o SDK oficial do Power BI (powerbi-client), evitando uma solução baseada apenas em iframes, para garantir maior flexibilidade e evolução futura.

O código deve ser modular, reutilizável e preparado para suportar dezenas de dashboards e centenas de usuários.

Resultado esperado

Ao final do projeto, a empresa terá um Portal Corporativo de BI que centraliza todos os dashboards do Power BI em uma única plataforma moderna, mantendo o Power BI como a única ferramenta de desenvolvimento e atualização dos relatórios, enquanto o portal oferece uma experiência de navegação, autenticação e gestão de acesso profissional e escalável.

This project was built with [Lovable](https://lovable.dev).

**Live app**: https://portalmundovem.lovable.app

## Build with Lovable

Continue developing this project in the [Lovable editor](https://lovable.dev/projects/8fc7ecd5-8e8c-4d22-98ea-241666f5cbd8).

- **Ship faster**: describe what you want to build and Lovable handles the code.
- **Stay in sync**: every change made in Lovable is committed straight to this repository.
- **Full ownership**: this code is yours. Push to `main` on GitHub and your changes sync back into Lovable, ready for your next prompt.

## Development

Prefer working locally? You need Node.js and npm — [install with nvm](https://github.com/nvm-sh/nvm#installing-and-updating).

```sh
git clone <this-repository-url>
cd <repository-name>
npm i
npm run dev
```
