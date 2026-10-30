# Fluxo de Caixa — Planilha Inteligente com CRUD VSTO

Planilha Excel (`Fluxo_Caixa.xlsx`) com add-in VSTO (.NET 4.8) e formulários WebView2 para gestão completa de dados.

## Requisitos

- Microsoft Excel (Office 2016+)
- [Visual Studio Tools for Office Runtime](https://www.microsoft.com/download/details.aspx?id=48217)
- [WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/)
- Visual Studio 2022 com carga de trabalho **Office/SharePoint** (para compilar)

### Verificar / instalar Visual Studio 2022

Execute **como Administrador** na raiz do repositório:

```powershell
.\install-vs2022-vsto.ps1
```

Componentes necessários (incluídos em `.vsconfig`):

- `Microsoft.VisualStudio.Workload.Office` — VSTO / Office Tools
- `Microsoft.VisualStudio.Workload.ManagedDesktop` — .NET desktop
- `Microsoft.Net.Component.4.8.TargetingPack` — .NET Framework 4.8

Após a instalação inicial, **reinicie o PC** se o instalador solicitar.

### Confiar no add-in VSTO (obrigatório na 1ª execução)

Se aparecer erro *"local não consta na sua lista de locais confiáveis"*, execute:

```powershell
.\trust-vsto.ps1
```

Isso adiciona pastas confiáveis no Excel, instala o certificado de assinatura e registra a inclusion list VSTO.

Se ainda falhar, execute **como Administrador**:

```powershell
.\trust-vsto.ps1 -AllUsers
```

**Importante:** abra sempre a planilha publicada em `fluxo_caixa_crud\bin\Release\Fluxo_Caixa.xlsx` (junto com `fluxo_caixa_crud.dll` e `.vsto`). Não abra a cópia solta na raiz `EXCEL\Fluxo_Caixa.xlsx`.

**Abertura normal:** o componente abre como qualquer outra planilha do Excel, sem instância dedicada — não fecha, move ou esconde as planilhas que o usuário já tenha abertas. Use `.\open-fluxo.ps1` no desenvolvimento; na distribuição, copie **`Atalho - Fluxo de Caixa.vbs`** para a Área de Trabalho — os `.dll` ficam na pasta de instalação.

**Abrir pelo atalho** (`.vbs` / `.bat`): além de abrir a planilha, o launcher confirma a configuração e remove apenas um lock órfão do Fluxo antes de abrir, evitando o erro "Arquivo em uso". Abrir o `.xlsx` pelo duplo clique direto no Excel também funciona, mas nesse caso não há essa limpeza de lock.

## Estrutura

```
fluxo_caixa_crud/
├── Fluxo_Caixa.xlsx          # Documento hospedeiro VSTO
├── fluxo_caixa_crud.csproj
├── FluxoCaixaForm.cs         # Formulário WebView2
├── Services/                 # Bridge Excel ↔ UI
├── frontend/                 # Interface HTML/CSS/JS
└── docs/PROTOCOL.md          # Protocolo de mensagens
```

## Como compilar

**Linha de comando** (após VS 2022 configurado):

```powershell
.\build.ps1
```

Ou abra `fluxo_caixa_crud.sln` no Visual Studio e compile em **Release | Any CPU**.

Saída: `fluxo_caixa_crud\bin\Release\Fluxo_Caixa.xlsx`

## Distribuir para colegas (sem compilar)

Gere um pacote pronto (ZIP + pasta com atalho):

```powershell
.\package-for-team.ps1
```

Saída:

- `dist\FluxoCaixa\` — pasta completa para copiar na rede ou extrair
- `dist\FluxoCaixa.zip` — enviar por e-mail / Teams / SharePoint

**O colega só precisa:**

1. Extrair o ZIP em uma pasta fixa (ex.: `C:\FluxoCaixa\`)
2. Dar duplo clique em **`Abrir Fluxo de Caixa.bat`**
3. Na 1ª vez a configuração roda sozinha; depois abre o Excel
4. Faixa **Fluxo de Caixa** → **Gestão de Dados**

Não é necessário Visual Studio, MSBuild nem scripts de build.

**Requisitos no PC do colega** (em geral já instalados com Office / Windows):

- Microsoft Excel
- [VSTO Runtime](https://aka.ms/vstoruntimedownload) (geralmente já presente)
- [WebView2 Runtime](https://go.microsoft.com/fwlink/p/?LinkId=2124703) (geralmente já presente no Windows 10/11)

Se aparecer erro de confiança, executar uma vez: **`Configurar (como administrador).bat`**

Para republicar após alterações: `.\package-for-team.ps1` (recompila e gera ZIP novo).

## Como usar (desenvolvimento)

1. Execute `.\trust-vsto.ps1` (uma vez por máquina).
2. Abra com **`.\open-fluxo.ps1`** ou **`fluxo_caixa_crud\bin\Release\Fluxo de Caixa.vbs`**
3. Na faixa de opções, aba **Fluxo de Caixa** → **Fluxo**.
3. Use os módulos na barra lateral:
   - **CAPEX**, **Novos Negócios**, **Premissas**, **Bônus** — CRUD tabular
   - **Lançamentos por Unidade** — edição matricial (Realizado/Previsto)
   - **Premissas Mensais** — premissas calibradas e CAPEX mensal
   - **Resumo** — visão consolidada e alertas

## Abas de suporte (criadas automaticamente)

- `_CONFIG` — metadados de mapeamento das 42 abas
- `_CRUD_LOG` — log de auditoria de alterações

## Protocolo

Comunicação JSON via WebView2. Detalhes em [docs/PROTOCOL.md](docs/PROTOCOL.md).

## Notas

- Valores monetários trafegam como centavos inteiros (`valorCents`).
- Abas consolidadas com fórmulas `#REF!` são sinalizadas no módulo Resumo.
- Dados são gravados diretamente nas células originais da planilha.
