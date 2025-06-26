# Ferramenta de Reposicionamento de Objetos e Configuração Automática do Excel (VBA)

Este documento reúne duas funcionalidades em VBA para otimizar o uso de planilhas no Microsoft Excel. A primeira é uma **ferramenta de reposicionamento de objetos**, permitindo ajustes precisos na posição de elementos visuais como imagens, formas e botões ActiveX. A segunda é uma **macro para configuração automática da interface do Excel** ao abrir o arquivo e restauração ao fechar, garantindo uma experiência limpa e controlada.

---

## 🧰 Ferramenta de Reposicionamento de Objetos para Excel (VBA)

### Visão Geral

Esta ferramenta foi desenvolvida para facilitar o reposicionamento preciso de qualquer objeto visual dentro de uma planilha do Excel, como imagens, formas, linhas e controles ActiveX. É ideal para usuários que desejam evitar ajustes manuais imprecisos e repetitivos.

### Funcionalidades Principais

- **Deteção Automática**: Identifica automaticamente o objeto selecionado.
- **Interface Gráfica Amigável**: Exibe as coordenadas atuais e permite edição direta.
- **Compatibilidade Universal**: Suporta Picture, Shape, Line, Rectangle e Controles ActiveX.
- **Confirmação de Ação**: Evita movimentos acidentais com um aviso de confirmação.
- **Validação de Dados**: Garante que apenas valores numéricos sejam inseridos nas coordenadas.

### Estrutura Técnica

A solução é composta por:

1. Um **UserForm** (`frmPosicionarObjeto`) para a interface gráfica.
2. Um **módulo de inicialização** (`IniciarReposicionamentoDeObjeto`).
3. Um **módulo de diagnóstico opcional** (`DiagnosticarSelecao`).

---

### 🔧 Passo a Passo: Como Implementar

#### 1. Crie o Formulário de Interação (UserForm)

1. Abra o Editor do Visual Basic pressionando `Alt + F11`.
2. No menu superior, vá a **Inserir > UserForm**.
3. Configure o formulário:
   - `(Name)` → `frmPosicionarObjeto`
   - `Caption` → `Reposicionar Objeto Selecionado`
4. Adicione os seguintes controles:

| Controle       | Propriedade (Name) | Propriedade (Caption/Text)               |
|----------------|--------------------|------------------------------------------|
| Label          | lblDescX           | Posição X (> direita, < esquerda)        |
| TextBox        | txtPosX            | (deixar em branco)                       |
| Label          | lblDescY           | Posição Y (> para baixo, < para cima)    |
| TextBox        | txtPosY            | (deixar em branco)                       |
| CommandButton  | btnMover           | Reposicionar                             |
| CommandButton  | btnCancelar        | Cancelar                                 |

#### 2. Código do UserForm

```vba
' ========================================================================================
' CÓDIGO DO USERFORM: frmPosicionarObjeto
' ========================================================================================
Option Explicit

Private Sub UserForm_Activate()
    If TypeOf Selection Is Range Or Selection Is Nothing Then
        MsgBox "Por favor, selecione um objeto (imagem, linha, forma) antes de continuar.", vbExclamation, "Nenhum Objeto Válido"
        Unload Me
        Exit Sub
    End If
    
    On Error Resume Next
    Me.txtPosX.Value = Round(Selection.Left, 2)
    Me.txtPosY.Value = Round(Selection.Top, 2)
    
    If Err.Number <> 0 Then
        MsgBox "O objeto selecionado (Tipo: " & TypeName(Selection) & ") não pode ser reposicionado por esta macro.", vbCritical, "Objeto Incompatível"
        Unload Me
        Exit Sub
    End If
    
    On Error GoTo 0
    Me.lblDescX.Caption = "Posição X (> direita, < esquerda)"
    Me.lblDescY.Caption = "Posição Y (> para baixo, < para cima)"
End Sub

Private Sub btnMover_Click()
    Dim novaPosX As Double
    Dim novaPosY As Double
    Dim resposta As VbMsgBoxResult

    If Not IsNumeric(Me.txtPosX.Value) Or Not IsNumeric(Me.txtPosY.Value) Then
        MsgBox "Por favor, insira apenas valores numéricos para as posições X e Y.", vbCritical, "Erro de Validação"
        Exit Sub
    End If

    novaPosX = CDbl(Me.txtPosX.Value)
    novaPosY = CDbl(Me.txtPosY.Value)

    resposta = MsgBox("Deseja realmente mover o objeto para as novas coordenadas?", _
                      vbYesNo + vbQuestion, "Confirmar Movimentação")

    If resposta = vbYes Then
        With Selection
            .Left = novaPosX
            .Top = novaPosY
        End With
        
        MsgBox "Objeto movido com sucesso!", vbInformation
        Unload Me
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub
```

#### 3. Macro de Inicialização

```vba
' =========================================================================
' MÓDULO DE INICIALIZAÇÃO (EX: Módulo1)
' =========================================================================
Option Explicit

Public Sub IniciarReposicionamentoDeObjeto()
    If Selection Is Nothing Then
        MsgBox "Selecione um objeto na folha de cálculo antes de executar a macro.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    frmPosicionarObjeto.Show
End Sub
```

#### 4. Ferramenta Opcional de Diagnóstico

```vba
' =========================================================================
' MÓDULO DE DIAGNÓSTICO
' =========================================================================
Option Explicit

Public Sub DiagnosticarSelecao()
    If Selection Is Nothing Then
        MsgBox "Nada está selecionado no momento."
    Else
        MsgBox "O tipo do objeto selecionado é: " & TypeName(Selection)
    End If
End Sub
```

---

### ✅ Como Usar a Ferramenta

1. Selecione o objeto na planilha.
2. Execute a macro `IniciarReposicionamentoDeObjeto` via `Alt + F8`.
3. Edite as coordenadas exibidas e clique em **Reposicionar**.

---

## 📐 Macro VBA para Múltiplas Planilhas (Foco na Restauração)

### Objetivo

Automatizar a configuração e restauração da interface do Excel ao abrir e fechar o arquivo. Ideal para dashboards ou documentos profissionais onde se deseja ocultar elementos da interface e proteger planilhas automaticamente.

### Funcionalidades

- Oculta barra de fórmulas e status
- Desativa grades, cabeçalhos e guias de planilhas
- Define áreas de rolagem específicas
- Protege planilhas com senha
- Restaura o ambiente original ao fechar

---

### Código Principal

#### Módulo Padrão (Ex: `Módulo1`)

```vba
' =================================================================================
' CÓDIGO A SER INSERIDO EM UM NOVO MÓDULO PADRÃO (EX: Módulo1)
' =================================================================================
Option Explicit

Private Const SENHA_PROTECAO As String = ""

Public Sub ConfigurarVisualizacaoApp()
    Dim ws As Worksheet
    Dim nomePlanilha As Variant
    Dim nomesPlanilhas As Variant

    nomesPlanilhas = Array("TITULAR", "INFORMES", "NOTAS", "CONFIGURACOES")

    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False

    On Error Resume Next

    For Each nomePlanilha In nomesPlanilhas
        Set ws = ThisWorkbook.Worksheets(nomePlanilha)
        If Not ws Is Nothing Then
            ws.Activate
            With ActiveWindow
                .DisplayGridlines = False
                .DisplayHeadings = False
                .DisplayWorkbookTabs = False
                .Zoom = 100
            End With
            ws.ScrollArea = "A1:Z100"
            ws.Protect Password:=SENHA_PROTECAO, UserInterfaceOnly:=True
        End If
    Next nomePlanilha

    ThisWorkbook.Worksheets(nomesPlanilhas(0)).Activate
    Application.Goto Reference:=ThisWorkbook.Worksheets(nomesPlanilhas(0)).Range("A1"), Scroll:=True

    On Error GoTo 0
End Sub

Public Sub RestaurarVisualizacaoPadrao()
    Dim ws As Worksheet
    Dim nomePlanilha As Variant
    Dim nomesPlanilhas As Variant

    nomesPlanilhas = Array("TITULAR", "INFORMES", "NOTAS", "CONFIGURACOES")

    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True

    On Error Resume Next

    For Each nomePlanilha In nomesPlanilhas
        Set ws = ThisWorkbook.Worksheets(nomePlanilha)
        If Not ws Is Nothing Then
            ws.Unprotect Password:=SENHA_PROTECAO
            ws.ScrollArea = ""
            ws.Activate
            With ActiveWindow
                .DisplayGridlines = True
                .DisplayHeadings = True
                .DisplayWorkbookTabs = True
            End With
        End If
    Next nomePlanilha

    ThisWorkbook.Worksheets(nomesPlanilhas(0)).Activate
    Application.Goto Reference:=ThisWorkbook.Worksheets(nomesPlanilhas(0)).Range("A1"), Scroll:=True

    On Error GoTo 0
End Sub
```

#### Módulo de Objeto `ThisWorkbook`

```vba
' =================================================================================
' CÓDIGO A SER INSERIDO NO MÓDULO DE OBJETO "ThisWorkbook"
' =================================================================================
Option Explicit

Private Sub Workbook_Open()
    Call ConfigurarVisualizacaoApp
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call RestaurarVisualizacaoPadrao
End Sub
```

---

### 💡 Dica Adicional

Para personalizar a senha de proteção:

```vba
Private Const SENHA_PROTECAO As String = "sua_senha_aqui"
```

Substitua `"sua_senha_aqui"` pela senha desejada. Deixe vazio (`""`) para proteção sem senha (menos seguro).

---

## ✅ Resumo de Funcionalidades

| Função | Acionamento | Finalidade |
|--------|-------------|------------|
| `IniciarReposicionamentoDeObjeto()` | Manualmente | Abre o formulário para reposicionar objetos |
| `ConfigurarVisualizacaoApp()` | Automático (ao abrir) | Configura interface e protege planilhas |
| `RestaurarVisualizacaoPadrao()` | Automático (ao fechar) | Restaura o ambiente padrão do Excel |
| `Workbook_Open()` | Evento interno | Chama `ConfigurarVisualizacaoApp()` |
| `Workbook_BeforeClose()` | Evento interno | Chama `RestaurarVisualizacaoPadrao()` |

---

## 📄 Exemplo Prático

Use este conjunto de macros em arquivos Excel usados como painéis de controle ou relatórios dinâmicos. Com isso, você garante:

- Interface limpa e focada
- Proteção contra alterações indesejadas
- Facilidade no posicionamento de elementos visuais
- Ambiente totalmente restaurado após o uso

---

## 📌 Observações Finais

Estas macros podem ser integradas diretamente ao seu projeto Excel e configuradas conforme suas necessidades específicas. Elas são ideais para automatizar tarefas repetitivas e criar uma experiência mais profissional e controlada para seus usuários finais.

--- 

> 📁 Este conteúdo pode ser salvo como `README.md` e usado como documentação interna ou pública para projetos baseados em VBA no Excel.