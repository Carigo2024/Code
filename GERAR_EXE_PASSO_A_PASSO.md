# Passo a passo — gerar o `AtualizarDashboard.exe` (para a TI)

Objetivo: transformar o projeto num **único executável** que qualquer pessoa roda
com duplo-clique, **sem Python instalado**. Isto é feito **uma vez**. Tempo: ~15 min.

> Precisa ser feito numa máquina **Windows** (o executável de Windows só é gerado no
> Windows). Depois, o `.exe` roda em qualquer Windows.

---

## Pré-requisitos (uma vez, na máquina que vai gerar)

**1. Python 3.9 ou superior.**
- Baixe em https://www.python.org/downloads/ → botão "Download Python".
- Ao instalar, **MARQUE a caixa "Add Python to PATH"** (canto inferior da 1ª tela) →
  clique "Install Now".
- Para conferir: abra o **Prompt de Comando** (tecla Windows, digite `cmd`, Enter) e
  digite `python --version`. Deve aparecer algo como `Python 3.12.x`.
  - Se aparecer "python não é reconhecido": o Python não entrou no PATH — reinstale
    marcando a caixa "Add Python to PATH".

**2. Os arquivos do projeto** numa pasta, por exemplo `C:\ForecastBuild\`:
```
C:\ForecastBuild\
   forecast_accuracy.py
   dashboard_build.py
   dashboard_template.html
   atualizar.py
   atualizar.spec
   build_exe.bat
   data\
      logo.png            (opcional — o logo da empresa)
```
> Onde pegar: baixe do repositório (branch do projeto) — GitHub → botão verde
> "Code" → "Download ZIP" — e extraia esses arquivos para `C:\ForecastBuild\`.
> A pasta `data\` você cria; coloque o `logo.png` dentro (opcional).

---

## Gerar o executável

**3.** Abra a pasta `C:\ForecastBuild\` no Explorador de Arquivos.

**4.** Dê **duplo-clique** em **`build_exe.bat`**.
- Abre uma **janela preta** (Prompt de Comando). Ela vai, sozinha:
  1. instalar as dependências (`pyinstaller pandas numpy openpyxl pillow`);
  2. gerar o ícone a partir do `data\logo.png` (se existir);
  3. empacotar tudo com o PyInstaller.
- Demora **2 a 5 minutos**. Vão passar muitas linhas de texto — é normal.
- Ao terminar aparece: **`Pronto! O executavel esta em: dist\AtualizarDashboard.exe`**.
  Pressione uma tecla para fechar.

> Se preferir pelo Prompt: abra o `cmd`, digite `cd C:\ForecastBuild` (Enter) e
> depois `build_exe.bat` (Enter).

**5.** O executável está em **`C:\ForecastBuild\dist\AtualizarDashboard.exe`**.

---

## Montar a pasta de uso (para o dia a dia)

**6.** Crie a pasta que o time vai usar, por exemplo `C:\Forecast\`, assim:
```
C:\Forecast\
   AtualizarDashboard.exe      (copie de dist\)
   data\
      Order_01_2026.xlsx ... Order_07_2026.xlsx
      Received_itens_2026.xlsx
      logo.png                 (opcional)
```

**7.** Teste: **duplo-clique** em `AtualizarDashboard.exe`.
- Abre uma janela mostrando o progresso e a linha **`OK - reconciliacao com
  diferenca zero`**.
- Os resultados aparecem em **`C:\Forecast\saida\`**:
  - `forecast_accuracy.xlsx` (a planilha com o Painel)
  - `dashboard.html` (abre no navegador)

**Rotina mensal:** substitua/coloque o `Order_NN` novo em `data\` → duplo-clique no
`.exe` → pegue o `saida\forecast_accuracy.xlsx` novo. Pronto, sem Python.

---

## Solução de problemas

| Sintoma | Causa / correção |
|---|---|
| "python não é reconhecido" | Python não está no PATH → reinstale marcando "Add Python to PATH". |
| A janela fecha muito rápido | Abra pelo `cmd` (passo 4, nota) para ver a mensagem de erro. |
| Antivírus bloqueia / apaga o `.exe` | PyInstaller às vezes é sinalizado como falso-positivo → libere a pasta `dist\` no antivírus e gere de novo. |
| Erro "ModuleNotFoundError: X" ao rodar o `.exe` | abra `atualizar.spec`, adicione `X` em `hiddenimports=[...]` e rode `build_exe.bat` de novo. |
| "Pasta 'data' nao encontrada" ao rodar o `.exe` | o `.exe` precisa de uma subpasta `data\` **ao lado dele** (passo 6). |

> Tamanho do `.exe`: ~60–120 MB (embute pandas/numpy). É normal.
