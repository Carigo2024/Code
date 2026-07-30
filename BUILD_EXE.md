# Gerar o executável (.exe) — atualizar o dashboard sem Python

Empacota o motor + o gerador do dashboard num **único `.exe`**. Quem tem Python
gera **uma vez**; depois o `.exe` roda em qualquer Windows **sem Python instalado**.

## Quem gera (uma vez)

Numa máquina **Windows** com Python 3.9+ e os arquivos do projeto
(`atualizar.py`, `forecast_accuracy.py`, `dashboard_build.py`,
`dashboard_template.html`, `atualizar.spec`, `build_exe.bat`):

```bat
build_exe.bat
```

O script instala as dependências (`pyinstaller pandas numpy openpyxl pillow`),
gera um ícone a partir de `data\logo.png` (se existir) e produz:

```
dist\AtualizarDashboard.exe
```

> **Importante:** o PyInstaller **não faz cross-compile**. Para um `.exe` de
> Windows, rode o build **no Windows**. (No Linux/macOS ele gera um binário
> daquele sistema.) O `.exe` costuma ter ~60–120 MB por embutir pandas/numpy —
> isso é normal.

## Quem usa (todo mês, sem Python)

1. Coloque `AtualizarDashboard.exe` numa pasta com uma subpasta **`data`**:
   ```
   MinhaPasta\
     AtualizarDashboard.exe
     data\
       Order_01_2026.xlsx ... Order_07_2026.xlsx
       Received_itens_2026.xlsx
       logo.png            (opcional)
   ```
2. **Duplo-clique** no `.exe`. Uma janela mostra o progresso e a linha de
   **reconciliação** (`DIF 0 OK`).
3. Pegue os resultados na subpasta **`saida`**:
   - `dashboard.html` (abrir no navegador / mandar para gestores)
   - `forecast_accuracy.xlsx`

Atualização mensal = jogar o novo `Order_NN` em `data\` e rodar o `.exe`.

## Como funciona (por dentro)

- `atualizar.py` detecta se está empacotado (`sys.frozen`) e resolve as pastas
  `data`/`saida` ao lado do `.exe` (não do diretório atual), e lê o
  `dashboard_template.html` de dentro do pacote (`sys._MEIPASS`).
- `atualizar.spec` embute o template e exclui libs pesadas não usadas
  (matplotlib, tkinter, scipy…) para reduzir o tamanho.

## Se faltar algum módulo no .exe

Raro, mas se o `.exe` reclamar de um módulo ausente, adicione-o em
`hiddenimports` no `atualizar.spec` e rode `build_exe.bat` de novo. Ex.:

```python
hiddenimports=["openpyxl.cell._writer", "pandas._libs.tslibs.timedeltas"],
```
