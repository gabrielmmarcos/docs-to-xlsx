# Converter DOCX → Excel (exames)

Script simples em Python para ler um arquivo `.docx` com uma tabela de exames e exportar para um arquivo Excel (`exames.xlsx`) com as colunas **Exame**, **Descrição** e **Valor**.

## Requisitos

- Python 3.8+ (recomendado)  
- VS Code (opcional)  
- O arquivo `.docx` deve estar na mesma pasta do script (ou você deve ajustar o caminho no script)

Bibliotecas Python utilizadas:
- `python-docx`  
- `pandas`  
- `openpyxl`

## Instalação rápida

Abra o terminal na pasta do projeto e rode:

```bash
# opcional: criar e ativar virtualenv
python -m venv .venv

# Linux / macOS
source .venv/bin/activate

# Windows (PowerShell)
.venv\Scripts\Activate.ps1

# instalar dependências
pip install python-docx pandas openpyxl
```

Se preferir, você pode instalar tudo de uma vez com:

pip install -r requirements.txt


## Como usar

1. Coloque o arquivo Word (exemplo: `Valores Particulares exames laboratoriais.docx`) na mesma pasta do `index.py`, ou ajuste o valor da variável `arquivo` no script para o caminho correto.  
2. Execute o script:

python index.py

3. O script vai gerar um arquivo `exames.xlsx` na mesma pasta com as colunas:
   - `Exame`
   - `Descrição`
   - `Valor` (o script converte `$` para `R$` quando encontra)

No final, o script imprime quantos registros foram processados.

## Estrutura esperada do `.docx`

- O ideal é que o `.docx` contenha uma **tabela** com pelo menos 3 colunas (Exame, Descrição e Valor).  
- Se o `.docx` não tiver tabela real (apenas colunas alinhadas por espaços), o script atual pode não extrair corretamente. Nesse caso, é possível adaptar o código para ler o texto e separar as colunas usando regex.

