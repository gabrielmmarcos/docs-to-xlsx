from docx import Document
import pandas as pd

# Caminho do arquivo .docx
arquivo = "Valores Particulares exames laboratoriais.docx"

# Lê o documento Word
doc = Document(arquivo)

dados = []
# Percorre as tabelas do documento
for tabela in doc.tables:
    for i, linha in enumerate(tabela.rows):
        # Ignora cabeçalho
        if i == 0:
            continue

        celulas = [c.text.strip() for c in linha.cells]

        # Garante que tem 3 colunas
        if len(celulas) >= 3:
            exame = celulas[0]
            descricao = celulas[1]
            valor = celulas[2]

            # Ajusta o formato do valor (ex: $77.76 → R$ 77.76)
            valor = valor.replace("$", "R$ ").strip()

            dados.append({
                "Exame": exame,
                "Descrição": descricao,
                "Valor": valor
            })

# Cria o DataFrame
df = pd.DataFrame(dados)

# Exporta para Excel
df.to_excel("exames.xlsx", index=False)

print(f"✅ Arquivo 'exames.xlsx' criado com {len(df)} registros!")
