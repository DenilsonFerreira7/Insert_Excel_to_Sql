import pandas as pd
from unidecode import unidecode


excel_file = 'C:/Users/Denilson Stratus tel/Downloads/AlunosPreMatriculados_Ricardo.xlsx'
df = pd.read_excel(excel_file)


table_name = 'discadora'


df.columns = df.columns.str.lower()


df.fillna('', inplace=True)

# Colunas da tabela SQL
columns = [
    'ra', 'nome', 'email', 'telefone1', 'telefone2',
    'cidade', 'estado', 'codcurso', 'curso', 'modalidade',
    'codcampus', 'polo'
]


columns = [unidecode(column) for column in columns]


values_statements = []
error_occurred = False
for index, row in df.iterrows():
    try:
        values = [repr(unidecode(str(row[column]))) for column in columns]  # Aplica unidecode aos valores
        values_statement = f"({', '.join(values)})"
        values_statements.append(values_statement)
    except Exception as e:
        print(f"Erro ao processar linha {index + 2}: {e}")
        error_occurred = True
        break

if not error_occurred:

    insert_statements = []
    for values_statement in values_statements:
        insert_statements.append(f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES {values_statement}")


    sql_script = "START TRANSACTION;\n\n"
    sql_script += "\n;\n".join(insert_statements)
    sql_script += ";\n\nCOMMIT;"


    sql_file = 'insert_script.sql'
    with open(sql_file, 'w') as file:
        file.write(sql_script)

    print(f"Instrução SQL gerada e salva em {sql_file}")
else:
    print("Ocorreu um erro durante a geração dos valores SQL. Nenhum arquivo foi gerado.")
