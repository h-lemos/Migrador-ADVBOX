import aspose.zip as az
import pandas as pd
import os
import re
import sys
import shutil
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox

# Setting path to templates
def resource_path(relative_path):
    """Get the absolute path to a resource, whether running as a script or as a bundled executable."""
    if hasattr(sys, "_MEIPASS"):
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Use the function to get the path to your templates
template_clientes_path = resource_path("templates\CLIENTES_template.xlsx")
template_processos_path = resource_path("templates\PROCESSOS_template.xlsx")

# Function to extract .rar file
def extract_archive(bdata, directory):
    if not os.path.exists(directory):
        os.makedirs(directory)
    try:
        with az.rar.RarArchive(bdata) as archive:
           archive.extract_to_directory(directory)
        print(f"Extraction completed successfully: {bdata}")
    except Exception as e:
        print(f"Error extracting archive: {e}")
        raise

# Function to read CSV files with multiple encoding attempts
def read_csv_file(file_path, encodings=('utf-8', 'ISO-8859-1'), sep=';'):
    for encoding in encodings:
        try:
            return pd.read_csv(file_path, encoding=encoding, sep=sep)
        except UnicodeDecodeError:
            continue
    print(f"Failed to read {file_path} with specified encodings.")
    return None

# Load CSV tables into a dictionary
def load_csv_tables(directory):
    return {
        os.path.splitext(file_name)[0]: read_csv_file(os.path.join(directory, file_name))
        for file_name in os.listdir(directory) if file_name.endswith('.csv')
    }

# Match a column header to a target field
def match_column(header, mapping):
    return next(
        (target for target, keywords in mapping.items() if any(re.search(keyword, header, re.IGNORECASE) for keyword in keywords)),
        None
    )

# Transform data based on specific rules
def transform_data(header, value, tables):
    if header == "DATA DE NASCIMENTO" and pd.notnull(value):
        match = re.search(r"\d{2}/\d{2}/\d{4}", str(value))
        return match.group(0) if match else None

    if header == "ORIGEM DO CLIENTE" and pd.notnull(value):
        lookup_table = tables.get("v_grupo_cliente_CodEmpresa_92577", pd.DataFrame())
        match_row = lookup_table[lookup_table["codigo"] == value]
        return match_row["descricao"].iloc[0] if not match_row.empty else "MIGRAÇÃO"
    
    if header == "NÚMERO DO PROCESSO" and pd.notnull(value):
        # Match the format "0000000-00.0000.0.00.0000"
        match = re.match(r"^\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}$", str(value))
        return value if match else None
    
    if header == "NOME DO CLIENTE" and pd.notnull(value) and type(int):
        lookup_table = tables.get("v_clientes_CodEmpresa_92577", pd.DataFrame())
        match_row = lookup_table[lookup_table["codigo"] == value]
        return match_row["razao_social"].iloc[0] if not match_row.empty else None
    
    if header == "COMARCA" and pd.notnull(value):
        lookup_table = tables.get("v_comarca_CodEmpresa_92577", pd.DataFrame())
        match_row = lookup_table[lookup_table["codigo"] == value]
        return match_row["descricao"].iloc[0] if not match_row.empty else None
    
    if header == "GRUPO DE AÇÃO" and pd.notnull(value):
        lookup_table = tables.get("v_grupo_processo_CodEmpresa_92577", pd.DataFrame())
        match_row = lookup_table[lookup_table["codigo"] == value]
        return match_row["descricao"].iloc[0] if not match_row.empty else None
    
    if header == "FASE PROCESSUAL" and pd.notnull(value):
        lookup_table = tables.get("v_fase_CodEmpresa_92577", pd.DataFrame())
        match_row = lookup_table[lookup_table["codigo"] == value]
        return match_row["fase"].iloc[0] if not match_row.empty else None
    
    if header == "DATA CADASTRO" and pd.notnull(value):
        match = re.search(r"\d{2}/\d{2}/\d{4}", str(value))
        return match.group(0) if match else None
    
    return value

# Extract data by key with transformations and obligatory checks
def extract_data_by_key(key_column, key_value, tables, mapping, target_headers, obligatory_columns):
    extracted_data = {header: None for header in target_headers}
    for df in tables.values():
        if key_column not in df.columns:
            continue
        table_rows = df[df[key_column] == key_value]
        if table_rows.empty:
            continue

        for column in table_rows.columns:
            target_header = match_column(column, mapping)
            if target_header and pd.notnull(table_rows[column].iloc[0]):
                value = transform_data(target_header, table_rows[column].iloc[0], tables)
                extracted_data[target_header] = value

    for col, default in obligatory_columns.items():
        extracted_data[col] = extracted_data.get(col) or default

    return extracted_data

# Merge duplicates by filling missing values
def merge_duplicates(df, key_column):
    return df.groupby(key_column, as_index=False).first()

# Constructor for the Clientes final table using a template
def constructor_clientes(tables):
    template_clients = pd.read_excel(template_clientes_path)
    template_headers = template_clients.columns.tolist()
    key_column = "razao_social"

    header_mapping = {
        "NOME": ["razao_social"],
        "CPF CNPJ": ["cpf" or "cnpj"],
        "RG": ["rg"],
        "NACIONALIDADE": ["nacionalidade"],
        "DATA DE NASCIMENTO": ["nascimento"],
        "ESTADO CIVIL": ["estado_civil"],
        "PROFISSÃO": ["profissao"],
        "CELULAR": ["telefone2"],
        "TELEFONE": ["telefone1", "telefone3", "telefone_comercial"],
        "EMAIL": ["email1", "email2"],
        "PAIS": ["nacionalidade"],
        "ESTADO": ["uf"],
        "CIDADE": ["cidade"],
        "BAIRRO": ["bairro"],
        "ENDEREÇO": ["logradouro"],
        "CEP": ["cep"],
        "PIS PASEP": ["pis"],
        "NOME DA MÃE": ["nome_mae"],
        "ORIGEM DO CLIENTE": ["grupo_cliente"]
    }

    obligatory_columns = {
        "ORIGEM DO CLIENTE": "MIGRAÇÃO",
        "NOME": "DESCONHECIDO"
    }

    clientes_table = "v_clientes_CodEmpresa_92577"
    if clientes_table not in tables:
        raise ValueError(f"Main table {clientes_table} not found.")

    main_table = tables[clientes_table]
    clientes = []

    for _, row in main_table.iterrows():
        if pd.isnull(row.get(key_column, None)) or not str(row[key_column]).strip():
            continue

        case_data = extract_data_by_key(key_column, row[key_column], tables, header_mapping, template_headers, obligatory_columns)
        clientes.append(case_data)

    final_clientes = pd.DataFrame(clientes, columns=template_headers)
    final_clientes = merge_duplicates(final_clientes, "NOME")
    return final_clientes

# Constructor for the Processos final table using a template
def constructor_processos(tables):
    template_processos = pd.read_excel(template_processos_path)
    template_headers = template_processos.columns.tolist()
    
    header_mapping = {
        "NOME DO CLIENTE": ["cod_cliente"],
        "PARTE CONTRÁRIA": ["parte_contraria"],
        "TIPO DE AÇÃO": ["tipo_acao"],
        "GRUPO DE AÇÃO": ["grupo_processo"],
        "FASE PROCESSUAL": ["codigo_fase"],
        "NÚMERO DO PROCESSO": ["numero_processo"],
        "PROCESSO ORIGINÁRIO": ["processo_originario"],
        "TRIBUNAL": ["tribunal"],
        "VARA": ["vara"],
        "COMARCA": ["comarca"],
        "PROTOCOLO": ["protocolo"],
        "EXPECTATIVA/VALOR DA CAUSA": ["expectativa_valor_causa"],
        "VALOR HONORÁRIOS": ["valor_honorarios"],
        "PASTA": ["pasta"],
        "DATA CADASTRO": ["inclusao"],
        "DATA FECHAMENTO": ["data_encerramento"],
        "DATA TRANSITO": ["data_transito"],
        "DATA ARQUIVAMENTO": ["data_arquivamento"],
        "DATA REQUERIMENTO": ["data_requerimento"],
        "RESPONSÁVEL": ["responsavel"],
        "ANOTAÇÕES GERAIS": ["anotacoes_gerais"]
    }
    
    obligatory_columns = {
        "NOME DO CLIENTE": "DESCONHECIDO",
        "TIPO DE AÇÃO": "DESCONHECIDO",
        "FASE PROCESSUAL": "DESCONHECIDO"
    }
    
    processos_table = "v_processos_CodEmpresa_92577"
    if processos_table not in tables:
        raise ValueError(f"Main table {processos_table} not found.")
    
    main_table = tables[processos_table]
    
    # Pre-filter rows to exclude invalid or empty "NÚMERO DO PROCESSO"
    valid_rows = main_table.dropna(subset=["numero_processo"]).copy()
    valid_rows = valid_rows[
        valid_rows["numero_processo"].str.match(r"^\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}$")
    ]
    
    processos = []
    
    for _, row in valid_rows.iterrows():
        case_data = extract_data_by_key(
            key_column="numero_processo", 
            key_value=row["numero_processo"], 
            tables=tables, 
            mapping=header_mapping, 
            target_headers=template_headers, 
            obligatory_columns=obligatory_columns
        )
        processos.append(case_data)
    
    final_processos = pd.DataFrame(processos, columns=template_headers)
    return final_processos


# Safely delete a folder
def delete_folder_safe(folder_path):
    try:
        if os.path.exists(folder_path):
            shutil.rmtree(folder_path)
            print(f"Deleted folder: {folder_path}")
    except Exception as e:
        print(f"Error deleting folder: {e}")

# Main processing function
def main(bdata):
    directory = f'./{os.path.splitext(os.path.basename(bdata))[0]}'
    try:
        # Extract and process data
        extract_archive(bdata, directory)
        tables = load_csv_tables(directory)

        # Build final Clientes table
        final_client_table = constructor_clientes(tables)
        final_client_table.to_excel("CLIENTES.xlsx", index=False)
        
        # Build final Cases table
        final_processos_table = constructor_processos(tables)
        final_processos_table.to_excel("PROCESSOS.xlsx", index=False)
        
        messagebox.showinfo("Success", f"Data migration of {os.path.basename(bdata)} completed successfully!")
        
    # Cleanup on error
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        delete_folder_safe(directory)
        raise
    finally:
        # Cleanup: Delete the directory and all its contents
        delete_folder_safe(directory)

# GUI Implementation
def gui():
    def select_file():
        file_path = filedialog.askopenfilename(
            title="Select Backup File",
            filetypes=[("RAR Files", "*.rar")]
        )
        file_entry.delete(0, "end")
        file_entry.insert(0, file_path)

    def start_process():
        bdata = file_entry.get()
        if not bdata:
            messagebox.showerror("Error", "Please select a valid .rar file!")
            return
        try:
            main(bdata)
            root.destroy()  # Close the GUI after successful migration
        except Exception as e:
            messagebox.showerror("Error", f"Processing failed. {str(e)}")

    global root  # Declare root as global if necessary
    root = Tk()
    root.title("Data Migration Tool")

    Label(root, text="Select the backup file (.rar):").grid(row=0, column=0, padx=10, pady=10)
    file_entry = Entry(root, width=50)
    file_entry.grid(row=1, column=0, padx=10, pady=5)
    Button(root, text="Browse", command=select_file).grid(row=1, column=1, padx=5, pady=5)
    Button(root, text="OK", command=start_process).grid(row=2, column=0, columnspan=2, pady=10)
    Button(root, text="Exit", command=root.quit).grid(row=3, column=0, columnspan=2, pady=10)

    root.mainloop()

# Run the GUI
if __name__ == "__main__":
    gui()
