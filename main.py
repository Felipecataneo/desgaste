import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, timedelta

st.set_page_config(page_title="Desgaste de Revestimento", layout="wide")

st.title("Processador de Operações de Perfuração")
st.markdown("Upload da tabela de entrada e geração da tabela agrupada por operações")

uploaded_file = st.file_uploader("Carregar tabela de entrada (Excel)", type=["xlsx", "xls", "csv"])

def time_str_to_hours(time_str):
    """Convert time string in HH:MM:SS format to decimal hours"""
    if pd.isna(time_str) or time_str == 'None':
        return 0.0
   
    if isinstance(time_str, (int, float)):
        return float(time_str)
       
    try:
        # Handle variations in time format
        if isinstance(time_str, str):
            parts = time_str.split(':')
            if len(parts) == 3:  # HH:MM:SS
                return int(parts[0]) + int(parts[1])/60 + int(parts[2])/3600
            elif len(parts) == 2:  # HH:MM
                return int(parts[0]) + int(parts[1])/60
        return 0.0
    except:
        return 0.0

def process_data(df):
    # Ensure required columns exist
    required_columns = ['Topo', 'Base', 'Operação', 'PSB', 'RPM', 'Total']
   
    # Map column variations to standard names
    column_mapping = {
        'Operação.1': 'Operação_Desc',
        'WOB': 'PSB',
    }
   
    # Rename columns if present
    for old_col, new_col in column_mapping.items():
        if old_col in df.columns:
            df[new_col] = df[old_col]
   
    # Check columns and setup proper structure
    if 'Operação' not in df.columns:
        if 'Operação.1' in df.columns:
            df['Operação'] = df['Operação.1']
        else:
            st.error("Coluna 'Operação' não encontrada na tabela.")
            return None
   
    # Ensure operation abbreviation is present (C, K, R, B)
    if df['Operação'].dtype == 'object' and all(len(str(op)) > 1 for op in df['Operação'].dropna()):
        # If we have full operation names like "Circulação" but no codes
        # Try to derive operation codes from operation names
        operation_map = {
            'Circulação': 'C',
            'Corte de Cimento': 'K',
            'Perfuração': 'R',
            'Backreaming': 'B'
        }
       
        # Try to extract code from operation name
        if 'Operação_Desc' not in df.columns:
            df['Operação_Desc'] = df['Operação']
           
        # Check if we need to create operation codes
        if not all(op in ['C', 'K', 'R', 'B'] for op in df['Operação'].dropna()):
            # Try to map operation descriptions to codes
            for i, row in df.iterrows():
                for key, code in operation_map.items():
                    if str(row['Operação_Desc']).startswith(key):
                        df.at[i, 'Operação'] = code
                        break
   
    # Ensure numeric columns are properly typed
    numeric_cols = ['Topo', 'Base', 'PSB', 'RPM']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        else:
            df[col] = np.nan
   
    # Convert time columns
    if 'Total' in df.columns:
        df['Duração em horas'] = df['Total'].apply(time_str_to_hours)
    else:
        df['Duração em horas'] = 0.0
   
    # Fill missing values with reasonable defaults
    df['Topo'].fillna(method='ffill', inplace=True)
    df['Base'].fillna(method='ffill', inplace=True)
    df['PSB'].fillna(0, inplace=True)
    df['RPM'].fillna(0, inplace=True)
   
    # Group operations by abbreviation and create operation groups
    # Map operation codes to full names
    operation_names = {
        'C': 'Circulação',
        'K': 'Corte de Cimento',
        'R': 'Perfuração',
        'B': 'Backreaming'
    }
   
    # Initialize operation counters
    operation_counters = {'C': 0, 'K': 0, 'R': 0, 'B': 0}
    current_op = None
    df['Grupo'] = ''
   
    # Assign groups
    for i, row in df.iterrows():
        op_code = row['Operação']
        if pd.isna(op_code) or op_code == 'None':
            continue
           
        # Convert to string to handle possible numeric codes
        op_code = str(op_code).strip()
       
        # Only consider the first character if it's a longer string
        if len(op_code) > 1:
            # Check if first char is a valid operation code
            if op_code[0] in operation_counters:
                op_code = op_code[0]
            else:
                # Try to determine from operation description
                if 'Circulação' in str(row.get('Operação_Desc', '')):
                    op_code = 'C'
                elif 'Cimento' in str(row.get('Operação_Desc', '')):
                    op_code = 'K'
                elif 'Perfuração' in str(row.get('Operação_Desc', '')):
                    op_code = 'R'
                elif 'Backreaming' in str(row.get('Operação_Desc', '')):
                    op_code = 'B'
       
        # Skip if operation code is not one of our expected codes
        if op_code not in operation_counters:
            continue
           
        # If operation changed, increment counter
        if op_code != current_op:
            operation_counters[op_code] += 1
            current_op = op_code
           
        # Create group name
        group = f"{op_code}{operation_counters[op_code]}"
        df.at[i, 'Grupo'] = group
   
    # Group and calculate statistics
    result_data = []
   
    for group, group_df in df[df['Grupo'] != ''].groupby('Grupo'):
        op_code = group[0]  # First character is the operation code
        op_num = group[1:]  # Remaining characters are the number
       
        # Calculate statistics
        min_topo = group_df['Topo'].min()
        max_base = group_df['Base'].max()
        avg_psb = group_df['PSB'].mean()
        avg_rpm = group_df['RPM'].mean()
        total_duration = group_df['Duração em horas'].sum()
       
        # Format duration
        hours = int(total_duration)
        minutes = int((total_duration - hours) * 60)
        seconds = int(((total_duration - hours) * 60 - minutes) * 60)
        duration_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
       
        # Get operation name
        op_name = f"{operation_names.get(op_code, op_code)}{op_num}"
       
        result_data.append({
            'Operação': op_name,
            'Topo': min_topo,
            'Base': max_base,
            'WOB': round(avg_psb, 3),
            'RPM': round(avg_rpm, 2),
            'Duração': duration_str,
            'Duração em horas': round(total_duration, 6)
        })
   
    # Create result DataFrame
    result_df = pd.DataFrame(result_data)
   
    # Sort by operation
    def get_sort_key(op_str):
        # Extract operation code and number for sorting
        code = op_str[0] if op_str else ''
        # Set priority for sorting (C, K, R, B)
        code_priority = {'C': 0, 'K': 1, 'R': 2, 'B': 3}.get(code, 4)
        # Extract number portion
        num_part = ''.join(filter(str.isdigit, op_str))
        num = int(num_part) if num_part else 0
        return (code_priority, num)
   
    # Sort result by operation code and number
    result_df['sort_key'] = result_df['Operação'].apply(lambda x: get_sort_key(x))
    result_df = result_df.sort_values('sort_key').drop('sort_key', axis=1)
   
    return result_df

if uploaded_file is not None:
    try:
        # Try to determine file type and read accordingly
        if uploaded_file.name.endswith('.csv'):
            df_input = pd.read_csv(uploaded_file)
        else:
            df_input = pd.read_excel(uploaded_file)
       
        st.subheader("Tabela de Entrada")
        st.dataframe(df_input)
       
        # Process the data
        df_output = process_data(df_input)
       
        if df_output is not None:
            st.subheader("Tabela de Saída")
            st.dataframe(df_output)
           
            # Create download button
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_output.to_excel(writer, index=False, sheet_name='Operações Agrupadas')
               
                # Auto-adjust columns' width
                workbook = writer.book
                worksheet = writer.sheets['Operações Agrupadas']
                for i, col in enumerate(df_output.columns):
                    column_width = max(df_output[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, column_width)
           
            buffer.seek(0)
            st.download_button(
                label="Baixar Tabela Processada (Excel)",
                data=buffer,
                file_name="operacoes_agrupadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
           
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        st.exception(e)
else:
    st.info("Carregue um arquivo Excel ou CSV para começar.")
   
    # Show example format
    st.subheader("Formato da Tabela de Entrada")
    example_data = {
        'Início': ['13:15:00', '13:30:00', '18:00:00'],
        'Fim': ['13:30:00', '18:00:00', '19:45:00'],
        'Total': ['00:15:00', '04:30:00', '01:45:00'],
        'Topo': [None, 5514, 5517],
        'Base': [5514, 5517, 5518],
        'Operação': ['C', 'K', 'K'],
        'PSB': [None, 30, 25],
        'RPM': [50, 50, 50],
        'Operação.1': ['Circulação', 'Corte de Cimento', 'Corte de Cimento']
    }
    st.dataframe(pd.DataFrame(example_data))
   
    st.subheader("Formato da Tabela de Saída")
    output_example = {
        'Operação': ['Circulação1', 'Corte de Cimento1', 'Perfuração1'],
        'Topo': [None, 5514, 5585],
        'Base': [5514, 5585, 5779],
        'WOB': [0, 17.125, 22],
        'RPM': [50, 47.5, 108.75],
        'Duração': ['00:15:00', '17:29:00', '31:02:00'],
        'Duração em horas': [0.25, 17.483, 31.033]
    }
    st.dataframe(pd.DataFrame(output_example))