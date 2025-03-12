import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, timedelta

st.set_page_config(page_title="Processador de Operações de Perfuração", layout="wide")

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

def format_duration(hours):
    """Format hours as HH:MM:SS"""
    total_seconds = int(hours * 3600)
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def process_data(df):
    # Make a copy to avoid modifying the original
    df = df.copy()
    
    # Map column variations to standard names
    if 'Operação.1' in df.columns:
        df['Operação_Desc'] = df['Operação.1']
        
    # Ensure required columns exist and convert to proper types
    # For Topo and Base
    for col in ['Topo', 'Base']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        else:
            df[col] = np.nan
    
    # For PSB/WOB and RPM
    if 'PSB' in df.columns:
        df['WOB'] = pd.to_numeric(df['PSB'], errors='coerce').fillna(0)
    elif 'WOB' in df.columns:
        df['WOB'] = pd.to_numeric(df['WOB'], errors='coerce').fillna(0)
    else:
        df['WOB'] = 0
        
    if 'RPM' in df.columns:
        df['RPM'] = pd.to_numeric(df['RPM'], errors='coerce').fillna(0)
    else:
        df['RPM'] = 0
    
    # For duration calculation
    if 'Total' in df.columns:
        df['Duração em horas'] = df['Total'].apply(time_str_to_hours)
    else:
        # Try to calculate from start and end times if available
        if 'Início' in df.columns and 'Fim' in df.columns:
            def calculate_duration(row):
                try:
                    # Parse start and end times
                    if isinstance(row['Início'], str) and isinstance(row['Fim'], str):
                        format_str = '%H:%M:%S' if ':' in row['Início'] else '%H%M'
                        
                        start = datetime.strptime(row['Início'], format_str)
                        end = datetime.strptime(row['Fim'], format_str)
                        
                        # Handle crossing midnight
                        if end < start:
                            end += timedelta(days=1)
                            
                        # Calculate duration in hours
                        duration = (end - start).total_seconds() / 3600
                        return duration
                    return 0
                except:
                    return 0
                    
            df['Duração em horas'] = df.apply(calculate_duration, axis=1)
        else:
            df['Duração em horas'] = 0
    
    # Determine operation codes and descriptions
    if 'Operação' not in df.columns:
        st.error("Coluna 'Operação' não encontrada na tabela.")
        return None
        
    # Map abbreviations to full operation names
    operation_map = {
        'C': 'Circulação',
        'K': 'Corte de Cimento',
        'R': 'Perfuração',
        'B': 'Backreaming'
    }
    
    # Create reverse mapping for operation descriptions
    reverse_map = {
        'Circulação': 'C',
        'Corte de Cimento': 'K', 
        'Perfuração': 'R',
        'Backreaming': 'B'
    }
    
    # Check for operation codes
    for i, row in df.iterrows():
        op = str(row['Operação']).strip() if not pd.isna(row['Operação']) else ''
        
        # Handle abbreviated operation codes
        if op in operation_map:
            df.at[i, 'Op_Code'] = op
            if 'Operação_Desc' not in df.columns or pd.isna(row.get('Operação_Desc')):
                df.at[i, 'Operação_Desc'] = operation_map[op]
        # Handle operation descriptions
        elif 'Operação_Desc' in df.columns and not pd.isna(row['Operação_Desc']):
            desc = str(row['Operação_Desc']).strip()
            for key, code in reverse_map.items():
                if desc.startswith(key):
                    df.at[i, 'Op_Code'] = code
                    break
            else:
                # If no match found in reverse_map, try to extract first letter
                if op and len(op) > 0 and op[0] in operation_map:
                    df.at[i, 'Op_Code'] = op[0]
                else:
                    df.at[i, 'Op_Code'] = ''
        else:
            # Try to determine from operation string if it's longer
            if len(op) > 1:
                if op[0] in operation_map:
                    df.at[i, 'Op_Code'] = op[0]
                else:
                    for key, code in reverse_map.items():
                        if op.startswith(key):
                            df.at[i, 'Op_Code'] = code
                            break
                    else:
                        df.at[i, 'Op_Code'] = ''
            else:
                df.at[i, 'Op_Code'] = ''
    
    # Group the operations
    operation_counters = {'C': 0, 'K': 0, 'R': 0, 'B': 0}
    current_op = None
    df['Grupo'] = ''
    
    for i, row in df.iterrows():
        op_code = row['Op_Code']
        if not op_code:
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
        
        # Calculate WOB and RPM properly - using weighted average by duration
        total_duration = group_df['Duração em horas'].sum()
        
        if total_duration > 0:
            # Weighted average for WOB and RPM
            avg_wob = (group_df['WOB'] * group_df['Duração em horas']).sum() / total_duration
            avg_rpm = (group_df['RPM'] * group_df['Duração em horas']).sum() / total_duration
        else:
            # Simple average if no duration
            avg_wob = group_df['WOB'].mean()
            avg_rpm = group_df['RPM'].mean()
        
        # Format duration string
        duration_str = format_duration(total_duration)
        
        # Get operation name
        op_name = f"{operation_map.get(op_code, op_code)}{op_num}"
        
        result_data.append({
            'Operação': op_name,
            'Topo': min_topo,
            'Base': max_base,
            'WOB': round(avg_wob, 3),
            'RPM': round(avg_rpm, 2),
            'Duração': duration_str,
            'Duração em horas': round(total_duration, 6)
        })
    
    # Create result DataFrame
    result_df = pd.DataFrame(result_data)
    
    # Sort by operation code and number
    operation_order = {'Circulação': 0, 'Corte de Cimento': 1, 'Perfuração': 2, 'Backreaming': 3}
    
    def get_sort_key(op_str):
        # Find the operation type
        for op_type in operation_order:
            if op_str.startswith(op_type):
                prefix = op_type
                break
        else:
            prefix = op_str
            
        # Extract number portion
        num_part = ''.join(filter(str.isdigit, op_str))
        num = int(num_part) if num_part else 0
        
        # Get the priority
        priority = operation_order.get(prefix, 99)
        
        return (priority, num)
    
    # Sort result
    result_df['sort_key'] = result_df['Operação'].apply(lambda x: get_sort_key(x))
    result_df = result_df.sort_values('sort_key').drop('sort_key', axis=1)
    
    return result_df

if uploaded_file is not None:
    try:
        # Determine file type and read accordingly
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
    
    # Show expected formats
    st.subheader("Exemplo da Tabela de Entrada")
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