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
    if pd.isna(time_str) or time_str == 'None' or time_str == '':
        return 0.0
   
    if isinstance(time_str, (int, float)):
        return float(time_str)
       
    try:
        # Handle variations in time format
        if isinstance(time_str, str):
            # Clean the string
            time_str = time_str.strip()
           
            # Try parsing with different formats
            parts = time_str.split(':')
            if len(parts) == 3:  # HH:MM:SS
                return int(parts[0]) + int(parts[1])/60 + int(parts[2])/3600
            elif len(parts) == 2:  # HH:MM
                return int(parts[0]) + int(parts[1])/60
           
            # Try direct conversion if it looks like hours
            try:
                return float(time_str)
            except:
                pass
        return 0.0
    except:
        return 0.0

def parse_time(time_str):
    """Parse a time string into a datetime object"""
    if pd.isna(time_str) or time_str == 'None' or time_str == '':
        return None
       
    try:
        if isinstance(time_str, str):
            time_str = time_str.strip()
           
            # Try different formats
            if ':' in time_str:
                return datetime.strptime(time_str, '%H:%M:%S')
            else:
                # Handle other formats
                if len(time_str) == 4:  # HHMM
                    return datetime.strptime(time_str, '%H%M')
        return None
    except:
        return None

def format_duration(hours):
    """Format hours as HH:MM:SS"""
    if hours == 0:
        return "00:00:00"
       
    total_seconds = int(hours * 3600)
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def calculate_duration_from_times(inicio, fim):
    """Calculate duration between two time strings in hours"""
    start_time = parse_time(inicio)
    end_time = parse_time(fim)
   
    if start_time is None or end_time is None:
        return 0.0
       
    # Handle crossing midnight
    if end_time < start_time:
        end_time += timedelta(days=1)
       
    # Calculate duration in hours
    duration = (end_time - start_time).total_seconds() / 3600
    return duration

def process_data(df):
    # Make a copy to avoid modifying the original
    df = df.copy()
   
    # Print columns for debugging
    st.write("Colunas originais:", df.columns.tolist())
   
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
   
    # For duration calculation - ALWAYS calculate from Início and Fim columns
    if 'Início' in df.columns and 'Fim' in df.columns:
        # Calculate durations from start and end times
        df['Duração em horas'] = df.apply(
            lambda row: calculate_duration_from_times(row['Início'], row['Fim']),
            axis=1
        )
       
        # Also try from Total column as backup if available
        if 'Total' in df.columns:
            df['Duração do Total'] = df['Total'].apply(time_str_to_hours)
           
            # If the calculation from Início/Fim is 0 but Total has a value, use Total instead
            zero_mask = (df['Duração em horas'] == 0) & (df['Duração do Total'] > 0)
            df.loc[zero_mask, 'Duração em horas'] = df.loc[zero_mask, 'Duração do Total']
           
            # Display sample of duration calculation for comparison
            st.write("Comparação de cálculos de duração:")
            st.write(df[['Início', 'Fim', 'Total', 'Duração em horas', 'Duração do Total']].head(10))
    elif 'Total' in df.columns:
        # Fallback to Total if Início/Fim are not available
        df['Duração em horas'] = df['Total'].apply(time_str_to_hours)
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
   
    # Display sample of grouped data with durations
    st.write("Dados agrupados com durações:")
    st.write(df[['Operação', 'Op_Code', 'Grupo', 'Duração em horas']].head(10))
   
    # Group and calculate statistics
    result_data = []
    
    # Armazenar a primeira linha de cada grupo para ordenação cronológica
    first_rows_by_group = {}
    for group, group_df in df[df['Grupo'] != ''].groupby('Grupo'):
        if not group_df.empty:
            first_rows_by_group[group] = group_df.iloc[0]
   
    # Group and calculate statistics
    for group, group_df in df[df['Grupo'] != ''].groupby('Grupo'):
        op_code = group[0]  # First character is the operation code
        op_num = group[1:]  # Remaining characters are the number
       
        # Calculate statistics
        min_topo = group_df['Topo'].min()
        max_base = group_df['Base'].max()
        
        # Obter o timestamp do início da primeira operação no grupo para ordenação
        first_time_str = None
        if group in first_rows_by_group and 'Início' in first_rows_by_group[group]:
            first_time_str = first_rows_by_group[group]['Início']
       
        # Sum durations for this group
        total_duration = group_df['Duração em horas'].sum()
       
        # Display for debugging
        st.write(f"Grupo {group}: Duração total = {total_duration} horas")
        st.write(group_df[['Duração em horas']].sum())
       
        # Store group stats for debugging
        group_stats = {
            "total_duration": total_duration,
            "num_rows": len(group_df),
            "durations": group_df['Duração em horas'].tolist(),
            "first_time": first_time_str
        }
        
        st.write(f"Grupo {group} stats:", group_stats)
       
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
            'Duração em horas': round(total_duration, 6),
            'Op_Code': op_code,
            'Op_Num': int(op_num) if op_num else 0,
            'First_Time': first_time_str  # Armazenar para ordenação
        })
   
    # Create result DataFrame
    result_df = pd.DataFrame(result_data)

    # Abordagem de ordenação combinada:
    # 1. Ordenar primeiro por profundidade (Topo) - operações mais rasas primeiro
    # 2. Para operações na mesma profundidade, ordenar pelo número da operação
    # 3. Se os dois critérios falharem, usar o timestamp como fallback

    # Garantir que Op_Num seja numérico
    result_df['Op_Num'] = pd.to_numeric(result_df['Op_Num'], errors='coerce').fillna(0).astype(int)

    # Converter a coluna First_Time para um formato que possa ser ordenado
    result_df['Timestamp'] = pd.to_datetime(result_df['First_Time'], format='%H:%M:%S', errors='coerce')

    # Ordenação em várias etapas
    result_df = result_df.sort_values(by=['Topo', 'Op_Num', 'Timestamp'], ascending=[True, True, True])

    # Alternativa: Criar uma ordem personalizada baseada no tipo de operação
    operation_order = {'C': 1, 'R': 2, 'K': 3, 'B': 4}  # Circulação, Perfuração, Corte de Cimento, Backreaming
    result_df['Op_Order'] = result_df['Op_Code'].map(operation_order).fillna(99)

    # Ordenar por Topo e depois pela ordem de operação e número
    result_df = result_df.sort_values(by=['Topo', 'Op_Order', 'Op_Num'], ascending=[True, True, True])

    # Remover colunas auxiliares antes de retornar
    if 'Op_Code' in result_df.columns:
        result_df = result_df.drop(['Op_Code', 'Op_Num', 'First_Time', 'Timestamp', 'Op_Order'], axis=1)

    return result_df

if uploaded_file is not None:
    try:
        # Determine file type and read accordingly
        if uploaded_file.name.endswith('.csv'):
            df_input = pd.read_csv(uploaded_file)
        else:
            # Force reading types correctly for time columns
            df_input = pd.read_excel(
                uploaded_file,
                dtype={
                    'Início': str,
                    'Fim': str,
                    'Total': str
                }
            )
       
        st.subheader("Tabela de Entrada")
        st.dataframe(df_input)
       
        # Process the data
        with st.expander("Detalhes do processamento (expandir para debug)"):
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