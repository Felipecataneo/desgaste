import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Drilling Operations Processor", layout="wide")

st.title("Processador de Operações de Perfuração")
st.markdown("Upload da tabela de entrada e geração da tabela agrupada por operações")

uploaded_file = st.file_uploader("Carregar tabela de entrada (Excel)", type=["xlsx", "xls"])

def process_data(df):
    # Ensure column names match expected input format
    expected_columns = ['Operação', 'Topo', 'Base', 'PSB', 'RPM', 'Total']
    
    # Check if all expected columns exist
    for col in expected_columns:
        if col not in df.columns:
            st.error(f"Coluna '{col}' não encontrada. Verifique o formato da tabela de entrada.")
            return None
    
    # Select only needed columns and convert to appropriate types
    df = df[expected_columns].copy()
    
    # Convert numeric columns to appropriate types
    df['Topo'] = pd.to_numeric(df['Topo'], errors='coerce')
    df['Base'] = pd.to_numeric(df['Base'], errors='coerce')
    df['PSB'] = pd.to_numeric(df['PSB'], errors='coerce')
    df['RPM'] = pd.to_numeric(df['RPM'], errors='coerce')
    
    # Convert time column to float hours
    def convert_time_to_hours(time_str):
        if pd.isna(time_str):
            return np.nan
        
        # Check if it's already a float or integer
        if isinstance(time_str, (float, int)):
            return float(time_str)
        
        try:
            # Split by ':' for HH:MM format
            if ':' in time_str:
                parts = time_str.split(':')
                if len(parts) == 2:
                    hours = float(parts[0])
                    minutes = float(parts[1])
                    return hours + (minutes / 60)
            # Try direct conversion for decimal hours
            return float(time_str)
        except:
            return np.nan
    
    df['Duração em horas'] = df['Total'].apply(convert_time_to_hours)
    
    # Create a new column to track operation groups
    df['Grupo Operação'] = ''
    current_operation = ''
    operation_count = {}
    
    # Assign operation groups
    for i, row in df.iterrows():
        op = row['Operação']
        
        if op != current_operation:
            # New operation type encountered
            current_operation = op
            if op in operation_count:
                operation_count[op] += 1
            else:
                operation_count[op] = 1
            
            group_name = f"{op}{operation_count[op]}"
        else:
            # Same operation continues
            group_name = f"{op}{operation_count[op]}"
        
        df.at[i, 'Grupo Operação'] = group_name
    
    # Group by operation group and calculate statistics
    result = pd.DataFrame(columns=['Operação', 'Topo', 'Base', 'WOB', 'RPM', 'Duração', 'Duração em horas'])
    
    for group_name, group_data in df.groupby('Grupo Operação'):
        min_topo = group_data['Topo'].min()
        max_base = group_data['Base'].max()
        avg_psb = group_data['PSB'].mean()
        avg_rpm = group_data['RPM'].mean()
        total_duration_hours = group_data['Duração em horas'].sum()
        
        # Format duration in HH:MM:SS format
        hours = int(total_duration_hours)
        minutes = int((total_duration_hours - hours) * 60)
        seconds = int(((total_duration_hours - hours) * 60 - minutes) * 60)
        
        duration_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
        new_row = {
            'Operação': group_name,
            'Topo': min_topo,
            'Base': max_base,
            'WOB': avg_psb,
            'RPM': avg_rpm,
            'Duração': duration_str,
            'Duração em horas': total_duration_hours
        }
        
        result = pd.concat([result, pd.DataFrame([new_row])], ignore_index=True)
    
    return result

if uploaded_file is not None:
    try:
        # Read the Excel file
        df_input = pd.read_excel(uploaded_file)
        
        st.subheader("Tabela de Entrada")
        st.dataframe(df_input)
        
        # Process the data
        df_output = process_data(df_input)
        
        if df_output is not None:
            st.subheader("Tabela de Saída")
            st.dataframe(df_output)
            
            # Create a download button for the processed data
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
else:
    st.info("Carregue um arquivo Excel para começar.")
    
    # Example table format
    st.subheader("Formato Esperado da Tabela de Entrada")
    example_data = {
        'Operação': ['Circulação', 'Corte de Cimento', 'Perfuração', 'Perfuração'],
        'Topo': [5514, 5585, 5595, 5779],
        'Base': [5514, 5585, 5779, 5819],
        'PSB': [30, 20, 25, 28],
        'RPM': [50, 48, 109, 110],
        'Total': ['00:15', '17:08', '31:02', '05:37']
    }
    st.dataframe(pd.DataFrame(example_data))