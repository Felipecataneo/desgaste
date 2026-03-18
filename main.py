import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, timedelta

st.set_page_config(
    page_title="Processador de Operações de Perfuração",
    layout="wide"
)

st.title("Processador de Operações de Perfuração")
st.markdown(
    "Upload da tabela de entrada e geração da tabela agrupada em ordem cronológica"
)

uploaded_file = st.file_uploader(
    "Carregar tabela de entrada (Excel ou CSV)",
    type=["xlsx", "xls", "csv"]
)


def time_str_to_hours(time_str):
    """Convert time string in HH:MM:SS format to decimal hours"""
    if pd.isna(time_str) or time_str in ['None', '']:
        return 0.0

    if isinstance(time_str, (int, float)):
        return float(time_str)

    try:
        if isinstance(time_str, str):
            time_str = time_str.strip()
            parts = time_str.split(':')

            if len(parts) == 3:  # HH:MM:SS
                return int(parts[0]) + int(parts[1]) / 60 + int(parts[2]) / 3600
            elif len(parts) == 2:  # HH:MM
                return int(parts[0]) + int(parts[1]) / 60

            try:
                return float(time_str)
            except:
                pass

        return 0.0
    except:
        return 0.0


def parse_time(time_str):
    """Parse a time string into a datetime object"""
    if pd.isna(time_str) or time_str in ['None', '']:
        return None

    try:
        if isinstance(time_str, str):
            time_str = time_str.strip()

            if ':' in time_str:
                return datetime.strptime(time_str, '%H:%M:%S')
            else:
                if len(time_str) == 4:  # HHMM
                    return datetime.strptime(time_str, '%H%M')

        return None
    except:
        return None


def calculate_duration_from_times(inicio, fim):
    """Calculate duration between two time strings in hours"""
    start_time = parse_time(inicio)
    end_time = parse_time(fim)

    if start_time is None or end_time is None:
        return 0.0

    if end_time < start_time:
        end_time += timedelta(days=1)

    duration = (end_time - start_time).total_seconds() / 3600
    return duration


def format_duration(hours):
    """Format decimal hours back to HH:MM:SS string"""
    if pd.isna(hours) or hours == 0:
        return "00:00:00"

    total_seconds = int(hours * 3600)
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60

    return f"{h:02d}:{m:02d}:{s:02d}"


def process_data(df):
    df = df.copy()

    # 1. Padronização de Colunas
    if 'Operação.1' in df.columns:
        df['Operação_Desc'] = df['Operação.1']
    elif 'Operação_Desc' not in df.columns:
        df['Operação_Desc'] = ''

    for col in ['Topo', 'Base']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

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

    # 2. Cálculo de Duração
    if 'Início' in df.columns and 'Fim' in df.columns:
        df['Duração em horas'] = df.apply(
            lambda row: calculate_duration_from_times(
                row['Início'], row['Fim']
            ),
            axis=1
        )

        if 'Total' in df.columns:
            df['Duração do Total'] = df['Total'].apply(time_str_to_hours)

            zero_mask = (
                (df['Duração em horas'] == 0)
                & (df['Duração do Total'] > 0)
            )

            df.loc[zero_mask, 'Duração em horas'] = df.loc[
                zero_mask, 'Duração do Total'
            ]

    elif 'Total' in df.columns:
        df['Duração em horas'] = df['Total'].apply(time_str_to_hours)
    else:
        df['Duração em horas'] = 0

    # 3. Mapeamento de Operações
    operation_map = {
        'D': 'Desviando',
        'C': 'Circulação',
        'K': 'Corte de Cimento',
        'R': 'Perfuração',
        'B': 'Backreaming'
    }

    if 'Operação' in df.columns:
        df['Op_Code'] = (
            df['Operação']
            .astype(str)
            .str.strip()
            .str.upper()
            .str[0]
        )
    else:
        st.error("Coluna 'Operação' não encontrada na tabela.")
        return None

    # 4. Agrupamento Sequencial
    df['Block_ID'] = (
        (df['Op_Code'] != df['Op_Code'].shift(1))
        .cumsum()
    )

    block_summary = df.drop_duplicates('Block_ID').copy()
    block_summary['Op_Num'] = (
        block_summary.groupby('Op_Code')
        .cumcount() + 1
    )

    df = df.merge(
        block_summary[['Block_ID', 'Op_Num']],
        on='Block_ID',
        how='left'
    )

    result_data = []

    # 5. Iterar blocos
    for block_id in df['Block_ID'].unique():
        group_df = df[df['Block_ID'] == block_id]

        op_code = group_df['Op_Code'].iloc[0]
        op_num = group_df['Op_Num'].iloc[0]

        topo = group_df['Topo'].iloc[0]
        base = group_df['Base'].iloc[-1]

        total_duration = group_df['Duração em horas'].sum()

        if total_duration > 0:
            avg_wob = (
                (group_df['WOB'] * group_df['Duração em horas']).sum()
                / total_duration
            )
            avg_rpm = (
                (group_df['RPM'] * group_df['Duração em horas']).sum()
                / total_duration
            )
        else:
            avg_wob = group_df['WOB'].mean()
            avg_rpm = group_df['RPM'].mean()

        base_name = operation_map.get(op_code, op_code)
        op_name = f"{base_name} {op_num}"

        desc = ""
        if (
            'Operação_Desc' in group_df.columns
            and not pd.isna(group_df['Operação_Desc'].iloc[0])
        ):
            desc = group_df['Operação_Desc'].iloc[0]

        result_data.append({
            'Operação': op_name,
            'Descrição': desc,
            'Topo': topo,
            'Base': base,
            'WOB': round(avg_wob, 2),
            'RPM': round(avg_rpm, 2),
            'Duração': format_duration(total_duration),
            'Duração em horas': round(total_duration, 4)
        })

    return pd.DataFrame(result_data)


if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_input = pd.read_csv(uploaded_file)
        else:
            df_input = pd.read_excel(
                uploaded_file,
                dtype={'Início': str, 'Fim': str, 'Total': str}
            )

        st.subheader("Tabela de Entrada")
        st.dataframe(df_input)

        df_output = process_data(df_input)

        if df_output is not None:
            st.subheader("Tabela de Saída (Agrupamento Correto)")
            st.dataframe(df_output)

            buffer = io.BytesIO()

            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_output.to_excel(
                    writer,
                    index=False,
                    sheet_name='Operações Agrupadas'
                )

                workbook = writer.book
                worksheet = writer.sheets['Operações Agrupadas']
                format_num = workbook.add_format({'num_format': '#,##0.00'})

                for i, col in enumerate(df_output.columns):
                    column_width = max(
                        df_output[col].astype(str).map(len).max(),
                        len(col)
                    ) + 2

                    if col in ['Topo', 'Base', 'WOB', 'RPM', 'Duração em horas']:
                        worksheet.set_column(i, i, column_width, format_num)
                    else:
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
