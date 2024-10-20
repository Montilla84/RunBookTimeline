import pandas as pd
import plotly.figure_factory as ff
from datetime import datetime
import plotly.io as pio
from random import randint
import numpy as np
import os

def try_read_csv(file_path):
    """
    Intenta leer un archivo CSV con diferentes codificaciones.
    Retorna el DataFrame si tiene éxito.
    """
    encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']
    
    for encoding in encodings:
        try:
            print(f"Intentando leer con codificación: {encoding}")
            # Especificamos el orden exacto de las columnas
            df = pd.read_csv(file_path, encoding=encoding)
            print(f"Éxito al leer con codificación: {encoding}")
            return df
        except UnicodeDecodeError:
            print(f"Fallo al leer con codificación: {encoding}")
            continue
        except Exception as e:
            print(f"Error inesperado al leer con codificación {encoding}: {str(e)}")
            continue
    
    raise ValueError("No se pudo leer el archivo CSV con ninguna de las codificaciones intentadas")

def clean_and_validate_data(df):
    """
    Limpia y valida los datos del DataFrame.
    """
    # Eliminar filas completamente vacías
    df = df.dropna(how='all')
    
    # Verificar columnas requeridas en el orden correcto
    required_columns = [
        'Activity Number',
        'Activity',
        'Milestone/task',
        'Start Date (CET)',
        'End Date (CET)',
        'Responsible Person'
    ]
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Columnas faltantes en el archivo: {missing_columns}")
    
    return df

def parse_dates(date_str):
    """
    Intenta parsear fechas en diferentes formatos.
    """
    if pd.isna(date_str):
        return None
        
    date_formats = [
        '%d/%m/%y %H:%M'
    ]
    
    for date_format in date_formats:
        try:
            return datetime.strptime(str(date_str), date_format)
        except ValueError:
            continue
            
    print(f"No se pudo parsear la fecha: {date_str}")
    return None

def create_gantt_chart(df):
    """
    Crea un diagrama de Gantt con los datos procesados.
    """
    # Preparar los datos para el diagrama de Gantt
    gantt_data = []
    
    for _, row in df.iterrows():
        # Combinar la información de la tarea y la actividad para la descripción
        task_description = f"{row['Milestone/task']} - {row['Activity']}"
        
        gantt_data.append(dict(
            Task=row['Activity Number'],
            Description=task_description,
            Start=row['Start Date (CET)'],
            Finish=row['End Date (CET)'],
            Resource=row['Responsible Person']
        ))
    
    colors = {}
    for resource in df['Responsible Person'].unique():
        colors[resource] = f'rgb({randint(0,255)}, {randint(0,255)}, {randint(0,255)})'
    
    fig = ff.create_gantt(
        gantt_data,
        colors=colors,
        index_col='Resource',
        show_colorbar=True,
        group_tasks=True,
        showgrid_x=True,
        showgrid_y=True
    )
    
    # Personalizar el diseño
    fig.update_layout(
        title='Cronograma de Actividades',
        xaxis_title='Fecha',
        height=400 + (len(df) * 30),  # Altura dinámica basada en el número de tareas
        font=dict(size=10)
    )
    
    return fig

def read_and_process_data(file_path, start_date_filter, end_date_filter):
    """
    Lee y procesa el archivo de datos.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"El archivo {file_path} no existe.")

    file_extension = os.path.splitext(file_path)[1].lower()
    
    try:
        # Leer el archivo
        if file_extension == '.xlsx':
            df = pd.read_excel(file_path, engine='openpyxl')
        elif file_extension == '.xls':
            df = pd.read_excel(file_path, engine='xlrd')
        elif file_extension == '.csv':
            df = try_read_csv(file_path)
        else:
            raise ValueError(f"Formato de archivo no soportado: {file_extension}")
        
        print("\nNombres de columnas en el archivo:")
        print(df.columns.tolist())
        print("\nPrimeras filas de datos:")
        print(df.head())
        
        # Limpiar y validar datos
        df = clean_and_validate_data(df)
        
        # Convertir fechas con reporte detallado de errores
        for date_col in ['Start Date (CET)', 'End Date (CET)']:
            print(f"\nProcesando {date_col}...")
            df[date_col] = df[date_col].apply(parse_dates)
            invalid_dates = df[df[date_col].isna()][['Activity Number', date_col]]
            if not invalid_dates.empty:
                print(f"Advertencia: Se encontraron fechas inválidas en {date_col}:")
                print(invalid_dates)
        
        # Eliminar filas con fechas inválidas
        df = df.dropna(subset=['Start Date (CET)', 'End Date (CET)'])
        
        if df.empty:
            raise ValueError("No quedan datos válidos después de la limpieza")
        
        # Filtrar tareas que comienzan con 'Prod'
        df = df[df['Activity Number'].str.startswith('Prod', na=False)]
        
        if df.empty:
            raise ValueError("No se encontraron tareas que comiencen con 'Prod'")
        
        # Convertir fechas de filtro a datetime con validación
        try:
            start_date = datetime.strptime(start_date_filter, '%d/%m/%Y %H:%M')
        except ValueError as e:
            raise ValueError(f"Formato de fecha de inicio inválido: {start_date_filter}. Error: {str(e)}")
            
        try:
            end_date = datetime.strptime(end_date_filter, '%d/%m/%Y %H:%M')
        except ValueError as e:
            raise ValueError(f"Formato de fecha de fin inválido: {end_date_filter}. Error: {str(e)}")
        
        # Filtrar por rango de fechas
        df = df[
            (df['Start Date (CET)'] >= start_date) & 
            (df['End Date (CET)'] <= end_date)
        ]
        
        if df.empty:
            raise ValueError(f"No se encontraron tareas en el rango de fechas {start_date_filter} a {end_date_filter}")
        
        # Calcular duración en horas
        df['Duration (Hours)'] = (
            df['End Date (CET)'] - df['Start Date (CET)']
        ).dt.total_seconds() / 3600
        
        return df
        
    except Exception as e:
        raise Exception(f"Error al procesar el archivo: {str(e)}")

def main():
    # Ejemplo de uso
    file_path = 'C:/Users/Admin/Documents/MinBook.csv'
    start_date_filter = '08/11/2024 17:00'
    end_date_filter = '10/11/2024 23:59'
    
    try:
        print(f"Leyendo archivo: {file_path}")
        print(f"Filtro de rango de fechas: {start_date_filter} a {end_date_filter}")
        
        df = read_and_process_data(file_path, start_date_filter, end_date_filter)
        
        print("\nResumen de datos procesados:")
        print(f"Total de tareas: {len(df)}")
        print("\nRango de fechas en los datos:")
        print(f"Inicio más temprano: {df['Start Date (CET)'].min()}")
        print(f"Fin más tardío: {df['End Date (CET)'].max()}")
        print("\nPersonas responsables únicas:")
        print(df['Responsible Person'].unique())
        
        print("\nCreando diagrama de Gantt...")
        fig = create_gantt_chart(df)
        
        output_file = 'gantt_chart.html'
        pio.write_html(fig, output_file)
        print(f"\nDiagrama de Gantt generado exitosamente: {output_file}")
        
        print("\nResumen de tareas:")
        for _, row in df.iterrows():
            print(f"\nTarea: {row['Activity Number']}")
            print(f"Actividad: {row['Activity']}")
            print(f"Descripción: {row['Milestone/task']}")
            print(f"Responsable: {row['Responsible Person']}")
            print(f"Duración: {row['Duration (Hours)']:.2f} horas")
            print("-" * 50)
            
    except FileNotFoundError as e:
        print(f"Error: {str(e)}")
        print("Por favor, verifica que el archivo existe y la ruta es correcta.")
    except ValueError as e:
        print(f"Error: {str(e)}")
        print("Por favor, verifica tus datos de entrada y el formato de fecha.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {str(e)}")
        print("Por favor, verifica el formato y los datos de tu archivo de entrada.")

if __name__ == "__main__":
    main()