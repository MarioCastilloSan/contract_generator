import streamlit as st
from data import  readData, singular_data_to_contract, multiple_data_to_contract,singular_path_clean, multiple_path_clean, create_zip_file
df =readData()


def main():
    st.title("Contract Generator")
    st.dataframe(df)
    row_selection = st.selectbox("Si quieres crear un solo contrato selecciona una fila:", df.index)
    if row_selection is not None:
        selected_data = df.iloc[row_selection]
        st.write("Detalles de la fila seleccionada:")
        st.write(selected_data)
       
    
        if st.button('Generar Contrato'):
            end_path = singular_data_to_contract(df, row_selection)
            with open(end_path, "rb") as file:
                btn = st.download_button(
                    label="Descargar Contrato",
                    data=file,
                    file_name=end_path.split("/")[-1],
                    mime="application/pdf"
                )
            singular_path_clean(end_path)

    st.write("Si quieres crear multiples contratos selecciona una fila")
    start = st.number_input('Inicio del rango', min_value=0, max_value=len(df)-1, value=0, key='range_start')
    end = st.number_input('Fin del rango', min_value=0, max_value=len(df), value=len(df), key='range_end')
    if st.button('Generar Contratos para el Rango'):
        if start < end:
            paths = multiple_data_to_contract(df, start, end)
            zip_path = create_zip_file(paths)  # Asume que esta función crea un archivo ZIP de los contratos
            with open(zip_path, "rb") as fp:
                st.download_button(
                    label="Descargar Todos los Contratos",
                    data=fp,
                    file_name="contratos.zip"
                )
            singular_path_clean(zip_path)  # Limpiar los archivos generados
            multiple_path_clean(paths)
        else:
            st.error("¡El inicio del rango debe ser menor que el fin del rango! .")
if __name__ == "__main__":
    main()