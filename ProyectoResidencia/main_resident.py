from model.metods import registers, categorizarPorDias, save_to_excel, registers_wfm

import pandas as pd

if __name__ == '__main__':
    base_query = (
        '^location=11dae7de476c3110d9d3a03a536d4370^ORlocation=f8dc277347ac3910d9d3a03a536d43e6^ORlocation=f5036756476c3110d9d3a03a536d4327'
        '^ORlocation=f3377b3e47f07d50d9d3a03a536d43d7^ORlocation=f1036756476c3110d9d3a03a536d4326^ORlocation=e04d9a0f1be53510d5166392b24bcb78'
        '^ORlocation=7d036756476c3110d9d3a03a536d4326^ORlocation=79036756476c3110d9d3a03a536d4325^ORlocation=202daf7347ac3910d9d3a03a536d4329'
        '^assignment_group=56068a6d1bf16410f720524f034bcbdf'
        '^company=b482d6a91b66a8101df3bb7f034bcbf0'
    )
    params_inc = {
        'sysparm_query': f"{base_query}^state=3^ORstate=2^ORstate=1",
        'sysparm_fields': [
            'sys_id,'
            'sys_updated_on,'
            'number,'
            'state,'
            'sys_created_on,'
            'short_description,'
            'description,'
            'location,'
        ],
    }
    params_ritm = {
        'sysparm_query': 
        f"{base_query}^state=5^ORstate=1^ORstate=2",
        'sysparm_fields': [
            'sys_id,'
            'number,'
            'state,'
            'sys_created_on,'
            'location,'
            'sys_updated_on,'
            'short_description,'
            'description,'
        ],
    }
    params_wfm_pendientes = {
        
        'sysparm_query': 
        f"{base_query}^state=1^ORstate=2^ORstate=-5",
        'sysparm_fields': [
            'parent,'
            'number,'
            'state,'
            'u_sla_start,'
            ],
    }
    params_wfm_todos = {
        
        'sysparm_query': 
        f"{base_query}^state=1^ORstate=2^ORstate=-5^ORstate=4",
        'sysparm_fields': [
            'parent,'
            'number,'
            'state,'
            'u_sla_start,'
            ],
    }
    
    # Obtener registros de tablas
    incident_records = registers('incident', params_inc)
    req_item_records = registers('sc_req_item', params_ritm)
    pend_ent_agendador = registers_wfm('u_ent_agendador', params_wfm_pendientes)
    all_ent_agendador = registers_wfm('u_ent_agendador', params_wfm_todos)

    # Convertir registros obtenidos en dataframe
    incident_df = pd.DataFrame(incident_records)
    req_item_df = pd.DataFrame(req_item_records)
    pend_ent_agendador_df = pd.DataFrame(pend_ent_agendador)
    todos_ent_agendador_df = pd.DataFrame(all_ent_agendador)

    #Consolidar registros en un dataframe
    reporte_sentral = pd.concat([incident_df, req_item_df], ignore_index=True)

    #Combinar datos de consolidado con agendamientos
    reporte_consolidado = pd.merge(
        reporte_sentral,        # DataFrame principal (la hoja consolidada)
        pend_ent_agendador_df,  # DataFrame de referencia (la hoja WFM)
        left_on='sys_id',       # Columna en la hoja consolidada
        right_on='parent',      # Columna en la hoja WFM
        how='left'              # Tipo de merge: 'left' mantiene todos los valores de la izquierda
    )
    reporte_wfm = pd.merge(
        todos_ent_agendador_df,       # DataFrame de referencia (la hoja WFM)
        reporte_sentral,        # DataFrame principal (la hoja consolidada)
        
        left_on='parent',       # Columna en la hoja consolidada
        right_on='sys_id',      # Columna en la hoja WFM
        how='left'              # Tipo de merge: 'left' mantiene todos los valores de la izquierda
    )

    #categorizar
    reporte_consolidado = categorizarPorDias(reporte_consolidado)

    #Identificación de hojas excel
    dict_sheetname = {
        'Consolidado': reporte_sentral,
        'En Gestión': reporte_consolidado,
        'Cantidad de Gestiones': reporte_wfm
        }
    
    # Genera consolidado de INC y RITM
    save_to_excel(dict_sheetname)
