from model.metods import registers, categorizarPorDias, save_to_excel_global, registers_wfm

import pandas as pd


if __name__ == '__main__':

    params_inc = {
        'sysparm_query':
        '^assignment_group=56068a6d1bf16410f720524f034bcbdf'
        '^state=3^ORstate=2^ORstate=1'
        '^company=b482d6a91b66a8101df3bb7f034bcbf0',
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
        '^assignment_group=56068a6d1bf16410f720524f034bcbdf'
        '^state=5^ORstate=1^ORstate=5'
        '^company=b482d6a91b66a8101df3bb7f034bcbf0',
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
    params_wfm = {
        'sysparm_query': 
        '^company=b482d6a91b66a8101df3bb7f034bcbf0'
        '^assignment_group=56068a6d1bf16410f720524f034bcbdf'
        '^state=2^ORstate=-5',
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
    all_ent_agendador = registers_wfm('u_ent_agendador', params_wfm)

    # Convertir registros obtenidos en dataframe
    incident_df = pd.DataFrame(incident_records)
    req_item_df = pd.DataFrame(req_item_records)
    ent_agendador_df = pd.DataFrame(all_ent_agendador)

    #Consolidar registros en un dataframe
    reporte_sentral = pd.concat([incident_df, req_item_df], ignore_index=True)

    #Combinar datos de consolidado con agendamientos
    reporte_consolidado = pd.merge(
        reporte_sentral,        # DataFrame principal (la hoja consolidada)
        ent_agendador_df,       # DataFrame de referencia (la hoja WFM)
        left_on='sys_id',       # Columna en la hoja consolidada
        right_on='parent',      # Columna en la hoja WFM
        how='left'              # Tipo de merge: 'left' mantiene todos los valores de la izquierda
    )

    #categorizar
    reporte_consolidado = categorizarPorDias(reporte_consolidado)

    #Identificación de hojas excel
    dict_sheetname = {
        'Consolidado': reporte_sentral,
        'Merge_TKT': reporte_consolidado
        }
    
    # Genera consolidado de INC y RITM
    save_to_excel_global(dict_sheetname)