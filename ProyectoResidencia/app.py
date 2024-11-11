from flask import Flask, render_template, jsonify, request, send_file
from model.metods import registers, registers_wfm
import pandas as pd
import os

app = Flask(__name__)

# Ruta de renderizado
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/request', methods=['POST'])
def handle_request():
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
            'sys_id',
            'sys_updated_on',
            'number',
            'state',
            'sys_created_on',
            'short_description',
            'description',
            'location',
        ],
    }
    params_ritm = {
        'sysparm_query': f"{base_query}^state=5^ORstate=1^ORstate=2",
        'sysparm_fields': [
            'sys_id',
            'number',
            'state',
            'sys_created_on',
            'location',
            'sys_updated_on',
            'short_description',
            'description',
        ],
    }
    params_wfm = {
        'sysparm_query': f"{base_query}^state=2^ORstate=-5",
        'sysparm_fields': [
            'parent',
            'number',
            'state',
            'u_sla_start',
            
        ],
    }

    try:
        incident_records = registers('incident', params_inc)
        req_item_records = registers('sc_req_item', params_ritm)
        all_ent_agendador = registers_wfm('u_ent_agendador', params_wfm)
    except Exception as e:
        print(f"Error en respuesta de API: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)})

    incident_df = pd.DataFrame(incident_records)
    req_item_df = pd.DataFrame(req_item_records)
    ent_agendador_df = pd.DataFrame(all_ent_agendador)

    reporte_sentral = pd.concat([incident_df, req_item_df], ignore_index=True)

    reporte_consolidado = pd.merge(
        reporte_sentral,
        ent_agendador_df,
        left_on='sys_id',
        right_on='parent',
        how='left'
    )

    dict_sheetname = {
        'Consolidado': reporte_sentral,
        'Merge_TKT': reporte_consolidado
    }

    incident_df = pd.DataFrame(incident_records)
    req_item_df = pd.DataFrame(req_item_records)
    ent_agendador_df = pd.DataFrame(all_ent_agendador)


    file_path = 'reporteResidencia.xlsx'
    with pd.ExcelWriter(file_path) as writer:
        for sheet_name, data in dict_sheetname.items():
            data.to_excel(writer, sheet_name=sheet_name)

    return send_file(file_path, as_attachment=True)

@app.after_request
def after_request(response):
    if response.status_code == 200 and 'reporteResidencia.xlsx' in response.headers.get('Content-Disposition', ''):
        try:
            os.remove('reporteResidencia.xlsx')
        except OSError as e:
            print(f"Error al eliminar archivo: {e}")
    return response

if __name__ == '__main__':
    app.run(debug=True)
