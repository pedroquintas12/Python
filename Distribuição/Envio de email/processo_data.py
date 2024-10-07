import uuid
from datetime import datetime
from db_conexão import get_db_connection
import mysql.connector
import logging
from logger_config import logger

def fetch_processes_and_clients():


    db_connection = get_db_connection()
    db_cursor = db_connection.cursor()

    try:
        query = (
            "SELECT c.Cliente_VSAP as clienteVSAP, p.Cod_escritorio, p.numero_processo, "
            "MAX(p.data_distribuicao) as data_distribuicao, "
            "p.orgao_julgador, p.tipo_processo, p.status, "
            "GROUP_CONCAT(DISTINCT a.nome ORDER BY a.nome SEPARATOR ', ') AS nomesAutores, "
            "GROUP_CONCAT(DISTINCT r.nome ORDER BY r.nome SEPARATOR ', ') AS nomesReus, "
            "GROUP_CONCAT(distinct l.link_documento order by l.link_documento separator ' | ') as Link, "
            "p.uf, p.sigla_sistema, MAX(p.instancia), p.tribunal, MAX(p.ID_processo), MAX(p.LocatorDB), p.tipo_processo, MAX(c.emails) "
            "FROM apidistribuicao.processo AS p "
            "LEFT JOIN apidistribuicao.clientes AS c ON p.Cod_escritorio = c.Cod_escritorio "
            "LEFT JOIN apidistribuicao.processo_autor AS a ON p.ID_processo = a.ID_processo "
            "LEFT JOIN apidistribuicao.processo_reu AS r ON p.ID_processo = r.ID_processo "
            "LEFT JOIN apidistribuicao.processo_docinicial as l ON p.ID_processo = l.ID_processo "
            "WHERE l.doc_peticao_inicial = 0 " 
            "AND p.status = 'P' "
            "GROUP BY p.numero_processo, clienteVSAP, p.Cod_escritorio, p.orgao_julgador, p.tipo_processo, p.uf, p.sigla_sistema, p.tribunal;"
        )

        db_cursor.execute(query)
        results = db_cursor.fetchall()
        clientes_data = {}

        for result in results:
            # Process the result and add to clientes_data
            process_result(result, clientes_data)

        return clientes_data

    except mysql.connector.Error as err:
        logger.error(f"Erro ao executar a consulta: {err}")
        return {}
    finally:
        db_cursor.close()
        db_connection.close()

def process_result(result, clientes_data):
    clienteVSAP = result[0]
    num_processo = result[2]
    data_distribuicao = datetime.strptime(str(result[3]), '%Y-%m-%d').strftime('%d/%m/%Y')
    tribunal = result[13]
    uf = result[10]
    instancia = result[12]
    comarca = result[11]
    polo_ativo = result[7] if result[7] else "[Nenhum dado disponível]"
    polo_passivo = result[8] if result[8] else "[Nenhum dado disponível]"
    emails= result[17]

    # Collect links for the process
    links_list = fetch_links(result[14])

    # Add process data to the clients' data
    if clienteVSAP not in clientes_data:
        clientes_data[clienteVSAP] = []

    clientes_data[clienteVSAP].append({
        'ID_processo' : result[14],
        'cod_escritorio': result[1],
        'numero_processo': num_processo,
        'data_distribuicao': data_distribuicao,
        'orgao': result[4],
        'classe_judicial': result[5],
        'polo_ativo': polo_ativo,
        'polo_passivo': polo_passivo,
        'links': links_list,
        'tribunal': tribunal,
        'uf': uf,
        'instancia': instancia,
        'comarca': comarca,
        'localizador': result[15],
        'tipo_processo': result[16],
        'emails': emails
    })

def fetch_links(process_id):
    db_connection = get_db_connection()
    db_cursor = db_connection.cursor()
    
    querylinks = "SELECT * FROM apidistribuicao.processo_docinicial WHERE ID_PROCESSO = %s AND doc_peticao_inicial= 0"
    db_cursor.execute(querylinks, (process_id,))
    results_links = db_cursor.fetchall()

    links_list = []
    for links in results_links:
        id_link = links[1]
        link_doc = links[2]
        tipoLink = links[3]
        links_list.append({'link_doc': link_doc, 'tipoLink': tipoLink, 'id_link': id_link})

    db_cursor.close()
    db_connection.close()

    return links_list
