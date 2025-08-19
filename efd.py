from flask import Blueprint, request, jsonify, send_file
from src.efd_generator import EFDGenerator
import os
import tempfile
from datetime import datetime

efd_bp = Blueprint('efd', __name__)

@efd_bp.route('/generate', methods=['POST'])
def generate_efd():
    try:
        data = request.get_json()
        
        # Validar dados de entrada
        required_fields = ['data_inicio', 'data_fim', 'id_inventario']
        for field in required_fields:
            if field not in data:
                return jsonify({'error': f'Campo obrigatório: {field}'}), 400
        
        data_inicio = data['data_inicio']
        data_fim = data['data_fim']
        id_inventario = data['id_inventario']
        
        # Validar formato das datas
        try:
            datetime.strptime(data_inicio, '%Y-%m-%d')
            datetime.strptime(data_fim, '%Y-%m-%d')
        except ValueError:
            return jsonify({'error': 'Formato de data inválido. Use YYYY-MM-DD'}), 400
        
        # Criar diretório temporário para o arquivo
        temp_dir = tempfile.mkdtemp()
        
        # Gerar arquivo EFD
        generator = EFDGenerator()
        generator.generate_efd_icms_ipi(
            data_inicio, 
            data_fim, 
            temp_dir, 
            data_inicio,  # data_contabil
            id_inventario
        )
        
        # Encontrar o arquivo gerado
        dt_inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
        filename = f"EFD_ICMS_IPI_{dt_inicio.month}_{dt_inicio.year}.txt"
        file_path = os.path.join(temp_dir, filename)
        
        if os.path.exists(file_path):
            return jsonify({
                'success': True,
                'message': 'Arquivo EFD gerado com sucesso',
                'filename': filename,
                'download_url': f'/api/efd/download/{filename}'
            })
        else:
            return jsonify({'error': 'Erro ao gerar arquivo EFD'}), 500
            
    except Exception as e:
        return jsonify({'error': f'Erro interno: {str(e)}'}), 500

@efd_bp.route('/download/<filename>', methods=['GET'])
def download_efd(filename):
    try:
        # Em um ambiente real, você salvaria os arquivos em um local específico
        # Por enquanto, vamos retornar um arquivo de exemplo
        temp_dir = tempfile.mkdtemp()
        file_path = os.path.join(temp_dir, filename)
        
        # Criar um arquivo de exemplo se não existir
        if not os.path.exists(file_path):
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("|0000|017|0|01012023|31012023|EMPRESA EXEMPLO LTDA|12345678000199|SP|123456789|3550308|||||A|0|\n")
                f.write("|0001|0|\n")
                f.write("|0990|2|\n")
                f.write("|B001|1|\n")
                f.write("|B990|1|\n")
                f.write("|C001|1|\n")
                f.write("|C990|1|\n")
                f.write("|D001|1|\n")
                f.write("|D990|1|\n")
                f.write("|E001|0|\n")
                f.write("|E990|1|\n")
        
        return send_file(file_path, as_attachment=True, download_name=filename)
        
    except Exception as e:
        return jsonify({'error': f'Erro ao baixar arquivo: {str(e)}'}), 500

@efd_bp.route('/test-connection', methods=['GET'])
def test_connection():
    try:
        generator = EFDGenerator()
        generator.db_conn.connect_to_database()
        
        # Testar uma consulta simples
        result = generator.db_conn.fetch_all("SELECT 1 as test")
        generator.db_conn.disconnect_from_database()
        
        if result:
            return jsonify({'success': True, 'message': 'Conexão com banco de dados OK'})
        else:
            return jsonify({'success': False, 'message': 'Erro na consulta de teste'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'Erro de conexão: {str(e)}'})

@efd_bp.route('/status', methods=['GET'])
def status():
    return jsonify({
        'status': 'online',
        'service': 'EFD Generator',
        'version': '1.0.0'
    })

