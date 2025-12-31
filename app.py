import io
import base64
import pyodbc
from flask import Flask, render_template, request
from barcode import Code128  # Code128 é o padrão mais seguro para diversos tamanhos de EAN
from barcode.writer import ImageWriter

app = Flask(__name__)

# --- CONFIGURAÇÃO DE CONEXÃO (Ajuste conforme seu ambiente) ---
def get_db_connection():
    conn_str = (
        "Driver={SQL Server};"
        "Server=SEU_SERVIDOR;"
        "Database=SEU_BANCO;"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str)

# --- FUNÇÃO DE GERAÇÃO DO CÓDIGO DE BARRAS ---
def gerar_barcode_base64(valor_ean):
    """
    Recebe o valor do campo Cod_EAN e transforma em imagem Base64.
    """
    if not valor_ean:
        return None
    
    try:
        # Transformamos em string e limpamos espaços
        ean_str = str(valor_ean).strip()
        
        # Gerando o código na memória (BytesIO) para não precisar salvar arquivos no disco
        buffer = io.BytesIO()
        # Usamos Code128 pois ele aceita o EAN13 e também códigos menores/internos
        Code128(ean_str, writer=ImageWriter()).write(buffer)
        
        # Converte o conteúdo binário para Base64 para o HTML ler
        return base64.b64encode(buffer.getvalue()).decode('utf-8')
    except Exception as e:
        print(f"Erro ao gerar código de barras para o EAN {valor_ean}: {e}")
        return None

# --- ROTA PRINCIPAL ---
@app.route('/etiqueta/<cod_prod>')
def gerar_etiqueta(cod_prod):
    conn = get_db_connection()
    cursor = conn.cursor()

    # AJUSTE NA QUERY: Pegando o pr.Cod_EAN explicitamente
    query = """
        SELECT 
            pr.Cod_Prod as ID_Interno, 
            pr.Descricao, 
            pr.Cod_EAN as Codigo_Barra 
        FROM Produtos pr 
        WHERE pr.Cod_Prod = ?
    """
    
    cursor.execute(query, (cod_prod,))
    row = cursor.fetchone()
    
    if not row:
        conn.close()
        return "Produto não encontrado", 404

    # Mapeando os dados
    produto = {
        "id": row.ID_Interno,
        "descricao": row.Descricao,
        "ean": row.Codigo_Barra
    }

    # AQUI ESTÁ A CORREÇÃO: Passamos o Codigo_Barra (Cod_EAN) para a função
    # e não o ID_Interno (cod_prod).
    barcode_base64 = gerar_barcode_base64(produto["ean"])

    conn.close()

    return render_template('etiqueta.html', produto=produto, barcode_img=barcode_base64)

if __name__ == '__main__':
    app.run(debug=True)