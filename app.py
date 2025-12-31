import io
import base64
import pyodbc
import barcode
from flask import Flask, render_template, request, redirect, url_for, session, flash
from barcode.writer import ImageWriter
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'chave_secreta_artenioreis_2025_modelos_v1'

# --- FUNÇÃO DE CONEXÃO COM O BANCO ---
def get_db_connection():
    driver = session.get('driver')
    server = session.get('server')
    database = session.get('database')
    trusted = session.get('trusted')

    if not all([driver, server, database]):
        raise Exception("Dados de conexão incompletos na sessão.")

    conn_str = f"Driver={{{driver}}};Server={server};Database={database};"

    if trusted == 'yes':
        conn_str += "Trusted_Connection=yes;"
    else:
        user = session.get('user')
        pwd = session.get('pwd')
        conn_str += f"UID={user};PWD={pwd};"
    
    return pyodbc.connect(conn_str)

# --- GERADOR DE CÓDIGO DE BARRAS (EAN-13 Preferencial) ---
def gerar_barcode_base64(valor_ean):
    if not valor_ean or str(valor_ean).strip() == "":
        return None
    
    ean_limpo = str(valor_ean).strip()
    buffer = io.BytesIO()
    writer = ImageWriter()

    try:
        # Tenta EAN-13 padrão
        if len(ean_limpo) in [12, 13] and ean_limpo.isdigit():
            barcode_instance = barcode.get('ean13', ean_limpo, writer=writer)
            barcode_instance.write(buffer, options={"module_height": 8.0, "font_size": 0, "text_distance": 1.0, "write_text": False})
        else:
             raise ValueError("Não é EAN-13")

    except Exception:
        # Fallback para Code128
        buffer.seek(0); buffer.truncate()
        try:
            barcode_instance = barcode.get('code128', ean_limpo, writer=writer)
            barcode_instance.write(buffer, options={"module_height": 8.0, "font_size": 0, "text_distance": 0, "write_text": False})
        except: return None

    return base64.b64encode(buffer.getvalue()).decode('utf-8')

# ================= ROTAS =================

@app.route('/', methods=['GET', 'POST'])
def login_banco():
    if request.method == 'POST':
        session['server'] = request.form.get('server')
        session['database'] = request.form.get('database')
        session['driver'] = request.form.get('driver')
        session['trusted'] = request.form.get('trusted')
        session['user'] = request.form.get('user')
        session['pwd'] = request.form.get('pwd')
        try:
            conn = get_db_connection(); conn.close()
            return redirect(url_for('busca'))
        except Exception as e: flash(f"Falha na conexão: {e}", "danger")
    return render_template('conexao.html')

@app.route('/busca')
def busca():
    if 'server' not in session: return redirect(url_for('login_banco'))
    return render_template('busca.html')

@app.route('/pesquisar', methods=['POST'])
def pesquisar():
    termo = request.form.get('termo', '').strip()
    if not termo: return redirect(url_for('busca'))

    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        # (Sua Query de busca complexa aqui - mantida a mesma da versão anterior)
        query = f"""
        SELECT pc.Cod_Produt, pr.Descricao AS Des_Produt, pr.Cod_EAN,
            Prc_Venda_V = (SELECT CASE WHEN (Round(pc.Prc_Promoc * 100, 2) > 0) THEN CASE WHEN (Round(pc.Prc_Promoc * 100, 2) > 0) THEN Round(pc.Prc_Promoc * (1 - pc.Per_Descon / 100), 2) ELSE pc.Prc_Promoc END ELSE CASE WHEN (Round(px.Prc_Venda * 100, 2) > 0) THEN Round(px.Prc_Venda * (1 - pc.Per_Descon / 100), 2) ELSE px.Prc_Venda END END)
        FROM PCXPR pc INNER JOIN PRODU pr ON (pc.Cod_Produt = pr.Codigo) INNER JOIN PRXES px ON (pc.Cod_Produt = px.Cod_Produt)
        WHERE pc.Id_PolCom = 432 AND px.Cod_Estabe = 0 AND (CAST(pc.Cod_Produt AS VARCHAR) = ? OR pr.Descricao LIKE ? OR pr.Cod_EAN = ?)
        """
        cursor.execute(query, (termo, f'%{termo}%', termo))
        resultados = cursor.fetchall()
        conn.close()
        return render_template('busca.html', resultados=resultados, termo=termo)
    except Exception as e:
        flash(f"Erro na busca: {e}", "danger")
        return redirect(url_for('busca'))

# --- NOVA ROTA DE ETIQUETA COM SUPORTE A MODELOS ---
@app.route('/etiqueta/<cod_prod>')
def etiqueta(cod_prod):
    # Captura o modelo da URL (ex: ?modelo=atacado). Padrão é 'gondola'.
    modelo_selecionado = request.args.get('modelo', 'gondola')

    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Usamos a sua Query Completa original para garantir que temos todos os dados (Atacado, Qtd Min, Unidade)
        query = """
       SELECT 1 AS Qtd_Produto, pc.Cod_Produt, pr.Descricao AS Des_Resumi, pr.Descricao AS Des_Produt, fa.Fantasia AS Des_Fabric, pc.Prc_Promoc, px.Prc_Venda, pc.Qtd_Minimo, pc.Per_Descon,
       Prc_Venda_V = (SELECT CASE WHEN (Round(pc.Prc_Promoc * 100, 2) > 0) THEN CASE WHEN (Round(pc.Prc_Promoc * 100, 2) > 0) THEN Round(pc.Prc_Promoc * (1 - pc.Per_Descon / 100), 2) WHEN (Round(pl.Per_AcrAutPrc * 100, 2) > 0) THEN Round(pc.Prc_Promoc * (1 + pl.Per_AcrAutPrc / 100), 2) WHEN (Round(pl.Per_DscAutPrc * 100, 2) > 0) THEN Round(pc.Prc_Promoc * (1 - pl.Per_DscAutPrc/ 100), 2) ELSE pc.Prc_Promoc END ELSE CASE WHEN (Round(px.Prc_Venda *100, 2) > 0) THEN Round(px.Prc_Venda * (1 - pc.Per_Descon / 100), 2) WHEN (Round(pl.Per_AcrAutPrc * 100, 2) > 0) THEN Round(px.Prc_Venda * (1 + pl.Per_AcrAutPrc / 100), 2) WHEN (Round(pl.Per_DscAutPrc * 100, 2) > 0) THEN Round(px.Prc_Venda * (1 - pl.Per_DscAutPrc/ 100), 2) ELSE px.Prc_Venda END END),
       pc.Qtd_Min2, pc.Per_Dsc2,
       Prc_Venda_A = (SELECT CASE WHEN (Round(pc.Prc_Promoc * 100, 2) > 0) THEN CASE WHEN (Round(pc.Prc_Promoc * 100, 2) > 0) THEN Round(pc.Prc_Promoc * (1 - pc.Per_Dsc2 / 100), 2) WHEN (Round(pl.Per_AcrAutPrc * 100, 2) > 0) THEN Round(pc.Prc_Promoc * (1 + pl.Per_AcrAutPrc / 100), 2) WHEN (Round(pl.Per_DscAutPrc * 100, 2) > 0) THEN Round(pc.Prc_Promoc * (1 - pl.Per_DscAutPrc/ 100), 2) ELSE pc.Prc_Promoc END ELSE CASE WHEN (Round(px.Prc_Venda *100, 2) > 0) THEN Round(px.Prc_Venda * (1 - pc.Per_Dsc2 / 100), 2) WHEN (Round(pl.Per_AcrAutPrc * 100, 2) > 0) THEN Round(px.Prc_Venda * (1 + pl.Per_AcrAutPrc / 100), 2) WHEN (Round(pl.Per_DscAutPrc * 100, 2) > 0) THEN Round(px.Prc_Venda * (1 - pl.Per_DscAutPrc/ 100), 2) ELSE px.Prc_Venda END END),
       pc.Per_DscVis, pc.Qtd_Maximo, pc.Per_DscVis2, pl.Per_AcrAutPrc, pl.Per_DscAutPrc, pr.Unidade_Venda, pr.Cod_EAN, pa.Fat_CnvApr
FROM PCXPR pc INNER JOIN PRODU pr ON (pc.Cod_Produt = pr.Codigo) INNER JOIN PRXAP pa ON (pc.Cod_Produt = pa.Cod_Produt AND pa.Flg_Padrao = 1) INNER JOIN FABRI fa ON (pr.Cod_Fabricante = fa.Codigo) INNER JOIN PRXES px ON (pc.Cod_Produt = px.Cod_Produt) INNER JOIN POCOM pl ON (pc.Id_PolCom = pl.Id_PolCom)
WHERE pc.Id_PolCom = 432 AND IsNull(pr.Cod_Classif, '') <> '' AND px.Cod_Estabe = 0
  AND pc.Cod_Produt = ?
        """
        cursor.execute(query, (cod_prod,))
        row = cursor.fetchone()
        conn.close()

        if row:
            # Mapeamos todos os dados necessários para os dois modelos
            produto = {
                "id": row.Cod_Produt,
                "descricao": row.Des_Produt,
                "ean": row.Cod_EAN,
                "unidade": row.Unidade_Venda,
                "preco_de": float(row.Prc_Venda) if row.Prc_Venda else 0.0, # Usando preço tabela como 'DE'
                "preco_por": float(row.Prc_Venda_V) if row.Prc_Venda_V else 0.0, # Preço Venda Varejo
                "preco_atacado": float(row.Prc_Venda_A) if row.Prc_Venda_A else 0.0, # Preço Venda Atacado
                "qtd_atacado": int(row.Qtd_Min2) if row.Qtd_Min2 else 1 # Quantidade mínima para atacado
            }

            # Gera o código de barras apenas se for o modelo de gôndola
            barcode_img = None
            if modelo_selecionado == 'gondola':
                 barcode_img = gerar_barcode_base64(produto["ean"])
            
            data_atual = datetime.now().strftime("%d/%m/%y %H:%M:%S")

            # Passamos a variável 'modelo' para o template decidir qual HTML usar
            return render_template('etiqueta.html', produto=produto, barcode_img=barcode_img, data_atual=data_atual, modelo=modelo_selecionado)
        
        return "Produto não encontrado.", 404
    except Exception as e:
        return f"Erro interno: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)