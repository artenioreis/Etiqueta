import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import barcode
# IMPORTAÇÃO EXPLÍCITA PARA O PYINSTALLER NÃO FALHAR
from barcode.writer import ImageWriter 
import os
import tempfile
import json
import sys

try:
    import win32print
    import win32api
    WINDOWS_PRINT_AVAILABLE = True
except Exception:
    WINDOWS_PRINT_AVAILABLE = False

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".etiqueta_58mm_exe_fix.json")

class EtiquetaTermicaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador 58mm - VarejãoFarma")
        self.root.geometry("1000x750")
        
        self.largura_papel = 58
        self.altura_etiqueta = 40
        self.logo_path = os.path.join(BASE_DIR, "logo.png")

        self.conn = None
        self.cursor = None
        self.config = self._carregar_config()
        self.setup_ui()
        self._preencher_config_inicial()

    def _carregar_config(self):
        try:
            if os.path.exists(CONFIG_PATH):
                with open(CONFIG_PATH, "r", encoding="utf-8") as f: return json.load(f)
        except Exception: pass
        return {}

    def _salvar_config(self):
        conf = {"server": self.servidor.get(), "database": self.banco.get(), "username": self.usuario.get()}
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f: json.dump(conf, f, indent=2)
        except Exception: pass

    def _preencher_config_inicial(self):
        if self.config.get("server"): self.servidor.insert(0, self.config.get("server"))
        if self.config.get("database"): self.banco.insert(0, self.config.get("database"))
        if self.config.get("username"): self.usuario.insert(0, self.config.get("username"))

    def setup_ui(self):
        f_conn = ttk.LabelFrame(self.root, text=" SQL Server ", padding=10)
        f_conn.pack(fill="x", padx=10, pady=5)
        ttk.Label(f_conn, text="Servidor:").grid(row=0, column=0); self.servidor = ttk.Entry(f_conn); self.servidor.grid(row=0, column=1)
        ttk.Label(f_conn, text="Banco:").grid(row=0, column=2); self.banco = ttk.Entry(f_conn); self.banco.grid(row=0, column=3)
        ttk.Label(f_conn, text="Usuário:").grid(row=1, column=0); self.usuario = ttk.Entry(f_conn); self.usuario.grid(row=1, column=1)
        ttk.Label(f_conn, text="Senha:").grid(row=1, column=2); self.senha = ttk.Entry(f_conn, show="*"); self.senha.grid(row=1, column=3)
        ttk.Button(f_conn, text="Conectar", command=self.conectar_banco).grid(row=0, column=4, rowspan=2, padx=10)

        f_search = ttk.LabelFrame(self.root, text=" Busca de Produtos ", padding=10)
        f_search.pack(fill="x", padx=10, pady=5)
        self.filtro_busca = ttk.Entry(f_search, width=50)
        self.filtro_busca.pack(side="left", padx=5)
        ttk.Button(f_search, text="Buscar", command=self.buscar_produtos).pack(side="left")

        f_table = ttk.Frame(self.root, padding=10)
        f_table.pack(fill="both", expand=True)
        cols = ("Código", "Descrição", "Fabricante", "Preço", "EAN")
        self.tree = ttk.Treeview(f_table, columns=cols, show="headings")
        for c in cols: self.tree.heading(c, text=c); self.tree.column(c, width=90)
        self.tree.column("Descrição", width=400)
        scroll = ttk.Scrollbar(f_table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")

        f_act = ttk.Frame(self.root, padding=10)
        f_act.pack(fill="x")
        ttk.Label(f_act, text="Qtd:").pack(side="left")
        self.quantidade = ttk.Spinbox(f_act, from_=1, to=100, width=5); self.quantidade.set(1); self.quantidade.pack(side="left", padx=5)
        ttk.Button(f_act, text="Gerar PDF", command=self.gerar_pdf).pack(side="left", padx=10)
        ttk.Button(f_act, text="Imprimir Direto", command=self.imprimir_direto).pack(side="left")

    def conectar_banco(self):
        try:
            conn_str = f'DRIVER={{SQL Server}};SERVER={self.servidor.get()};DATABASE={self.banco.get()};UID={self.usuario.get()};PWD={self.senha.get()}'
            self.conn = pyodbc.connect(conn_str, timeout=5); self.cursor = self.conn.cursor()
            self._salvar_config(); messagebox.showinfo("Sucesso", "Conectado!")
        except Exception as e: messagebox.showerror("Erro", str(e))

    def buscar_produtos(self):
        if not self.conn: return
        try:
            for i in self.tree.get_children(): self.tree.delete(i)
            termo = f"%{self.filtro_busca.get().strip()}%"
            query = """
            SELECT pc.Cod_Produt, pr.Descricao, fb.Fantasia, pc.Prc_Promoc, pc.Per_Descon, pr.Cod_EAN
            FROM PCXPR pc
            INNER JOIN PRODU pr on (pc.Cod_Produt = pr.Codigo)
            LEFT OUTER JOIN FABRI fb on (pr.Cod_Fabricante = fb.Codigo)
            WHERE pc.Id_PolCom IN (432) AND (pc.Cod_Produt LIKE ? OR pr.Descricao LIKE ?)
            """
            self.cursor.execute(query, (termo, termo))
            for row in self.cursor.fetchall():
                p_promo = float(row[3]) if row[3] and float(row[3]) > 0 else 0
                p_desc = float(row[4]) if row[4] and float(row[4]) > 0 else 0
                preco = p_promo if p_promo > 0 else p_desc
                self.tree.insert("", "end", values=(row[0], row[1], row[2], f"{preco:.2f}", row[5]))
        except Exception as e: messagebox.showerror("Erro", str(e))

    def gerar_barcode_fix(self, codigo):
        """Geração robusta de barcode para o EXE."""
        try:
            cod = str(codigo).strip()
            if not cod or cod == "None": return None
            
            # Força Code128 para códigos curtos (ex: 3922)
            if cod.isdigit() and len(cod) in [12, 13]:
                bc_class = barcode.get_barcode_class('ean13')
            else:
                bc_class = barcode.get_barcode_class('code128')
            
            # Usa ImageWriter explicitamente instanciado
            bc = bc_class(cod, writer=ImageWriter())
            temp_path = os.path.join(tempfile.gettempdir(), f"barcode_{cod}")
            return bc.save(temp_path)
        except Exception as e:
            print(f"Erro ao gerar barcode: {e}")
            return None

    def criar_pdf_etiqueta(self, row_data, qtd, out_file):
        try:
            c = canvas.Canvas(out_file, pagesize=(self.largura_papel*mm, self.altura_etiqueta*mm))
            cod_int, nome, preco, ean = str(row_data[0]), str(row_data[1]).upper(), float(row_data[3]), str(row_data[4])
            centro_x = (self.largura_papel / 2) * mm

            for _ in range(qtd):
                c.setFont("Helvetica-Oblique", 8)
                c.drawString(2*mm, 36*mm, nome[:28]) 
                if len(nome) > 28: c.drawString(2*mm, 32*mm, nome[28:56])

                if os.path.exists(self.logo_path):
                    c.drawImage(self.logo_path, 0.5*mm, 17*mm, width=20*mm, height=14*mm, mask='auto', preserveAspectRatio=True)

                c.setFont("Helvetica", 7)
                c.drawString(1*mm, 13*mm, f"Cód: {cod_int}")

                c.saveState()
                c.setFont("Helvetica", 4); c.translate(57*mm, 20*mm); c.rotate(90)
                c.drawString(0, 0, datetime.datetime.now().strftime("%d/%m/%y %H:%M:%S"))
                c.restoreState()

                c.setFont("Helvetica-Bold", 14)
                c.drawCentredString(centro_x, 18*mm, f"Por R${preco:.2f}")

                bc_img = self.gerar_barcode_fix(ean if ean != "None" else cod_int)
                if bc_img:
                    largura_bc = 30
                    c.drawImage(bc_img, centro_x - (largura_bc/2)*mm, 4.5*mm, width=largura_bc*mm, height=8*mm, mask='auto')
                    c.setFont("Helvetica", 6)
                    c.drawCentredString(centro_x, 2*mm, ean if ean != "None" else cod_int)

                c.showPage()
            c.save()
            return True
        except Exception as e:
            messagebox.showerror("Erro PDF", str(e))
            return False

    def gerar_pdf(self):
        sel = self.tree.selection()
        if not sel: return
        row = self.tree.item(sel[0])['values']
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"etiqueta_{row[0]}.pdf")
        if path and self.criar_pdf_etiqueta(row, int(self.quantidade.get()), path): os.startfile(path)

    def imprimir_direto(self):
        if not WINDOWS_PRINT_AVAILABLE: return
        sel = self.tree.selection()
        if not sel: return
        temp_pdf = os.path.join(tempfile.gettempdir(), "etiqueta_direta.pdf")
        if self.criar_pdf_etiqueta(self.tree.item(sel[0])['values'], int(self.quantidade.get()), temp_pdf):
            win32api.ShellExecute(0, "print", temp_pdf, None, ".", 0)

if __name__ == "__main__":
    root = tk.Tk(); app = EtiquetaTermicaApp(root); root.mainloop()