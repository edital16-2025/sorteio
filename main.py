import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import hashlib
import time

class InterfaceSorteio:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Sorteador de Desempate")
        self.opcao = None
        self.semente = None
        
        largura, altura = 350, 300
        x = (self.root.winfo_screenwidth() // 2) - (largura // 2)
        y = (self.root.winfo_screenheight() // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{x}+{y}")

        tk.Label(self.root, text="Selecione o método de desempate:", font=("Arial", 11, "bold"), pady=20).pack()

        tk.Button(self.root, text="1 - ALEATÓRIO (Sistema)", width=30, height=2, 
                  command=self.escolher_aleatorio).pack(pady=5)
        
        tk.Button(self.root, text="2 - SEMENTE NUMÉRICA", width=30, height=2, 
                  command=self.escolher_numerica).pack(pady=5)
        
        tk.Button(self.root, text="3 - SEMENTE DE HASH", width=30, height=2, 
                  command=self.escolher_hash).pack(pady=5)

    def escolher_aleatorio(self):
        self.opcao = 1
        self.semente = int(time.time())
        self.root.quit()

    def escolher_numerica(self):
        res = simpledialog.askinteger("Semente Numérica", "Digite o número da semente:", parent=self.root)
        if res is not None:
            self.opcao = 2
            self.semente = res
            self.root.quit()

    def escolher_hash(self):
        res = simpledialog.askstring("Semente Hash", "Cole o código HASH:", parent=self.root)
        if res:
            self.opcao = 3
            self.semente = int(hashlib.sha256(res.strip().encode()).hexdigest(), 16) % (2**32)
            self.root.quit()

def realizar_desempate():
    root_init = tk.Tk()
    root_init.withdraw()
    caminho = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Excel", "*.xlsx *.xls")])
    root_init.destroy()
    
    if not caminho: return

    app = InterfaceSorteio()
    app.root.mainloop()
    
    opcao = app.opcao
    semente_final = app.semente
    app.root.destroy()

    if opcao is None: return

    try:
        df = pd.read_excel(caminho)
        
        # Sorteio garantindo que a semente afete cada grupo
        np.random.seed(semente_final)
        res_df = df.copy()
        
        # Realiza o sorteio grupo a grupo
        for _, grupo_data in res_df.groupby('GRUPO'):
            indices = grupo_data.index
            res_df.loc[indices, 'ORDEM DESEMPATE'] = np.random.permutation(np.arange(1, len(indices) + 1))

        res_df['ORDEM DESEMPATE'] = res_df['ORDEM DESEMPATE'].astype(int)
        
        # ORDENAÇÃO CRÍTICA: Primeiro por Grupo, depois pela Ordem Sorteada
        res_df = res_df.sort_values(by=['GRUPO', 'ORDEM DESEMPATE'], ascending=[True, True])

        # Definir nome do arquivo de saída
        diretorio = os.path.dirname(caminho)
        nome_base = os.path.join(diretorio, "RESULTADO_DESEMPATE.xlsx")
        contador = 1
        while os.path.exists(nome_base):
            nome_base = os.path.join(diretorio, f"RESULTADO_DESEMPATE_{contador}.xlsx")
            contador += 1
        
        # Escrita manual para garantir os espaços
        writer = pd.ExcelWriter(nome_base, engine='xlsxwriter')
        workbook = writer.book
        worksheet = workbook.add_worksheet('Resultado')
        
        # Escrever Cabeçalhos
        for col_num, value in enumerate(res_df.columns.values):
            worksheet.write(0, col_num, value)

        # Escrever os dados respeitando o agrupamento e inserindo a linha em branco
        row_idx = 1
        grupos_unicos = res_df['GRUPO'].unique()
        
        for g in grupos_unicos:
            dados_grupo = res_df[res_df['GRUPO'] == g]
            for _, row in dados_grupo.iterrows():
                for col_num, val in enumerate(row):
                    worksheet.write(row_idx, col_num, val)
                row_idx += 1
            row_idx += 1 # PULA UMA LINHA após terminar cada grupo

        # Ajuste automático de largura de colunas
        for i, col in enumerate(res_df.columns):
            # Encontra o maior valor na coluna (dados ou cabeçalho)
            max_len = max(res_df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

        writer.close()
        messagebox.showinfo("Sucesso", f"Sorteio concluído!\nArquivo: {os.path.basename(nome_base)}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro no processamento: {str(e)}")

if __name__ == "__main__":
    realizar_desempate()
