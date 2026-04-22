import pandas as pd
from handler import SeleniumHandler
import os
from datetime import datetime
from tqdm import tqdm # Biblioteca para a barra de progresso

def executar_automacao():
    # 1. Carregamento dos processos
    try:
        if not os.path.exists("processos.txt"):
            print("❌ Arquivo processos.txt não encontrado!")
            return

        with open("processos.txt", "r") as f:
            lista_processos = [linha.strip() for linha in f if linha.strip()]
    except Exception as e:
        print(f"❌ Erro ao ler arquivo: {e}")
        return

    if not lista_processos:
        print("⚠️ A lista no TXT está vazia!")
        return

    resultados = []
    total = len(lista_processos)
    
    # Gerar carimbo de data e hora para o nome do arquivo
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    file_name = f"Relatorio_SEI_{timestamp}.xlsx"

    print(f"\n🚀 Iniciando Robô SEI ARTESP")
    print(f"📋 Total de processos: {total}")
    print(f"📂 O resultado será salvo em: {file_name}\n")

    # 2. Execução com Barra de Progresso
    with SeleniumHandler() as bot:
        # tqdm cria a barra visual no terminal
        for num_proc in tqdm(lista_processos, desc="Processando", unit="proc"):
            try:
                if bot.buscar_processo(num_proc):
                    info = bot.clicar_e_extrair()
                    resultados.append({"Processo": num_proc, "Especificação": info})
                else:
                    resultados.append({"Processo": num_proc, "Especificação": "❌ Erro na busca"})
            except Exception:
                resultados.append({"Processo": num_proc, "Especificação": "⚠️ Erro técnico"})

    # 3. Salvamento Formatado
    if resultados:
        print(f"\n\n📊 Finalizando e formatando planilha...")
        df = pd.DataFrame(resultados)
        
        with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Relatorio')
            workbook = writer.book
            worksheet = writer.sheets['Relatorio']
            
            # Formato profissional
            format_celula = workbook.add_format({
                'align': 'left', 
                'valign': 'vcenter', 
                'text_wrap': False,
                'border': 1
            })
            
            # Cabeçalho em negrito
            format_header = workbook.add_format({
                'bold': True,
                'bg_color': '#D7E4BC',
                'border': 1
            })

            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 4
                worksheet.set_column(i, i, max_len, format_celula)
                # Aplica o formato de cabeçalho manualmente na primeira linha
                worksheet.write(0, i, col, format_header)
                
        print(f"💎 Trabalho concluído! Verifique o arquivo: {file_name}")

if __name__ == "__main__":
    executar_automacao()