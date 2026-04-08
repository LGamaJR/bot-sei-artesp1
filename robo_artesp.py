import pandas as pd
import os
import time
from handler import SeleniumHandler

def executar():
    arquivo_lista = "processos.txt"
    arquivo_saida = "Relatorio_Final_ARTESP.xlsx"

    # 1. Verifica se o arquivo de texto existe
    if not os.path.exists(arquivo_lista):
        print(f"❌ Erro: O arquivo '{arquivo_lista}' não foi encontrado.")
        return

    # 2. Lê a lista de processos
    with open(arquivo_lista, "r") as f:
        lista_processos = [linha.strip() for linha in f if linha.strip()]

    if not lista_processos:
        print("⚠️ A lista de processos está vazia.")
        return

    resultados = []

    # 3. Inicia a automação
    with SeleniumHandler() as bot:
        print(f"🚀 Iniciando processamento de {len(lista_processos)} processos...")
        
        for i, num_proc in enumerate(lista_processos):
            print(f"🔄 [{i+1}/{len(lista_processos)}] Processando: {num_proc}")
            
            if bot.buscar_processo(num_proc):
                info = bot.clicar_e_extrair()
            else:
                info = "❌ Erro na busca (Campo não localizado)"
            
            print(f"   ✅ Resultado: {info}")
            
            resultados.append({
                "Processo": num_proc, 
                "Especificação": info
            })
            
            # Pequena pausa para o servidor da ARTESP "respirar"
            time.sleep(2)

    # 4. Salvamento Final com Ajuste Automático de Colunas
    print("\n" + "="*30)
    print("💾 Formatando e salvando arquivo final...")
    
    try:
        # Usando XlsxWriter como engine para permitir formatação de colunas
        with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
            df_final = pd.DataFrame(resultados)
            df_final.to_excel(writer, index=False, sheet_name='Relatorio')
            
            workbook  = writer.book
            worksheet = writer.sheets['Relatorio']
            
            # Formato: Alinhado à esquerda e SEM quebra de linha (Single Line)
            formato_celula = workbook.add_format({
                'text_wrap': False, 
                'align': 'left',
                'valign': 'vcenter'
            })
            
            # Ajusta largura da Coluna A (Processo)
            worksheet.set_column('A:A', 25, formato_celula)
            
            # Ajusta largura da Coluna B (Especificação) com base no maior texto
            # Calculamos o comprimento máximo ou definimos um mínimo de 50
            if not df_final.empty:
                largura_maxima = max(df_final['Especificação'].astype(str).map(len).max(), 50)
                worksheet.set_column('B:B', largura_maxima + 2, formato_celula)
            
            print(f"💎 PROCESSO CONCLUÍDO COM SUCESSO!")
            print(f"📊 Total: {len(resultados)} processos.")
            print(f"📂 Arquivo gerado: {arquivo_saida}")

    except Exception as e:
        print(f"⚠️ Erro ao salvar Excel formatado: {e}")
        # Fallback simples caso o xlsxwriter não esteja instalado
        pd.DataFrame(resultados).to_excel(arquivo_saida, index=False)
    
    print("="*30)

if __name__ == "__main__":
    executar()