import pyautogui
import time
import pyperclip
import google.generativeai as genai
import webbrowser
import json
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
import os


def salvar_planilha(response_text):
    """Adiciona os temas extraídos à planilha Excel existente ou cria uma nova."""
    try:
        # Extrai JSON do texto (remove markdown code blocks se existirem)
        texto_limpo = response_text.strip()
        if '```' in texto_limpo:
            inicio = texto_limpo.find('{')
            fim = texto_limpo.rfind('}') + 1
            if inicio != -1 and fim > inicio:
                texto_limpo = texto_limpo[inicio:fim]
        
        dados_json = json.loads(texto_limpo)
        temas = dados_json.get('top_themes', [])
        
        if not temas:
            print("⚠️ Nenhum tema encontrado no JSON.")
            return
        
        arquivo_planilha = 'planilha_temas.xlsx'
        headers = ['Tema', 'Descrição', 'Relevância', 'Roteiro', 'Video Pronto', 'Video Postado', 'Data']
        
        # Verifica se o arquivo existe
        if os.path.exists(arquivo_planilha):
            workbook = load_workbook(arquivo_planilha)
            worksheet = workbook.active
            
            # Garante que os cabeçalhos existam (atualiza se necessário)
            if worksheet.max_row == 0 or worksheet.cell(1, 1).value != 'Tema':
                for col, header in enumerate(headers, start=1):
                    cell = worksheet.cell(1, col)
                    cell.value = header
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
            
            # Encontra a próxima linha vazia
            proxima_linha = worksheet.max_row + 1
        else:
            # Cria nova planilha
            workbook = Workbook()
            worksheet = workbook.active
            
            # Adiciona cabeçalhos
            for col, header in enumerate(headers, start=1):
                cell = worksheet.cell(1, col)
                cell.value = header
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
            
            proxima_linha = 2
        
        # Adiciona os novos temas
        for tema in temas:
            if isinstance(tema, dict):
                worksheet.cell(proxima_linha, 1, tema.get('tema', tema.get('Tema', '')))
                worksheet.cell(proxima_linha, 2, tema.get('descricao', tema.get('Descrição', '')))
                worksheet.cell(proxima_linha, 3, tema.get('relevancia', tema.get('Relevância', '')))
                # Deixa Roteiro, Video Pronto, Video Postado e Data em branco
                worksheet.cell(proxima_linha, 4, '')  # Roteiro
                worksheet.cell(proxima_linha, 5, '')  # Video Pronto
                worksheet.cell(proxima_linha, 6, '')  # Video Postado
                worksheet.cell(proxima_linha, 7, '')  # Data
                proxima_linha += 1
        
        # Ajusta largura das colunas
        worksheet.column_dimensions['A'].width = 30
        worksheet.column_dimensions['B'].width = 50
        worksheet.column_dimensions['C'].width = 15
        worksheet.column_dimensions['D'].width = 50  # Roteiro
        worksheet.column_dimensions['E'].width = 15  # Video Pronto
        worksheet.column_dimensions['F'].width = 15  # Video Postado
        worksheet.column_dimensions['G'].width = 12  # Data
        
        workbook.save(arquivo_planilha)
        print(f"\n✅ Planilha atualizada com sucesso! ({len(temas)} tema(s) adicionado(s))")
        
    except json.JSONDecodeError as e:
        print(f"❌ Erro ao fazer parse do JSON: {e}")
    except Exception as e:
        print(f"❌ Erro ao salvar planilha: {e}")


# Abre TikTok Studio
url = 'https://www.tiktok.com/tiktokstudio/inspiration'
webbrowser.open(url)
time.sleep(10)

# Navega até o conteúdo
for _ in range(28):
    pyautogui.press('tab')
    time.sleep(0.2)

pyautogui.press('enter')
time.sleep(2)

# Copia o conteúdo
pyautogui.hotkey('ctrl', 'a')
pyautogui.hotkey('ctrl', 'c')
time.sleep(0.5)

conteudo = pyperclip.paste()

# Configura e usa API do Gemini
GEMINI_API_KEY = "AIzaSyDZ_6FweRyBza_TuiWQ1W9zgubhfzHqRyY"

if not GEMINI_API_KEY:
    print("❌ Erro: GEMINI_API_KEY não foi definida.")
else:
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel('gemini-2.5-flash')

        prompt = f"""Analise o texto a seguir e identifique os 3 temas mais relevantes.
Retorne a resposta em formato JSON, com a seguinte estrutura:
{{
  "top_themes": [
    {{"tema": "nome do tema", "descricao": "explicação", "relevancia": "alta|média|baja"}},
    {{"tema": "nome do tema", "descricao": "explicação", "relevancia": "alta|média|baja"}},
    {{"tema": "nome do tema", "descricao": "explicação", "relevancia": "alta|média|baja"}}
  ]
}}

Texto para análise:
---
{conteudo}
---"""

        print("\n🤖 Enviando texto para análise do Gemini...")
        response = model.generate_content(prompt)
        
        print("\n--- Análise do Gemini ---")
        print(response.text)
        print("--- Fim da Análise ---\n")

        salvar_planilha(response.text)

    except Exception as e:
        if "API key" in str(e):
            print("❌ Erro de autenticação com a API do Gemini. Verifique sua API Key.")
        else:
            print(f"❌ Erro ao usar a API do Gemini: {e}")

print("\n✅ Processo concluído.")

