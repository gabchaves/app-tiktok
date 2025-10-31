import pyautogui
import time
import pyperclip
import google.generativeai as genai
import webbrowser
import json
import xlsxwriter



def salvar_planilha(response_text):
    """Salva os temas extraídos em uma planilha Excel usando xlsxwriter."""
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
        
        # Cria planilha Excel
        workbook = xlsxwriter.Workbook('planilha_temas.xlsx')
        worksheet = workbook.add_worksheet()
        
        # Cabeçalhos
        headers = ['Tema', 'Descrição', 'Relevância']
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
        
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        
        # Dados
        for row, tema in enumerate(temas, start=1):
            if isinstance(tema, dict):
                worksheet.write(row, 0, tema.get('tema', tema.get('Tema', '')))
                worksheet.write(row, 1, tema.get('descricao', tema.get('Descrição', '')))
                worksheet.write(row, 2, tema.get('relevancia', tema.get('Relevância', '')))
        
        # Ajusta largura das colunas
        worksheet.set_column(0, 0, 30)
        worksheet.set_column(1, 1, 50)
        worksheet.set_column(2, 2, 15)
        
        workbook.close()
        print(f"\n✅ Planilha 'planilha_temas.xlsx' salva com sucesso! ({len(temas)} temas encontrados)")
        
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

