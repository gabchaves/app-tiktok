import pyautogui
import time
import pyperclip
import google.generativeai as genai
import webbrowser
import os

# --- Interação com a Web ---
# Abrir a URL no navegador padrão. É uma abordagem mais robusta que simular o 'Executar'.
url = 'https://www.tiktok.com/tiktokstudio/inspiration'
webbrowser.open(url)

# ATENÇÃO: O uso de time.sleep e a navegação por 'tab' são frágeis.
# Uma pequena mudança na página pode quebrar o script.
# Para uma automação mais robusta, considere usar bibliotecas como Selenium ou Playwright,
# que permitem esperar por elementos específicos da página.
time.sleep(10)  # Aumentado para dar mais tempo para a página carregar.

# Navegação por 'tab' para chegar ao conteúdo.
# Este número de 'tabs' é um "chute" e provavelmente precisará de ajuste.
for _ in range(28):
    pyautogui.press('tab')
    time.sleep(0.2)

pyautogui.press('enter')
time.sleep(2)

# Selecionar e copiar o conteúdo
pyautogui.hotkey('ctrl', 'a')
pyautogui.hotkey('ctrl', 'c')
time.sleep(0.5)

# --- Processamento do Conteúdo ---
conteudo = pyperclip.paste()

with open('conteudo_copiado.txt', 'w', encoding='utf-8') as arquivo:
    arquivo.write(conteudo)

print("Conteúdo salvo em 'conteudo_copiado.txt'")
print(f"Conteúdo (primeiros 1000 caracteres): {conteudo[:1000]}...")


GEMINI_API_KEY = "AIzaSyDZ_6FweRyBza_TuiWQ1W9zgubhfzHqRyY"

if not GEMINI_API_KEY:
    print("❌ Erro: A variável de ambiente GEMINI_API_KEY não foi definida.")
else:
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel('gemini-2.5-flash')

        prompt = f"""
Analise o texto a seguir e identifique os 3 temas mais relevantes.
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
---
"""

        print("\n🤖 Enviando texto para análise do Gemini...")
        response = model.generate_content(prompt)

        print("\n--- Análise do Gemini ---")
        print(response.text)
        print("--- Fim da Análise ---")

    except Exception as e:
        if "API key" in str(e):
            print("❌ Erro de autenticação com a API do Gemini. Verifique sua API Key.")
        else:
            print(f"❌ Ocorreu um erro ao usar a API do Gemini: {e}")


print("\n✅ Processo concluído.")