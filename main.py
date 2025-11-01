from Tema_generator import GeradorDeTemas
from video_creator import CriadorDeVideo
import openpyxl
import os
from datetime import datetime

class App:
    def __init__(self, planilha_path='planilha_temas.xlsx'):
        self.planilha_path = planilha_path

    def _ler_proximo_tema(self):
        """Lê a planilha e retorna o primeiro tema que ainda não tem um roteiro."""
        if not os.path.exists(self.planilha_path):
            return None, -1

        workbook = openpyxl.load_workbook(self.planilha_path)
        worksheet = workbook.active

        # Itera a partir da segunda linha (pulando o cabeçalho)
        for row in range(2, worksheet.max_row + 1):
            tema = worksheet.cell(row=row, column=1).value
            roteiro = worksheet.cell(row=row, column=4).value
            if tema and not roteiro:
                print(f"🔍 Tema encontrado para criar vídeo: '{tema}'")
                return tema, row
        
        return None, -1

    def _atualizar_planilha(self, row_num, roteiro, video_criado=False):
        """Atualiza o status do tema na planilha."""
        workbook = openpyxl.load_workbook(self.planilha_path)
        worksheet = workbook.active
        
        # Adiciona o roteiro (simulado por enquanto)
        worksheet.cell(row=row_num, column=4).value = roteiro
        
        # Marca se o vídeo foi criado
        if video_criado:
            worksheet.cell(row=row_num, column=5).value = "Sim"
            worksheet.cell(row=row_num, column=7).value = datetime.now().strftime("%Y-%m-%d %H:%M")

        workbook.save(self.planilha_path)
        print(f"📊 Planilha atualizada para a linha {row_num}.")

    def executar(self):
        """Executa o fluxo principal da aplicação."""
        print("--- Iniciando Automação TikTok ---")

        # 1. Gera novos temas se não houver nenhum pendente
        proximo_tema, _ = self._ler_proximo_tema()
        if not proximo_tema:
            print("\nNenhum tema pendente encontrado. Buscando novos temas...")
            gerador_temas = GeradorDeTemas(planilha=self.planilha_path)
            gerador_temas.executar()
            # Tenta ler novamente após a geração
            proximo_tema, _ = self._ler_proximo_tema()

        # 2. Processa o próximo tema disponível
        if proximo_tema:
            tema, linha = self._ler_proximo_tema()
            if tema:
                # Etapa de gerar roteiro (a ser implementada)
                print(f"\n📝 Gerando roteiro para o tema: '{tema}' (simulação)")
                roteiro_gerado = f"Este é um roteiro sobre {tema}."
                
                # 3. Cria o vídeo
                criador_video = CriadorDeVideo(roteiro=roteiro_gerado)
                sucesso = criador_video.executar()
                
                # 4. Atualiza a planilha
                if sucesso:
                    self._atualizar_planilha(linha, roteiro_gerado, video_criado=True)
        else:
            print("\nNenhum tema para processar, mesmo após a busca. Encerrando.")

        print("\n--- Automação TikTok Concluída ---")

if __name__ == "__main__":
    app = App()
    app.executar()
