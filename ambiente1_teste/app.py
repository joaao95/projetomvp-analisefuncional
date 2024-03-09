import requests
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt
import os

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    if request.method == 'POST':
        erp = request.form['erp']
        processos = request.form.getlist('processos')
        informacoes_erp = request.form.getlist('informacoes_erp')

        # Formular a pergunta para o ChatGPT
        pergunta = (f"Qual seria a melhor análise funcional para realizar a integração do ERP {erp} com o nosso SaaS Paytrack? "
            f"Processos desejados: {', '.join(processos)}. "
            f"Informações necessárias pelo ERP repassadas pelo cliente: {', '.join(informacoes_erp)}. "
            "Preciso que no retorno do documento contenha um mapeamento de campos olhando para o ERP selecionado, "
            "além de me retornar um JSON de exemplo. "
            "É importante salientar, que o mapeamento de campos precisa vir no formato tabela e o JSON, "
            "Precisa ser formatado com as nomenclaturas do ERP, "
            "como por exemplo bukrs que significa a empresa para o SAP e assim sucessivamente. "
            "Seguem algumas diretrizes que o nosso sistema Paytrack tem para integrar com os ERPs. "
            "1. - Utilizamos comunicação Sincrona com os Webservices do cliente. "
            "2. - A Paytrack é ativa nas integrações, ou seja, após este passo o cliente irá disponibilizar um Webservice para consumirmos. "
            "3. - A Análise funcional precisará ser separada no documento por cenário selecionado, ou seja, uma análise para adiantamento, uma para prestação de contas etc.")


        headers = {'Authorization': 'Bearer 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'}
        data = {
            "model": "gpt-3.5-turbo",
            "messages": [{"role": "user", "content": pergunta}],
            "temperature": 0.7
        }

        response = requests.post('https://api.openai.com/v1/chat/completions', headers=headers, json=data)
        
        if response.status_code == 200:
            resposta_chatgpt = response.json()['choices'][0]['message']['content']

            # Criar o documento com os dados do formulário e a resposta do ChatGPT
            document = Document()
            document.add_heading('Integração ERP com Paytrack', 0)
            document.add_paragraph(f"ERP: {erp}")
            document.add_paragraph(f"Processos desejados: {', '.join(processos)}")
            document.add_paragraph(f"Informações necessárias pelo ERP: {', '.join(informacoes_erp)}")
            document.add_heading('Análise Funcional Recomendada:', level=1)
            document.add_paragraph(resposta_chatgpt)

            # Adicionando informações ERP selecionadas em uma tabela
            document.add_page_break()
            document.add_heading('Mapeamento de Campos ERP:', level=1)
            table = document.add_table(rows=1, cols=2)
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Campo'
            hdr_cells[1].text = 'Descrição'
            for campo in informacoes_erp:
                row_cells = table.add_row().cells
                row_cells[0].text = campo
                row_cells[1].text = "Descrição para " + campo  # Aqui você pode personalizar as descrições

            # Salvar o documento
            diretorio_base = os.path.abspath(os.path.dirname(__file__))
            filename = os.path.join(diretorio_base, "analise_integracao.docx")
            document.save(filename)

            # Enviar o arquivo gerado para o usuário
            return send_file(filename, as_attachment=True)
        else:
            print(response.json())  # Ajuda a depurar em caso de erro na API
            return "Erro ao comunicar com a API do ChatGPT", 500

if __name__ == '__main__':
    app.run(debug=True)
