from flask import Flask, request, render_template, send_from_directory
from docx import Document
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

def formatar_valor(valor):
    return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory('static/outputs', filename, as_attachment=True)

@app.route('/atualizar_word', methods=['POST'])
def atualizar_word():
    try:
        nome_cliente = request.form.get('nome', 'Cliente não informado')
        endereco = request.form.get('endereco', 'Endereço não informado')
        dias = request.form.get('dias', '0')
        numero_pedido = request.form.get('n_pedido', 'Pedido não informado')
        data_atual = request.form.get('data_atual', 'Data não informada')
        data_prevista = request.form.get('data_prevista', 'Data não informada')
        data_vencimento = request.form.get('data_vencimento', 'Data não informada')
        forma_pagamento = request.form.get('forma_pagamento', 'Forma de pagamento não informada')

        data_atual = datetime.now().strftime('%d/%m/%Y')
        data_prevista = datetime.strptime(request.form.get('data_prevista'), '%Y-%m-%d').strftime('%d/%m/%Y')
        data_vencimento = datetime.strptime(request.form.get('data_vencimento'), '%Y-%m-%d').strftime('%d/%m/%Y')

        try:
            dias = int(dias)
        except ValueError:
            return "Erro: 'Dias' deve ser um número inteiro válido.", 400

        codigos = request.form.getlist('codigo[]')
        descricoes = request.form.getlist('descricao[]')
        unidades = request.form.getlist('un[]')
        quantidades = request.form.getlist('qtd[]')
        valores_unitarios = request.form.getlist('valor_unit[]')

        if not (len(codigos) == len(descricoes) == len(unidades) == len(quantidades) == len(valores_unitarios)):
            return "Erro: Todos os itens devem ser preenchidos e ter o mesmo comprimento.", 400

        itens = []
        numero_itens = 0
        soma_quantidades = 0
        total_produtos = 0.0
        for i, (codigo, descricao, un, qtd, valor_unit) in enumerate(zip(codigos, descricoes, unidades, quantidades, valores_unitarios)):
            try:
                qtd = int(qtd)
                valor_unit = float(valor_unit)
                total = qtd * valor_unit
                itens.append((codigo, descricao, un, qtd, formatar_valor(valor_unit), formatar_valor(total)))
                numero_itens += 1
                soma_quantidades += qtd
                total_produtos += total
            except ValueError:
                return f"Erro: Dados inválidos no item {i+1}. Certifique-se de que as quantidades e valores unitários sejam numéricos.", 400

        total_pedido = total_produtos
        arquivo_entrada = 'ORÇAMENTO_LUCAS.docx'
        arquivo_saida = f'static/outputs/ORÇAMENTO_{nome_cliente}.docx'
        editar_word(arquivo_entrada, arquivo_saida, nome_cliente, endereco, dias, numero_pedido, data_atual, data_prevista, data_vencimento, forma_pagamento, itens, numero_itens, soma_quantidades, formatar_valor(total_pedido), formatar_valor(total_produtos))
        return f"Documento Word atualizado com sucesso! Faça o download <a href='/download/{arquivo_saida}'>aqui</a>."
    except Exception as e:
        return f"Erro ao processar o documento Word: {e}", 500

def editar_word(arquivo_entrada, arquivo_saida, nome, endereco, dias, numero_pedido, data_atual, data_prevista, data_vencimento, forma_pagamento, itens, numero_itens, soma_quantidades, total_pedido, total_produtos):
    try:
        doc = Document(arquivo_entrada)
        for paragraph in doc.paragraphs:
            if '[nome]' in paragraph.text:
                paragraph.text = paragraph.text.replace('[nome]', nome)
            # Continue para outros placeholders...
        doc.save(arquivo_saida)
    except Exception as e:
        raise RuntimeError(f"Erro no processo de edição do Word: {e}")

if __name__ == '__main__':
    app.run(debug=True)
