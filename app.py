from flask import Flask, request, render_template
import pythoncom
from win32com.client import Dispatch
from datetime import datetime
app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

def formatar_valor(valor):
    return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

@app.route('/atualizar_word', methods=['POST'])
def atualizar_word():
    try:
        # Coletar dados do formulário
        nome_cliente = request.form.get('nome', 'Cliente não informado')
        endereco = request.form.get('endereco', 'Endereço não informado')
        dias = request.form.get('dias', '0')
        numero_pedido = request.form.get('n_pedido', 'Pedido não informado')
        data_atual = request.form.get('data_atual', 'Data não informada')
        data_prevista = request.form.get('data_prevista', 'Data não informada')
        data_vencimento = request.form.get('data_vencimento', 'Data não informada')
        forma_pagamento = request.form.get('forma_pagamento', 'Forma de pagamento não informada')
        

        # Formatar datas recebidas do formulário
        data_atual = datetime.now().strftime('%d/%m/%Y')  # Data atual no formato desejado
        data_prevista = datetime.strptime(request.form.get('data_prevista'), '%Y-%m-%d').strftime('%d/%m/%Y')
        data_vencimento = datetime.strptime(request.form.get('data_vencimento'), '%Y-%m-%d').strftime('%d/%m/%Y')
        
        # Verificar se 'dias' é um número válido
        try:
            dias = int(dias)
        except ValueError:
            return "Erro: 'Dias' deve ser um número inteiro válido.", 400
        
        # Coletar itens como listas
        codigos = request.form.getlist('codigo[]')
        descricoes = request.form.getlist('descricao[]')
        unidades = request.form.getlist('un[]')
        quantidades = request.form.getlist('qtd[]')
        valores_unitarios = request.form.getlist('valor_unit[]')
        
        # Verificar se todas as listas têm o mesmo comprimento
        if not (len(codigos) == len(descricoes) == len(unidades) == len(quantidades) == len(valores_unitarios)):
            return "Erro: Todos os itens devem ser preenchidos e ter o mesmo comprimento.", 400
        
        # Calcular os valores totais para cada item
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
                
                # Atualizar cálculos
                numero_itens += 1
                soma_quantidades += qtd
                total_produtos += total
                
            except ValueError:
                return f"Erro: Dados inválidos no item {i+1}. Certifique-se de que as quantidades e valores unitários sejam numéricos.", 400
        
        # Total do pedido é o mesmo que total de produtos
        total_pedido = total_produtos
        
        # Caminhos dos arquivos
        arquivo_entrada = r'C:\Users\bruno.salles\Desktop\Lucas Orçamento\ORÇAMENTO_LUCAS.docx'
        arquivo_saida = f'C:\\Users\\bruno.salles\\Desktop\\Lucas Orçamento\\ORÇAMENTO_{nome_cliente}.docx'
        
        editar_word(arquivo_entrada, arquivo_saida, nome_cliente, endereco, dias, numero_pedido, data_atual, data_prevista, data_vencimento, forma_pagamento, itens, numero_itens, soma_quantidades, formatar_valor(total_pedido), formatar_valor(total_produtos))
        
        return f"Documento Word atualizado com sucesso! Salvo como {arquivo_saida}."
    except Exception as e:
        return f"Erro ao processar o documento Word: {e}", 500

def editar_word(arquivo_entrada, arquivo_saida, nome, endereco, dias, numero_pedido, data_atual, data_prevista, data_vencimento, forma_pagamento, itens, numero_itens, soma_quantidades, total_pedido, total_produtos):
    try:
        pythoncom.CoInitialize()
        word = Dispatch("Word.Application")
        
        # Abrir o documento
        try:
            doc = word.Documents.Open(arquivo_entrada)
        except Exception as e:
            raise FileNotFoundError(f"Erro ao abrir o arquivo: {e}")
        
        # Substituir texto genérico (como [nome], [endereco], [dias])
        for paragraph in doc.Paragraphs:
            try:
                if '[nome]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[nome]', nome)
                if '[endereco]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[endereco]', endereco)
                if '[dias]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[dias]', str(dias))
                if '[n_pedido]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[n_pedido]', numero_pedido)
                if '[data_atual]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[data_atual]', data_atual)
                if '[data_prevista]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[data_prevista]', data_prevista)
                if '[data_vencimento]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[data_vencimento]', data_vencimento)
                if '[forma_pagamento]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[forma_pagamento]', forma_pagamento)
                if '[n_itens]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[n_itens]', f"{int(numero_itens):,.1f}")
                if '[soma_qtd]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[soma_qtd]', f"{int(soma_quantidades):,.1f}")
                if '[total_prod]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[total_prod]', total_produtos)
                if '[total_geral]' in paragraph.Range.Text:
                    paragraph.Range.Text = paragraph.Range.Text.replace('[total_geral]', total_pedido)
            except Exception as e:
                raise RuntimeError(f"Erro ao substituir texto no parágrafo: {e}")
        
        # Adicionar itens na tabela
        for table in doc.Tables:
            try:
                # Procura a linha com o placeholder e adiciona os itens abaixo dela
                for row in table.Rows:
                    if '[descricao]' in row.Cells[0].Range.Text:
                        placeholder_row = row
                        # Adiciona os itens abaixo da linha do placeholder
                        for item in itens:
                            codigo, descricao, un, qtd, valor_unit, total = item
                            new_row = table.Rows.Add()  # Adiciona uma nova linha
                            new_row.Cells[0].Range.Text = descricao
                            new_row.Cells[1].Range.Text = codigo
                            new_row.Cells[2].Range.Text = un
                            new_row.Cells[3].Range.Text = f"{int(qtd):,.1f}"
                            new_row.Cells[4].Range.Text = valor_unit
                            new_row.Cells[5].Range.Text = total
                        # Remove o placeholder da linha original para evitar duplicação
                        placeholder_row.Delete()
                        break  # Sai do loop após adicionar os itens
            except Exception as e:
                raise RuntimeError(f"Erro ao adicionar itens na tabela: {e}")
        
        # Substituir texto dentro de formas geométricas
        for shape in doc.Shapes:
            try:
                if shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text
                    if '[nome]' in text:
                        shape.TextFrame.TextRange.Text = text.replace('[nome]', nome)
            except Exception as e:
                raise RuntimeError(f"Erro ao substituir texto na forma geométrica: {e}")
        
        # Salvar o documento em DOCX
        try:
            doc.SaveAs(arquivo_saida)
        except Exception as e:
            raise IOError(f"Erro ao salvar o arquivo: {e}")
        
        # Salvar o documento em PDF
        try:
            arquivo_saida_pdf = arquivo_saida.replace('.docx', '.pdf')
            doc.SaveAs(arquivo_saida_pdf, FileFormat=17)  # 17 é o código para PDF
        except Exception as e:
            raise IOError(f"Erro ao salvar o arquivo em PDF: {e}")
        finally:
            doc.Close()
            word.Quit()
    except Exception as e:
        raise RuntimeError(f"Erro no processo de edição do Word: {e}")

if __name__ == '__main__':
    app.run(debug=True)