from flask import Flask, request, render_template, send_from_directory, url_for
from docx import Document
from datetime import datetime
import os
import re
import random

app = Flask(__name__)

# ==============================
# CONFIG
# ==============================
OUTPUT_DIR = os.path.join("static", "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

TEMPLATE_DOCX = "ORÇAMENTO_LUCAS.docx"


# ==============================
# GERADOR DE PEDIDO (4 dígitos)
# ==============================
def gerar_numero_pedido():
    return f"{random.randint(0, 9999):04d}"


@app.route("/")
def index():
    return render_template("index.html", pedido_auto=gerar_numero_pedido())


# ==============================
# HELPERS
# ==============================
def formatar_valor(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def nome_arquivo_seguro(nome: str) -> str:
    if not nome:
        return "CLIENTE"
    nome = nome.strip()
    nome = re.sub(r'[\\/:*?"<>|]+', "", nome)
    nome = re.sub(r"\s+", " ", nome)
    return nome or "CLIENTE"


def parse_data_form(campo: str) -> str:
    if not campo:
        return "Data não informada"
    try:
        return datetime.strptime(campo, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return "Data inválida"


# ==============================
# DOWNLOAD
# ==============================
@app.route("/download/<path:filename>")
def download_file(filename):
    filename = os.path.basename(filename)  # evita path traversal
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


# =====================================================
# 🔥 REPLACE DEFINITIVO (PEGA TEXTBOX/SHAPES TAMBÉM)
# =====================================================
def replace_in_xml_anywhere(doc: Document, replacements: dict):
    """
    Varre TODOS os nós do XML do DOCX e substitui em qualquer <w:t>.
    Isso pega:
    - parágrafos normais
    - tabelas
    - cabeçalho/rodapé
    - caixas de texto (TextBox/Shape)
    - muitos content controls
    """
    def _replace_in_element(element):
        for node in element.iter():
            # node.tag termina com '}t' nos textos (w:t)
            if isinstance(node.tag, str) and node.tag.endswith("}t") and node.text:
                txt = node.text
                for k, v in replacements.items():
                    txt = txt.replace(k, str(v))
                node.text = txt

    # Corpo
    _replace_in_element(doc.element)

    # Cabeçalho/Rodapé (todas as seções)
    for section in doc.sections:
        _replace_in_element(section.header._element)
        _replace_in_element(section.footer._element)

        # alguns modelos usam first_page_header/footer
        try:
            _replace_in_element(section.first_page_header._element)
            _replace_in_element(section.first_page_footer._element)
        except Exception:
            pass


# ==============================
# ITENS NA TABELA
# ==============================
def fill_items_table(doc: Document, itens):
    """
    Procura a linha MODELO que tenha [descricao] e [codigo],
    duplica para cada item e remove a linha modelo.
    """
    for table in doc.tables:
        for row in table.rows:
            row_text = " ".join(cell.text for cell in row.cells)

            if "[descricao]" in row_text and "[codigo]" in row_text:
                model_row = row

                for (codigo, descricao, un, qtd_int, valor_unit, total) in itens:
                    new_row = table.add_row()

                    # Copia conteúdo da linha modelo
                    for c in range(len(model_row.cells)):
                        new_row.cells[c].text = model_row.cells[c].text

                    rep_item = {
                        "[descricao]": descricao,
                        "[codigo]": codigo,
                        "[un]": un,
                        "[qtd]": str(qtd_int),
                        "[valor_unit]": str(valor_unit),
                        "[total]": str(total),
                    }

                    # Faz replace por XML dentro das células da nova linha
                    for cell in new_row.cells:
                        for node in cell._element.iter():
                            if isinstance(node.tag, str) and node.tag.endswith("}t") and node.text:
                                t = node.text
                                for k, v in rep_item.items():
                                    t = t.replace(k, str(v))
                                node.text = t

                # Remove a linha modelo
                table._tbl.remove(model_row._tr)
                return


# ==============================
# ROTA PRINCIPAL
# ==============================
@app.route("/atualizar_word", methods=["POST"])
def atualizar_word():
    try:
        nome_cliente = (request.form.get("nome") or "").strip()
        endereco = (request.form.get("endereco") or "").strip()
        dias_raw = (request.form.get("dias") or "0").strip()
        numero_pedido = (request.form.get("n_pedido") or "").strip()
        forma_pagamento = (request.form.get("forma_pagamento") or "").strip()

        # Se não vier pedido, gera automático
        if not numero_pedido:
            numero_pedido = gerar_numero_pedido()

        # Datas
        data_atual = datetime.now().strftime("%d/%m/%Y")
        data_prevista = parse_data_form(request.form.get("data_prevista"))
        data_vencimento = parse_data_form(request.form.get("data_vencimento"))

        # Dias
        try:
            dias = int(dias_raw or "0")
        except ValueError:
            return "Erro: 'Dias' deve ser número válido.", 400

        # Itens (listas)
        codigos = request.form.getlist("codigo[]")
        descricoes = request.form.getlist("descricao[]")
        unidades = request.form.getlist("un[]")
        quantidades = request.form.getlist("qtd[]")
        valores_unitarios = request.form.getlist("valor_unit[]")

        if not (len(codigos) == len(descricoes) == len(unidades) == len(quantidades) == len(valores_unitarios)):
            return "Erro: itens com tamanhos diferentes (campos faltando).", 400

        itens = []
        numero_itens = 0
        soma_quantidades = 0
        total_produtos = 0.0

        for i, (codigo, descricao, un, qtd, valor_unit) in enumerate(
            zip(codigos, descricoes, unidades, quantidades, valores_unitarios), start=1
        ):
            try:
                qtd_int = int(qtd)
                valor_float = float((valor_unit or "0").replace(",", "."))
            except ValueError:
                return f"Erro: Item {i} com quantidade/valor inválidos.", 400

            total = qtd_int * valor_float

            itens.append((
                codigo,
                descricao,
                un,
                qtd_int,
                formatar_valor(valor_float),
                formatar_valor(total),
            ))

            numero_itens += 1
            soma_quantidades += qtd_int
            total_produtos += total

        total_geral = total_produtos

        # Arquivo de saída
        cliente_safe = nome_arquivo_seguro(nome_cliente)
        nome_saida = f"ORÇAMENTO_{cliente_safe}.docx"
        arquivo_saida = os.path.join(OUTPUT_DIR, nome_saida)

        # Edita e salva
        editar_word(
            arquivo_entrada=TEMPLATE_DOCX,
            arquivo_saida=arquivo_saida,
            nome=nome_cliente,
            endereco=endereco,
            dias=dias,
            numero_pedido=numero_pedido,
            data_atual=data_atual,
            data_prevista=data_prevista,
            data_vencimento=data_vencimento,
            forma_pagamento=forma_pagamento,
            itens=itens,
            numero_itens=numero_itens,
            soma_qtd=soma_quantidades,
            total_prod=formatar_valor(total_produtos),
            total_geral=formatar_valor(total_geral),
        )

        download_url = url_for("download_file", filename=nome_saida)
        return f"Documento Word atualizado com sucesso! Faça o download <a href='{download_url}'>aqui</a>."

    except Exception as e:
        return f"Erro ao processar o documento Word: {e}", 500


# ==============================
# EDIÇÃO DO WORD
# ==============================
def editar_word(
    arquivo_entrada,
    arquivo_saida,
    nome,
    endereco,
    dias,
    numero_pedido,
    data_atual,
    data_prevista,
    data_vencimento,
    forma_pagamento,
    itens,
    numero_itens,
    soma_qtd,
    total_prod,
    total_geral,
):
    doc = Document(arquivo_entrada)

    # Placeholders exatamente como no seu template
    replacements = {
        "[nome]": nome,
        "[endereco]": endereco,
        "[n_pedido]": numero_pedido,
        "[data_atual]": data_atual,
        "[data_prevista]": data_prevista,
        "[dias]": str(dias),
        "[data_vencimento]": data_vencimento,
        "[forma_pagamento]": forma_pagamento,
        "[n_itens]": str(numero_itens),
        "[soma_qtd]": str(soma_qtd),
        "[total_prod]": str(total_prod),
        "[total_geral]": str(total_geral),
    }

    # ✅ Isso pega o [nome] na TextBox
    replace_in_xml_anywhere(doc, replacements)

    # Preenche itens (tabela)
    fill_items_table(doc, itens)

    doc.save(arquivo_saida)


if __name__ == "__main__":
    app.run(debug=True)