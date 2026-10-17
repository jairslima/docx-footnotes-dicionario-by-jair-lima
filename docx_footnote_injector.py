#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Footnote Injector by Jair Lima
Este script realiza a injeção automática de notas de rodapé em arquivos Word (.docx)
com base em um dicionário de termos técnicos e definições.
Preserva a formatação original do documento e constrói dinamicamente a
infraestrutura de notas de rodapé caso ela não exista no arquivo original.
"""

import os
import sys
import re
import zipfile
import tempfile
import shutil
import copy
import json
import csv
from lxml import etree


def resource_path(relative_path):
    """Obtém o caminho absoluto para o recurso, compatível com PyInstaller e desenvolvimento."""
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_run_xml_text(r_elem, new_text):
    """Define o texto de um elemento run XML (<w:r>) preservando espaços."""
    t_elem = r_elem.find(qn('w:t'))
    if t_elem is None:
        t_elem = OxmlElement('w:t')
        r_elem.append(t_elem)
    t_elem.text = new_text
    t_elem.set(qn('xml:space'), 'preserve')


def get_paragraph_runs_offsets(paragraph):
    """Calcula os offsets de início e fim de cada run no texto consolidado do parágrafo."""
    runs_offsets = []
    current_offset = 0
    for r in paragraph.runs:
        text_len = len(r.text)
        runs_offsets.append((r, current_offset, current_offset + text_len))
        current_offset += text_len
    return runs_offsets


def merge_identical_runs(paragraph):
    """
    Mescla runs adjacentes idênticos (com a mesma formatação) no parágrafo.
    Isso ajuda a unificar palavras que foram quebradas artificialmente pelo Word.
    """
    p_elem = paragraph._p
    runs = list(paragraph.runs)
    if len(runs) < 2:
        return

    i = 0
    while i < len(runs) - 1:
        r1 = runs[i]
        r2 = runs[i + 1]
        
        r1_elem = r1._r
        r2_elem = r2._r

        # Verifica se ambos contêm propriedades equivalentes
        rPr1 = r1_elem.find(qn('w:rPr'))
        rPr2 = r2_elem.find(qn('w:rPr'))
        
        rPr1_str = etree.tostring(rPr1) if rPr1 is not None else b''
        rPr2_str = etree.tostring(rPr2) if rPr2 is not None else b''

        # Verifica se contêm apenas w:t além de w:rPr
        allowed_tags = {qn('w:rPr'), qn('w:t')}
        r1_valid = all(child.tag in allowed_tags for child in r1_elem)
        r2_valid = all(child.tag in allowed_tags for child in r2_elem)

        if rPr1_str == rPr2_str and r1_valid and r2_valid:
            # Mescla r2 em r1
            t1 = r1_elem.find(qn('w:t'))
            t2 = r2_elem.find(qn('w:t'))
            
            text2 = t2.text if t2 is not None and t2.text else ""
            if t1 is None:
                t1 = OxmlElement('w:t')
                r1_elem.append(t1)
            
            t1.text = (t1.text if t1.text else "") + text2
            t1.set(qn('xml:space'), 'preserve')
            
            # Remove r2 da árvore XML e da lista de trabalho
            p_elem.remove(r2_elem)
            runs.pop(i + 1)
        else:
            i += 1


class FootnoteInjector:
    """Injetor de notas de rodapé automático baseado em dicionário."""

    # Estilos de cabeçalho comuns a serem ignorados no processamento de notas de rodapé
    IGNORED_STYLES = {
        "title", "subtitle", "heading 1", "heading 2", "heading 3", "heading 4", 
        "heading 5", "heading 6", "título 1", "título 2", "título 3", "título 4", 
        "título 5", "título 6", "título", "subtítulo", "caption", "legenda"
    }

    def __init__(self, dictionary_data):
        """
        Inicializa o injetor com os dados do dicionário.
        dictionary_data: dict de {termo: definicao}
        """
        self.dictionary = dictionary_data
        self.footnotes_to_add = []  # Lista de tuplas (id, matched_text, texto)
        self.next_footnote_id = 1
        self.insertion_report = {}  # Relatório de contagem de inserções por termo

        
        # Mapeamento do termo normalizado para busca rápida
        # Ordena por tamanho decrescente para que termos compostos tenham prioridade sobre termos simples
        # Ex: "Banco de Dados" antes de "Banco"
        self.sorted_terms = sorted(self.dictionary.keys(), key=len, reverse=True)

    def inject(self, doc_path, output_path, only_first_occurrence=True):
        """
        Abre o documento, injeta as referências e finaliza as notas de rodapé no zip.
        """
        doc = Document(doc_path)
        terms_inserted = set()
        self.insertion_report = {}

        for paragraph in doc.paragraphs:
            # Ignora títulos e cabeçalhos com base no nome do estilo
            style_name = (paragraph.style.name or "").lower().strip()
            if any(ignored in style_name for ignored in self.IGNORED_STYLES):
                continue

            # Primeiro mescla runs adjacentes idênticos para evitar termos fracionados
            merge_identical_runs(paragraph)
            
            # Constrói o texto do parágrafo a partir dos runs
            runs = list(paragraph.runs)
            full_text = "".join(r.text for r in runs)
            if not full_text:
                continue

            # Encontra todas as correspondências do dicionário neste parágrafo
            matches = []
            for term in self.sorted_terms:
                if only_first_occurrence and term.lower() in terms_inserted:
                    continue

                # Cria expressão regular para correspondência de palavra inteira
                # Suporta termos com caracteres acentuados usando flag unicode nativa
                pattern = rf"\b{re.escape(term)}\b"
                for match in re.finditer(pattern, full_text, re.IGNORECASE):
                    start_pos, end_pos = match.span()
                    matched_text = match.group(0)
                    
                    # Verifica se não há sobreposição com correspondências já encontradas
                    # (como termos compostos que já englobaram a posição)
                    overlap = False
                    for existing_start, existing_end, _, _ in matches:
                        if not (end_pos <= existing_start or start_pos >= existing_end):
                            overlap = True
                            break
                    
                    if not overlap:
                        matches.append((start_pos, end_pos, term, matched_text))
                        if only_first_occurrence:
                            # Adiciona o termo normalizado na lista de inseridos
                            terms_inserted.add(term.lower())

            if not matches:
                continue

            # Processa as correspondências da direita para a esquerda (ordem decrescente de início)
            # Isso garante que a divisão dos runs à direita não mude os índices à esquerda
            matches.sort(key=lambda x: x[0], reverse=True)

            for start_pos, end_pos, term, matched_text in matches:
                # Recalcula offsets dos runs do parágrafo atualizado
                runs_offsets = get_paragraph_runs_offsets(paragraph)
                
                # Localiza o run que contém o termo
                target_run = None
                run_start = 0
                
                for r, r_start, r_end in runs_offsets:
                    if r_start <= start_pos and end_pos <= r_end:
                        target_run = r
                        run_start = r_start
                        break
                
                # Se o termo estiver em um único run (caso geral mais comum e seguro)
                if target_run is not None:
                    r_elem = target_run._r
                    parent_p = r_elem.getparent()
                    run_idx = parent_p.index(r_elem)
                    
                    rel_start = start_pos - run_start
                    rel_end = end_pos - run_start
                    
                    text_before = target_run.text[:rel_start]
                    text_term = target_run.text[rel_start:rel_end]
                    text_after = target_run.text[rel_end:]
                    
                    inserted_elements = []
                    
                    # 1. Texto anterior
                    if text_before:
                        r_before = copy.deepcopy(r_elem)
                        set_run_xml_text(r_before, text_before)
                        inserted_elements.append(r_before)
                        
                    # 2. O próprio termo
                    r_term = copy.deepcopy(r_elem)
                    set_run_xml_text(r_term, text_term)
                    inserted_elements.append(r_term)
                    
                    # 3. Referência da Nota de Rodapé
                    r_footnote = OxmlElement('w:r')
                    rPr = OxmlElement('w:rPr')
                    
                    rStyle = OxmlElement('w:rStyle')
                    rStyle.set(qn('w:val'), 'FootnoteReference')
                    rPr.append(rStyle)
                    
                    # Força sobrescrito (superscript) nas propriedades locais do run
                    vertAlign = OxmlElement('w:vertAlign')
                    vertAlign.set(qn('w:val'), 'superscript')
                    rPr.append(vertAlign)
                    
                    r_footnote.append(rPr)
                    
                    footnote_ref = OxmlElement('w:footnoteReference')
                    footnote_ref.set(qn('w:id'), str(self.next_footnote_id))
                    r_footnote.append(footnote_ref)
                    inserted_elements.append(r_footnote)
                    
                    # Registra a nota na fila
                    definition = self.dictionary[term]
                    self.footnotes_to_add.append((self.next_footnote_id, matched_text, definition))
                    self.next_footnote_id += 1
                    
                    # Atualiza o relatório de inserções
                    self.insertion_report[term] = self.insertion_report.get(term, 0) + 1
                    
                    # 4. Texto posterior
                    if text_after:
                        r_after = copy.deepcopy(r_elem)
                        set_run_xml_text(r_after, text_after)
                        inserted_elements.append(r_after)
                        
                    # Remove o run antigo e insere os novos em sequência
                    parent_p.remove(r_elem)
                    for idx, new_elem in enumerate(inserted_elements):
                        parent_p.insert(run_idx + idx, new_elem)

        # Salva o arquivo temporário inicializado pelo python-docx
        doc.save(output_path)
        
        # Realiza a finalização injetando a infraestrutura de notas de rodapé no ZIP
        self.finalize_footnotes_zip(output_path)

    def gerar_relatorio(self, output_report_path):
        """Gera um arquivo de relatório de texto com o resumo dos termos inseridos."""
        try:
            with open(output_report_path, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("RELATORIO DE INSERCAO DE NOTAS DE RODAPE\n")
                f.write("Projeto by Jair Lima - Footnote Injector\n")
                f.write("=" * 60 + "\n\n")
                f.write(f"Total de notas de rodape adicionadas: {self.next_footnote_id - 1}\n")
                f.write(f"Total de termos unicos identificados: {len(self.insertion_report)}\n\n")
                f.write("Detalhamento de Termos Inseridos:\n")
                f.write("-" * 45 + "\n")
                f.write(f"{'Termo':<25} | {'Quantidade de Insercoes':<15}\n")
                f.write("-" * 45 + "\n")
                for term, count in sorted(self.insertion_report.items()):
                    f.write(f"{term:<25} | {count:<15}\n")
                f.write("-" * 45 + "\n")
        except Exception as e:
            print(f"[-] Erro ao salvar o arquivo de relatório: {e}")

    def finalize_footnotes_zip(self, docx_path):
        """Abre o arquivo ZIP .docx e insere/atualiza os arquivos de notas e relações."""
        if not self.footnotes_to_add:
            return

        extract_dir = tempfile.mkdtemp()
        try:
            # 1. Extrai o zip do DOCX
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)

            w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

            # 2. Configura/atualiza word/footnotes.xml
            footnotes_dir = os.path.join(extract_dir, "word")
            if not os.path.exists(footnotes_dir):
                os.makedirs(footnotes_dir)

            footnotes_path = os.path.join(footnotes_dir, "footnotes.xml")
            
            # Se o arquivo já existe, faz o parse. Senão, cria a estrutura base.
            if os.path.exists(footnotes_path):
                tree = etree.parse(footnotes_path)
                root = tree.getroot()
            else:
                # Cria raiz w:footnotes com os namespaces adequados
                root = etree.Element(
                    "{%s}footnotes" % w_ns,
                    nsmap={
                        'w': w_ns,
                        'r': r_ns
                    }
                )
                tree = etree.ElementTree(root)
                
                # Notas estruturais necessárias para separador do Word (IDs -1 e 0)
                # Separador (id=-1)
                fn_sep = etree.SubElement(root, "{%s}footnote" % w_ns)
                fn_sep.set("{%s}type" % w_ns, "separator")
                fn_sep.set("{%s}id" % w_ns, "-1")
                p_sep = etree.SubElement(fn_sep, "{%s}p" % w_ns)
                pPr_sep = etree.SubElement(p_sep, "{%s}pPr" % w_ns)
                spacing_sep = etree.SubElement(pPr_sep, "{%s}spacing" % w_ns)
                spacing_sep.set("{%s}after" % w_ns, "0")
                spacing_sep.set("{%s}line" % w_ns, "240")
                spacing_sep.set("{%s}lineRule" % w_ns, "auto")
                r_sep = etree.SubElement(p_sep, "{%s}r" % w_ns)
                etree.SubElement(r_sep, "{%s}separator" % w_ns)

                # Separador de continuação (id=0)
                fn_cont = etree.SubElement(root, "{%s}footnote" % w_ns)
                fn_cont.set("{%s}type" % w_ns, "continuationSeparator")
                fn_cont.set("{%s}id" % w_ns, "0")
                p_cont = etree.SubElement(fn_cont, "{%s}p" % w_ns)
                pPr_cont = etree.SubElement(p_cont, "{%s}pPr" % w_ns)
                spacing_cont = etree.SubElement(pPr_cont, "{%s}spacing" % w_ns)
                spacing_cont.set("{%s}after" % w_ns, "0")
                spacing_cont.set("{%s}line" % w_ns, "240")
                spacing_cont.set("{%s}lineRule" % w_ns, "auto")
                r_cont = etree.SubElement(p_cont, "{%s}r" % w_ns)
                etree.SubElement(r_cont, "{%s}continuationSeparator" % w_ns)

            # Insere as novas notas de rodapé da lista
            for fn_id, matched_text, fn_text in self.footnotes_to_add:
                # Remove nota existente se houver o mesmo ID (segurança contra duplicados)
                existing = root.find(f".//{{{w_ns}}}footnote[@{{{w_ns}}}id='{fn_id}']")
                if existing is not None:
                    root.remove(existing)

                footnote_node = etree.SubElement(root, "{%s}footnote" % w_ns)
                footnote_node.set("{%s}id" % w_ns, str(fn_id))

                p = etree.SubElement(footnote_node, "{%s}p" % w_ns)
                pPr = etree.SubElement(p, "{%s}pPr" % w_ns)
                pStyle = etree.SubElement(pPr, "{%s}pStyle" % w_ns)
                pStyle.set("{%s}val" % w_ns, "FootnoteText")

                # run 1: Marca da referência da nota de rodapé
                r1 = etree.SubElement(p, "{%s}r" % w_ns)
                rPr1 = etree.SubElement(r1, "{%s}rPr" % w_ns)
                rStyle1 = etree.SubElement(rPr1, "{%s}rStyle" % w_ns)
                rStyle1.set("{%s}val" % w_ns, "FootnoteReference")
                etree.SubElement(r1, "{%s}footnoteRef" % w_ns)

                # run 2: Espaço em branco após a referência
                r2 = etree.SubElement(p, "{%s}r" % w_ns)
                t2 = etree.SubElement(r2, "{%s}t" % w_ns)
                t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t2.text = " "

                # run 3: O termo correspondido (em negrito)
                r_term = etree.SubElement(p, "{%s}r" % w_ns)
                rPr_term = etree.SubElement(r_term, "{%s}rPr" % w_ns)
                etree.SubElement(rPr_term, "{%s}b" % w_ns)
                t_term = etree.SubElement(r_term, "{%s}t" % w_ns)
                t_term.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t_term.text = f"{matched_text}. "

                # run 4: O texto da definição
                r4 = etree.SubElement(p, "{%s}r" % w_ns)
                t4 = etree.SubElement(r4, "{%s}t" % w_ns)
                t4.text = fn_text

            # Salva o footnotes.xml atualizado
            tree.write(footnotes_path, xml_declaration=True, encoding="UTF-8", standalone="yes")

            # 3. Atualiza os relacionamentos em word/_rels/document.xml.rels
            rels_dir = os.path.join(extract_dir, "word", "_rels")
            if not os.path.exists(rels_dir):
                os.makedirs(rels_dir)
            rels_path = os.path.join(rels_dir, "document.xml.rels")

            rel_type_fn = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
            
            if os.path.exists(rels_path):
                rels_tree = etree.parse(rels_path)
                rels_root = rels_tree.getroot()
            else:
                rels_root = etree.Element("Relationships", xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
                rels_tree = etree.ElementTree(rels_root)

            # Verifica se já existe a relação de footnotes
            fn_rel_exists = False
            max_r_id = 0
            for rel in rels_root:
                r_id_str = rel.attrib.get('Id', '')
                if r_id_str.startswith('rId'):
                    try:
                        num = int(r_id_str[3:])
                        if num > max_r_id:
                            max_r_id = num
                    except ValueError:
                        pass
                if rel.attrib.get('Type') == rel_type_fn:
                    fn_rel_exists = True

            if not fn_rel_exists:
                # Adiciona relação usando um novo rId seguro
                new_rid = f"rId{max_r_id + 1}"
                new_rel = etree.SubElement(rels_root, "Relationship")
                new_rel.set("Id", new_rid)
                new_rel.set("Type", rel_type_fn)
                new_rel.set("Target", "footnotes.xml")
                
                # Salva o arquivo de relações atualizado
                rels_tree.write(rels_path, xml_declaration=True, encoding="UTF-8")

            # 4. Atualiza [Content_Types].xml na raiz do ZIP
            ct_path = os.path.join(extract_dir, "[Content_Types].xml")
            if os.path.exists(ct_path):
                ct_tree = etree.parse(ct_path)
                ct_root = ct_tree.getroot()
                
                ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"
                ct_type_fn = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
                
                # Procura se já há o override para /word/footnotes.xml
                fn_ct_exists = False
                for override in ct_root.findall(f".//{{{ns_ct}}}Override"):
                    if override.attrib.get('PartName') == "/word/footnotes.xml":
                        fn_ct_exists = True
                        break
                
                if not fn_ct_exists:
                    new_override = etree.SubElement(ct_root, "{%s}Override" % ns_ct)
                    new_override.set("PartName", "/word/footnotes.xml")
                    new_override.set("ContentType", ct_type_fn)
                    ct_tree.write(ct_path, xml_declaration=True, encoding="UTF-8")

            # 5. Higienização XML (namespaces do Mac e declarações XML)
            self._cleanup_docx_xml(extract_dir)

            # 6. Recompacta o arquivo com ordenação estrita
            self._repack_docx(extract_dir, docx_path)

        finally:
            shutil.rmtree(extract_dir)

    def _cleanup_docx_xml(self, extract_dir):
        """Remove namespaces de Mac e padroniza aspas simples nas declarações XML."""
        for root_dir, dirs, files in os.walk(extract_dir):
            for file in files:
                if file.endswith('.xml') or file.endswith('.rels'):
                    file_path = os.path.join(root_dir, file)
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()

                    modified = False

                    # Remove namespaces indesejados
                    if 'xmlns:mo=' in content or 'xmlns:mv=' in content:
                        content = re.sub(r'\s*xmlns:mo="[^"]*"', '', content)
                        content = re.sub(r'\s*xmlns:mv="[^"]*"', '', content)
                        modified = True

                    # Corrige delimitadores de atributos nas tags xml
                    if "<?xml version='1.0'" in content:
                        content = content.replace(
                            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>",
                            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        )
                        modified = True

                    if modified:
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(content)

    def _repack_docx(self, extract_dir, docx_path):
        """Recompacta a pasta no formato DOCX preservando a ordem exigida dos arquivos XML."""
        all_files = []
        for root_dir, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, extract_dir).replace('\\', '/')
                all_files.append((file_path, arcname))

        # Ordenação de arquivos padrão exigida pelo OpenXML do Microsoft Word
        priority_order = [
            '[Content_Types].xml',
            '_rels/.rels',
            'word/_rels/document.xml.rels',
            'word/document.xml',
            'word/footnotes.xml',
            'word/endnotes.xml',
        ]

        def sort_key(item):
            try:
                return (priority_order.index(item[1]), item[1])
            except ValueError:
                return (len(priority_order), item[1])

        all_files.sort(key=sort_key)

        temp_docx = docx_path + ".tmp"
        with zipfile.ZipFile(temp_docx, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path, arcname in all_files:
                zipf.write(file_path, arcname)

        # Substitui de forma atômica o original pelo temporário criado
        if os.path.exists(docx_path):
            os.remove(docx_path)
        os.rename(temp_docx, docx_path)


def carregar_dicionario(caminho_arquivo):
    """Carrega o dicionário a partir de JSON ou CSV."""
    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo de dicionário não encontrado: {caminho_arquivo}")
        
    ext = os.path.splitext(caminho_arquivo)[1].lower()
    dicionario = {}
    
    if ext == '.json':
        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            dicionario = json.load(f)
    elif ext == '.csv':
        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            # Tenta pular cabeçalho se existirem palavras como "termo,definicao" ou "word,definition"
            for row in reader:
                if len(row) >= 2:
                    termo, definicao = row[0].strip(), row[1].strip()
                    if termo.lower() in ('termo', 'word', 'key') and definicao.lower() in ('definicao', 'definição', 'definition', 'value'):
                        continue
                    dicionario[termo] = definicao
    else:
        raise ValueError("Formato de dicionário não suportado. Use JSON ou CSV.")
        
    return dicionario


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Footnote Injector by Jair Lima - Injeta notas de rodapé no DOCX.")
    parser.add_argument("input_docx", help="Caminho do arquivo DOCX de entrada.")
    parser.add_argument("output_docx", help="Caminho do arquivo DOCX de saída contendo as notas.")
    parser.add_argument("-d", "--dict", help="Caminho para o dicionário JSON ou CSV. Se omitido, usa o dicionario_padrao.json.", default=None)
    parser.add_argument("-a", "--all", action="store_true", help="Se definido, insere nota em todas as ocorrências de um termo (o padrão é apenas na primeira ocorrência).")

    args = parser.parse_args()

    # Define o arquivo de dicionário padrão se não for fornecido
    dict_path = args.dict
    if dict_path is None:
        dict_path = resource_path("dicionario_padrao.json")

    print(f"[*] Carregando dicionário a partir de: {dict_path}")
    try:
        dict_data = carregar_dicionario(dict_path)
        print(f"[+] Dicionário carregado com sucesso! Total de termos: {len(dict_data)}")
    except Exception as e:
        print(f"[-] Erro ao carregar o dicionário: {e}")
        return

    print(f"[*] Iniciando injeção em: {args.input_docx}")
    print(f"[*] Somente primeira ocorrência: {not args.all}")
    
    try:
        injector = FootnoteInjector(dict_data)
        injector.inject(args.input_docx, args.output_docx, only_first_occurrence=not args.all)
        print(f"[+] Sucesso! Arquivo gerado em: {args.output_docx}")
        print(f"[+] Total de notas de rodapé adicionadas: {injector.next_footnote_id - 1}")
        
        # Gera o relatório de injeção na pasta de saída
        report_path = os.path.splitext(args.output_docx)[0] + "_relatorio.txt"
        injector.gerar_relatorio(report_path)
        print(f"[+] Relatório detalhado salvo em: {report_path}")
    except Exception as e:
        print(f"[-] Erro durante o processo de injeção: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
