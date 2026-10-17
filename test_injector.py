#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script de testes automatizados para o Footnote Injector by Jair Lima.
Gera um documento Word de teste com termos e formatação variada, executa o
injetor e realiza asserções de integridade estrutural e de formatação no XML do DOCX.
"""

import os
import shutil
import unittest
import zipfile
from lxml import etree
from docx import Document

from docx_footnote_injector import FootnoteInjector


class TestFootnoteInjector(unittest.TestCase):

    def setUp(self):
        self.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_temp")
        if not os.path.exists(self.test_dir):
            os.makedirs(self.test_dir)

        self.input_docx = os.path.join(self.test_dir, "input_teste.docx")
        self.output_docx = os.path.join(self.test_dir, "output_teste.docx")

        # Dicionário de testes
        self.test_dict = {
            "Docker": "Tecnologia de virtualização baseada em contêineres.",
            "Kubernetes": "Orquestrador de contêineres de código aberto.",
            "API": "Interface de programação de aplicações."
        }

        # Cria um arquivo Word de teste do zero usando python-docx
        doc = Document()
        
        # Parágrafo 1: Termo simples sem formatação
        p1 = doc.add_paragraph("Este é um parágrafo de exemplo que menciona o Docker como ferramenta.")
        
        # Parágrafo 2: Termo no meio de texto formatado (negrito e itálico)
        p2 = doc.add_paragraph()
        p2.add_run("O orquestrador ")
        r_bold = p2.add_run("Kubernetes")
        r_bold.bold = True
        r_bold.italic = True
        p2.add_run(" é amplamente utilizado em ambientes de produção.")

        # Parágrafo 3: Múltiplos termos no mesmo parágrafo
        p3 = doc.add_paragraph("Uma API conecta o Docker com outros sistemas externos.")

        # Parágrafo 4: Termo como substring de outra palavra (não deve disparar nota)
        p4 = doc.add_paragraph("O dockerhub ou o importador de API não devem disparar notas de rodapé de forma parcial.")

        # Parágrafo 5: Título que contém o termo (deve ser ignorado)
        p5 = doc.add_paragraph("Título sobre Kubernetes", style='Heading 1')

        doc.save(self.input_docx)

    def tearDown(self):
        # Remove diretório temporário após os testes
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)

    def test_footnote_injection(self):
        """Valida que o injetor insere as notas e preserva a formatação XML."""
        injector = FootnoteInjector(self.test_dict)
        
        # Injeta as notas de rodapé
        injector.inject(self.input_docx, self.output_docx, only_first_occurrence=True)

        self.assertTrue(os.path.exists(self.output_docx), "O arquivo de saída não foi gerado.")

        # Realiza validações abrindo o arquivo ZIP do DOCX
        with zipfile.ZipFile(self.output_docx, 'r') as zip_ref:
            # 1. Verifica a existência de word/footnotes.xml
            self.assertIn("word/footnotes.xml", zip_ref.namelist(), "O arquivo word/footnotes.xml está ausente no ZIP.")

            # 2. Verifica a existência da relação no document.xml.rels
            self.assertIn("word/_rels/document.xml.rels", zip_ref.namelist(), "O arquivo de relacionamentos está ausente.")
            
            rels_data = zip_ref.read("word/_rels/document.xml.rels")
            rels_root = etree.fromstring(rels_data)
            
            rel_type_fn = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
            fn_relation = [rel for rel in rels_root if rel.attrib.get('Type') == rel_type_fn]
            self.assertEqual(len(fn_relation), 1, "Deve existir exatamente uma relação de notas de rodapé.")
            self.assertEqual(fn_relation[0].attrib.get('Target'), "footnotes.xml", "O alvo do relacionamento deve ser footnotes.xml.")

            # 3. Verifica o arquivo [Content_Types].xml
            ct_data = zip_ref.read("[Content_Types].xml")
            ct_root = etree.fromstring(ct_data)
            ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"
            
            override_fn = [
                ov for ov in ct_root.findall(f".//{{{ns_ct}}}Override")
                if ov.attrib.get('PartName') == "/word/footnotes.xml"
            ]
            self.assertEqual(len(override_fn), 1, "O tipo de conteúdo para word/footnotes.xml deve estar registrado no [Content_Types].xml.")
            self.assertEqual(
                override_fn[0].attrib.get('ContentType'),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
            )

            # 4. Verifica o arquivo word/footnotes.xml para as definições inseridas
            fn_data = zip_ref.read("word/footnotes.xml")
            fn_root = etree.fromstring(fn_data)
            w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            
            # Devem existir 3 notas customizadas (Docker, Kubernetes e API) + as 2 estruturais (-1 e 0)
            all_footnotes = fn_root.findall(f".//{{{w_ns}}}footnote")
            self.assertEqual(len(all_footnotes), 5, "Devem existir 5 notas de rodapé no XML (2 estruturais + 3 injetadas).")

            # Valida os textos das definições injetadas
            definitions = []
            for footnote in all_footnotes:
                fn_id = footnote.attrib.get(f"{{{w_ns}}}id")
                if fn_id not in ("-1", "0"):
                    texts = [t.text for t in footnote.findall(f".//{{{w_ns}}}t") if t.text]
                    definitions.append("".join(texts))

            self.assertIn(" Docker. Tecnologia de virtualização baseada em contêineres.", definitions)
            self.assertIn(" Kubernetes. Orquestrador de contêineres de código aberto.", definitions)
            self.assertIn(" API. Interface de programação de aplicações.", definitions)

        # 5. Abre o documento final no python-docx para validar que os estilos e a integridade foram mantidos
        doc = Document(self.output_docx)
        
        # O parágrafo 2 (Kubernetes) deve ter o run "Kubernetes" em negrito e itálico
        p2 = doc.paragraphs[1]
        
        # Encontra o run correspondente ao termo
        term_run = None
        for run in p2.runs:
            if run.text == "Kubernetes":
                term_run = run
                break

        self.assertIsNotNone(term_run, "O run do termo Kubernetes não foi localizado após o processamento.")
        self.assertTrue(term_run.bold, "O run Kubernetes deve manter o estilo negrito.")
        self.assertTrue(term_run.italic, "O run Kubernetes deve manter o estilo itálico.")

        # Valida que o parágrafo 4 não disparou notas para substrings
        p4 = doc.paragraphs[3]
        # O texto de p4 não deve possuir tags de footnoteReference no XML subjacente
        p4_xml = etree.tostring(p4._p).decode('utf-8')
        self.assertNotIn("footnoteReference", p4_xml, "A substring dockerhub ou API parciais não devem gerar notas de rodapé.")

        # Valida que o parágrafo 5 (Título) ignorou a injeção do termo
        p5 = doc.paragraphs[4]
        p5_xml = etree.tostring(p5._p).decode('utf-8')
        self.assertNotIn("footnoteReference", p5_xml, "Títulos ou cabeçalhos do tipo Heading não devem receber notas de rodapé.")

        # Valida as contagens corretas no relatório de inserção de termos
        self.assertEqual(injector.insertion_report.get("Docker"), 1)
        self.assertEqual(injector.insertion_report.get("Kubernetes"), 1)
        self.assertEqual(injector.insertion_report.get("API"), 1)


if __name__ == "__main__":
    unittest.main()
