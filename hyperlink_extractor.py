# hyperlink_extractor.py

import logging
from urllib.parse import urlparse
import re
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.shared import qn

logging.basicConfig(filename='hyperlink_extractor.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def extract_hyperlinks(document):
    hyperlinks = []
    logging.info("Kezdődik a hivatkozások kinyerése")

    # 1. Beágyazott hiperhivatkozások
    logging.info("1. módszer: Beágyazott hiperhivatkozások keresése")
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            try:
                # FONTOS: Kezeljük a 'CT_R' object has no attribute 'hyperlink' hibát
                if hasattr(run.element, 'hyperlink'):
                    if run.element.hyperlink is not None:
                        target = run.element.hyperlink.get(qn('r:id'))
                        if target:
                            target = document.part.rels[target].target_ref
                            hyperlinks.append({"text": run.text, "target": target, "type": determine_hyperlink_type(target)})
                            logging.debug(f"Talált beágyazott hivatkozás: {run.text} -> {target}")
                else:
                    # Alternatív módszer a hivatkozás kinyerésére
                    for elem in run._element.xpath('.//w:hyperlink'):
                        rId = elem.get(qn('r:id'))
                        if rId:
                            target = document.part.rels[rId].target_ref
                            hyperlinks.append({"text": run.text, "target": target, "type": determine_hyperlink_type(target)})
                            logging.debug(f"Talált beágyazott hivatkozás (alternatív módszer): {run.text} -> {target}")
            except AttributeError as e:
                logging.warning(f"AttributeError a hivatkozás kinyerése során: {str(e)}")

    # 2. Szöveges URL-ek
    logging.info("2. módszer: Szöveges URL-ek keresése")
    url_pattern = re.compile(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+')
    for paragraph in document.paragraphs:
        urls = url_pattern.findall(paragraph.text)
        for url in urls:
            hyperlinks.append({"text": url, "target": url, "type": determine_hyperlink_type(url)})
            logging.debug(f"Talált szöveges URL: {url}")

    # 3. Mezők (fields)
    logging.info("3. módszer: Mezők keresése")
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if run._element.get('fldChar') is not None:
                field_text = get_field_text(run._element)
                if field_text.startswith('HYPERLINK'):
                    parts = field_text.split('"')
                    if len(parts) > 1:
                        url = parts[1]
                        hyperlinks.append({"text": "Mező hivatkozás", "target": url, "type": determine_hyperlink_type(url)})
                        logging.debug(f"Talált mező hivatkozás: {url}")

    # 4. Könyvjelzők
    logging.info("4. módszer: Könyvjelzők keresése")
    for bookmark in document.element.xpath('//w:bookmarkStart'):
        name = bookmark.get(qn('w:name'))
        if name.startswith('_'):
            continue  # Belső könyvjelzők kihagyása
        hyperlinks.append({"text": f"Könyvjelző: {name}", "target": f"#{name}", "type": "belső"})
        logging.debug(f"Talált könyvjelző: {name}")

    # 5. Speciális belső hivatkozások
    logging.info("5. módszer: Speciális belső hivatkozások keresése")
    internal_link_pattern = re.compile(r'(.*?)\t.*?BKM_[A-F0-9_]+')
    for paragraph in document.paragraphs:
        matches = internal_link_pattern.findall(paragraph.text)
        for match in matches:
            text = match.strip()
            target = paragraph.text.split('\t')[1].strip()
            link_type = determine_internal_link_type(target)
            hyperlinks.append({"text": text, "target": target, "type": link_type})
            logging.debug(f"Talált speciális belső hivatkozás: {text} -> {target} ({link_type})")

    unique_links = remove_duplicates(hyperlinks)
    logging.info(f"Összesen {len(unique_links)} egyedi hivatkozás található")
    return unique_links

def get_field_text(element):
    field_text = ''
    for e in element.getparent().iter():
        if e.tag.endswith('fldChar') and e.get('fldCharType') == 'end':
            break
        if e.tag.endswith('instrText'):
            field_text += e.text
    return field_text

def determine_hyperlink_type(url):
    if url.startswith("#") or "BKM_" in url:
        return "belső"
    elif is_valid_url(url):
        return "külső"
    else:
        return "törött"

def determine_internal_link_type(target):
    if "BKM_" in target:
        if "részleges egyezés" in target.lower():
            return "Érvényes belső hivatkozás (részleges egyezés)"
        else:
            return "Valószínűleg törött belső hivatkozás (szellem-hivatkozás)"
    return "Ismeretlen belső hivatkozás típus"

def is_valid_url(url):
    try:
        result = urlparse(url)
        return all([result.scheme, result.netloc])
    except:
        return False

def remove_duplicates(hyperlinks):
    seen = set()
    unique_hyperlinks = []
    for link in hyperlinks:
        key = (link['text'], link['target'])
        if key not in seen:
            seen.add(key)
            unique_hyperlinks.append(link)
    return unique_hyperlinks
