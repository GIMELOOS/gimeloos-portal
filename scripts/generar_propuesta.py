#!/usr/bin/env python3
"""
Genera un documento Word de propuesta de animación a partir de la plantilla.
Uso: python3 generar_propuesta.py '<json_payload>' <output_path>
"""
import os
import sys
import json
import copy
import re
import zipfile
from pathlib import Path
from lxml import etree

_default_template = Path(__file__).parent / "plantilla.docx"
TEMPLATE = Path(os.environ.get("PROPUESTA_TEMPLATE_PATH", str(_default_template)))

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{NS}}}"

FONT_NAME = "Aptos"
FONT_SZ = "36"   # 18pt = 36 half-points


INTROS = {
    "busqueda del tesoro": (
        "Haremos una búsqueda del tesoro. El cofre secreto ha desaparecido. Deberán superar cada "
        "prueba para conseguir descubrir dónde está. Por el lugar, de manera aleatoria, estarán "
        "distribuidos diferentes sobres con actividades. En la medida en que vayan encontrando los "
        "sobres, realizarán una actividad u otra. El orden de los sobres será aleatorio, no seguirán "
        "este orden necesariamente. Tras cada prueba se les darán pistas de dónde está el cofre "
        "secreto representado con pasos que deben dar. Al acabar todas las pruebas, tendrán todas las "
        "pistas necesarias para poder encontrar el cofre secreto."
    ),
    "gymkana": (
        "Organizaremos una gymkana que consiste en una batería de juegos de diferente estilo. "
        "Divididos en 2 grupos y liderados por un capitán, los participantes deberán superar una serie "
        "de pruebas para sumar puntos y hacerse con la victoria en esta gymkana que combina diferentes "
        "habilidades. Al acabar todas las pruebas, sumaremos los puntos y nombraremos al equipo ganador."
    ),
    "furor": (
        "Se dividirá a los participantes en equipos. En cada equipo habrá un monitor GIMELOOS. Se "
        "plantearán una serie de pruebas que se alternarán con retos para ganar puntos extra u obtener "
        "beneficios para otras pruebas. Al acabar cada prueba recontaremos los puntos y los sumaremos. "
        "Ganará el equipo que más puntos tenga."
    ),
    "feria": (
        "Cada participante recibirá X billetes. Tendrán que pasar por la caseta de tickets a recoger "
        "los billetes. Con los billetes, deberán pasar por los puestos y probar suerte. Cada vez que "
        "participen en un puesto, deberán dar X billetes de entrada. Si ganan reciben los billetes de "
        "entrada duplicados. Si pierden, no recuperan los billetes de entrada. Previo a la finalización "
        "de la actividad se abrirá la tómbola/cantina para que puedan canjear los billetes."
    ),
}


def get_intro(tematica):
    """Devuelve el texto de intro para la temática, o None si no hay."""
    if not tematica:
        return None
    key = tematica.lower().strip()
    # Coincidencia exacta
    if key in INTROS:
        return INTROS[key]
    # Coincidencia parcial (la temática contiene la clave)
    for k, v in INTROS.items():
        if k in key:
            return v
    return None


def limpiar_tematica(tematica):
    return re.sub(r'\s*\(.*?\)\s*', '', tematica or '').strip()


def add_page_border(xml, color="FF5757", sz=128):
    """
    Añade w:pgBorders a todos los w:sectPr del documento.
    sz=128 → 16pt = 320 twips; debe coincidir con BORDER_TWIPS en main().
    """
    border_xml = (
        f'<w:pgBorders w:offsetFrom="page">'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:pgBorders>'
    )

    def inject_into_sectpr(m):
        s = m.group(0)
        if 'pgBorders' in s:
            return s
        # Insertar justo después de la etiqueta de apertura <w:sectPr ...>
        gt = s.index('>') + 1
        return s[:gt] + border_xml + s[gt:]

    return re.sub(r'<w:sectPr\b[^>]*>.*?</w:sectPr>', inject_into_sectpr, xml, flags=re.DOTALL)


def _make_rpr(bold=False):
    """Crea un elemento w:rPr con fuente Aptos y tamaño 18pt."""
    rpr = etree.SubElement(etree.Element(f'{W}rPr'), f'{W}rPr')
    rpr = etree.Element(f'{W}rPr')
    if bold:
        etree.SubElement(rpr, f'{W}b')
    rfonts = etree.SubElement(rpr, f'{W}rFonts')
    rfonts.set(f'{W}ascii', FONT_NAME)
    rfonts.set(f'{W}hAnsi', FONT_NAME)
    rfonts.set(f'{W}cs', FONT_NAME)
    sz = etree.SubElement(rpr, f'{W}sz')
    sz.set(f'{W}val', FONT_SZ)
    szcs = etree.SubElement(rpr, f'{W}szCs')
    szcs.set(f'{W}val', FONT_SZ)
    return rpr


def _make_run(text, bold=False):
    """Crea un w:r con el texto dado, en Aptos 18pt."""
    r = etree.Element(f'{W}r')
    r.append(_make_rpr(bold=bold))
    t = etree.SubElement(r, f'{W}t')
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def _get_cell_text(tc):
    """Extrae el texto visible de una celda, ignorando runs con drawings."""
    parts = []
    for p in tc.findall(f'.//{W}p'):
        for r in p.findall(f'{W}r'):
            # Ignorar runs que contienen drawings (el recuadro rojo)
            if r.find(f'{W}drawing') is not None:
                continue
            for t in r.findall(f'{W}t'):
                parts.append(t.text or '')
    return ''.join(parts).strip()


def _get_ppr(tc):
    """Obtiene el w:pPr del primer párrafo de la celda (para reutilizar formato)."""
    first_p = tc.find(f'{W}p')
    if first_p is not None:
        return first_p.find(f'{W}pPr')
    return None


def _clear_cell_text_runs(tc):
    """
    Elimina todos los runs (texto y drawings) del primer párrafo de la celda
    y todos los párrafos adicionales. El borde de página lo gestiona w:pgBorders,
    por lo que las formas antiguas del template ya no son necesarias.
    """
    ns_mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
    first_p = tc.find(f'{W}p')
    if first_p is None:
        return

    for r in list(first_p.findall(f'{W}r')):
        first_p.remove(r)
    for ac in list(first_p.findall(f'{{{ns_mc}}}AlternateContent')):
        first_p.remove(ac)

    for p in tc.findall(f'{W}p')[1:]:
        tc.remove(p)



def set_cell_text_xml(tc, text):
    """Escribe texto en la celda preservando el drawing si lo hubiera."""
    _clear_cell_text_runs(tc)

    first_p = tc.find(f'{W}p')
    if first_p is None:
        first_p = etree.SubElement(tc, f'{W}p')

    first_pPr = first_p.find(f'{W}pPr')
    _fix_ppr_indent(first_pPr)
    first_p.append(_make_run(text))


def _fix_ppr_indent(ppr, left="400", right="400"):
    """
    Garantiza márgenes mínimos en w:ind para que el texto no quede
    tapado por el borde rojo (320 twips). El valor por defecto (400)
    supera el grosor del borde y da un margen visual cómodo.
    """
    if ppr is None:
        return
    ind = ppr.find(f'{W}ind')
    if ind is not None:
        if not ind.get(f'{W}left') or int(ind.get(f'{W}left', 0)) < int(left):
            ind.set(f'{W}left', left)
        if not ind.get(f'{W}right') or int(ind.get(f'{W}right', 0)) < int(right):
            ind.set(f'{W}right', right)
    else:
        ind = etree.Element(f'{W}ind')
        ind.set(f'{W}left', left)
        ind.set(f'{W}right', right)
        spacing = ppr.find(f'{W}spacing')
        if spacing is not None:
            spacing.addnext(ind)
        else:
            ppr.insert(0, ind)


def _make_game_ppr(base_ppr):
    """Crea pPr para párrafo de juego heredando el formato base."""
    ppr = copy.deepcopy(base_ppr) if base_ppr is not None else etree.Element(f'{W}pPr')
    _fix_ppr_indent(ppr)
    return ppr


def add_description_xml(tc, juegos, tematica=None):
    """
    Escribe la intro de temática (si existe) seguida de la lista numerada de juegos.
    Preserva el run con el drawing (recuadro rojo) en el primer párrafo.
    """
    base_ppr = copy.deepcopy(_get_ppr(tc))
    _clear_cell_text_runs(tc)

    first_p = tc.find(f'{W}p')
    if first_p is None:
        first_p = etree.SubElement(tc, f'{W}p')

    # Arreglar el indent del primer párrafo (puede no tener w:left)
    first_pPr = first_p.find(f'{W}pPr')
    _fix_ppr_indent(first_pPr)

    intro = get_intro(tematica)
    parrafo_inicial_usado = False

    # Párrafo de intro (si hay)
    if intro:
        first_p.append(_make_run(intro))
        parrafo_inicial_usado = True

    # Lista numerada de juegos (cada uno con keepLines para no partirse)
    for i, j in enumerate(juegos, 1):
        nombre = j.get('nombre', '')
        desc = j.get('descripcion', '') or ''
        if desc:
            desc = desc[0].upper() + desc[1:]

        if not parrafo_inicial_usado:
            # Usar el primer párrafo (que puede tener el drawing)
            first_p.append(_make_run(f"{i}. {nombre}: ", bold=True))
            first_p.append(_make_run(desc))
            parrafo_inicial_usado = True
        else:
            new_p = etree.SubElement(tc, f'{W}p')
            new_p.append(_make_game_ppr(base_ppr))
            new_p.append(_make_run(f"{i}. {nombre}: ", bold=True))
            new_p.append(_make_run(desc))


def main():
    if len(sys.argv) < 3:
        print("Uso: generar_propuesta.py '<json>' <output_path>", file=sys.stderr)
        sys.exit(1)

    data_arg = json.loads(sys.argv[1])
    output_path = Path(sys.argv[2])
    output_path.parent.mkdir(parents=True, exist_ok=True)

    reserva = data_arg['reserva']
    juegos = data_arg['juegos']

    festejado = reserva.get('festejado', {})
    evento = reserva.get('evento', {})
    nombre_completo = ' '.join(filter(None, [
        festejado.get('nombre', ''), festejado.get('apellidos', '')
    ])).strip() or 'Sin nombre'

    tipo_actividad = limpiar_tematica(evento.get('tematica', ''))
    monitores = (str(reserva.get('monitoresEstimados', '')) + ' monitores') if reserva.get('monitoresEstimados') else ''
    participantes = (str(reserva.get('participantes', '')) + ' niños') if reserva.get('participantes') else ''
    fecha_horario = f"{evento.get('fecha', '')} · {evento.get('horario', '')}".strip(' ·')
    if evento.get('horasDuracion'):
        fecha_horario += f" ({evento['horasDuracion']}h)"
    materiales = sorted(set(m for j in juegos for m in j.get('materiales', [])))
    materiales_texto = ', '.join(materiales) if materiales else ''

    # Leer la plantilla como zip
    with zipfile.ZipFile(TEMPLATE, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    xml_doc = files['word/document.xml'].decode('utf-8')
    # Reemplazar el título en los cuadros de texto
    xml_doc = xml_doc.replace('TÍTULO', nombre_completo)

    # Añadir borde rojo de página nativo de Word
    xml_doc = add_page_border(xml_doc)

    # Parsear con lxml para modificar la tabla
    # Registrar todos los namespaces para preservarlos en la serialización
    root = etree.fromstring(xml_doc.encode('utf-8'))

    # Encontrar la primera tabla
    tables = root.findall(f'.//{W}tbl')
    if not tables:
        print("Error: no se encontró ninguna tabla en la plantilla", file=sys.stderr)
        sys.exit(1)

    table = tables[0]

    # Convertir la tabla de flotante a inline eliminando w:tblpPr.
    # Las tablas flotantes no pueden crecer más allá de una página;
    # hacerla inline permite que el contenido fluya a páginas siguientes.
    tblPr = table.find(f'{W}tblPr')
    if tblPr is not None:
        tblpPr = tblPr.find(f'{W}tblpPr')
        if tblpPr is not None:
            tblPr.remove(tblpPr)
        # Centrar la tabla horizontalmente en la página
        jc = tblPr.find(f'{W}jc')
        if jc is None:
            jc = etree.SubElement(tblPr, f'{W}jc')
        jc.set(f'{W}val', 'center')

    # La portada y la tabla están en la misma sección sin salto de página.
    # Creamos un salto de sección entre portada y tabla con márgenes distintos:
    #   Sección 1 (portada): márgenes 1701 — formas posicionadas correctamente
    #   Sección 2 (tabla):   márgenes 500  — tabla ancha con margen visible
    TABLE_MARGIN = 0   # tabla ancho completo — el borde rojo lo pone w:pgBorders
    TABLE_W = 11906    # ancho de página A4 en twips

    body = root.find(f'{W}body')
    if body is not None:
        final_sectPr = body.find(f'{W}sectPr')  # hijo directo de body = sección final
        if final_sectPr is not None:
            # Clonar sectPr para la portada (conserva márgenes 1701 originales)
            cover_sectPr = copy.deepcopy(final_sectPr)

            # BORDER_TWIPS debe coincidir con sz en add_page_border():
            # sz=128 → 128/8 = 16pt → 16×20 = 320 twips
            # Los márgenes top/bottom iguales al grosor del borde eliminan la
            # línea blanca entre el borde rojo y la tabla en cada página.
            BORDER_TWIPS = 320  # 16pt (sz=128 en pgBorders)
            pgMar = final_sectPr.find(f'{W}pgMar')
            if pgMar is not None:
                pgMar.set(f'{W}left', str(TABLE_MARGIN))
                pgMar.set(f'{W}right', str(TABLE_MARGIN))
                pgMar.set(f'{W}top', str(BORDER_TWIPS))
                pgMar.set(f'{W}bottom', str(BORDER_TWIPS))

            # Ajustar ancho de la tabla y todas sus celdas al nuevo TABLE_W
            tblPr_elem = table.find(f'{W}tblPr')
            if tblPr_elem is not None:
                tblW_elem = tblPr_elem.find(f'{W}tblW')
                if tblW_elem is not None:
                    tblW_elem.set(f'{W}w', str(TABLE_W))
            for tc_elem in table.findall(f'.//{W}tc'):
                tcPr_elem = tc_elem.find(f'{W}tcPr')
                if tcPr_elem is not None:
                    tcW_elem = tcPr_elem.find(f'{W}tcW')
                    if tcW_elem is not None:
                        tcW_elem.set(f'{W}w', str(TABLE_W))

            # Párrafo de salto de sección antes de la tabla.
            # Se le da altura cero (line=1 exact, sz=2) para que no sangre
            # en la página 2 creando la línea blanca delgada.
            sec_break_p = etree.Element(f'{W}p')
            sec_break_pPr = etree.SubElement(sec_break_p, f'{W}pPr')
            spacing_sb = etree.SubElement(sec_break_pPr, f'{W}spacing')
            spacing_sb.set(f'{W}before', '0')
            spacing_sb.set(f'{W}after', '0')
            spacing_sb.set(f'{W}line', '1')
            spacing_sb.set(f'{W}lineRule', 'exact')
            rPr_sb = etree.SubElement(sec_break_pPr, f'{W}rPr')
            sz_sb = etree.SubElement(rPr_sb, f'{W}sz')
            sz_sb.set(f'{W}val', '2')  # marca de párrafo de 1pt = altura cero efectiva
            sec_break_pPr.append(cover_sectPr)

            tbl_idx = next(
                (i for i, c in enumerate(body) if c.tag == f'{W}tbl'), None
            )
            if tbl_idx is not None:
                body.insert(tbl_idx, sec_break_p)

    # Eliminar la altura fija de la fila de descripción (la más alta)
    # para que pueda crecer según el contenido.
    for tr in table.findall(f'{W}tr'):
        trPr = tr.find(f'{W}trPr')
        if trPr is not None:
            trHeight = trPr.find(f'{W}trHeight')
            if trHeight is not None:
                val = int(trHeight.get(f'{W}val', 0))
                if val >= 3000:  # solo la fila de descripción (grande)
                    trPr.remove(trHeight)

    # Obtener todas las celdas en orden
    cells = table.findall(f'.//{W}tc')

    VALORES = {
        'TIPO DE ACTIVIDAD': tipo_actividad,
        'MONITORES': monitores,
        'PARTICIPANTES': participantes,
        'FECHA Y HORARIO': fecha_horario,
        'MATERIALES': materiales_texto,
    }

    for i, cell in enumerate(cells):
        label = _get_cell_text(cell)
        if label in VALORES and i + 1 < len(cells):
            set_cell_text_xml(cells[i + 1], VALORES[label])
        elif 'DESCRIPCI' in label and 'ACTIVIDAD' in label and i + 1 < len(cells):
            add_description_xml(cells[i + 1], juegos, tematica=tipo_actividad)

    # Serializar el XML modificado preservando las declaraciones de namespaces
    modified_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    files['word/document.xml'] = modified_xml

    with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name, file_data in files.items():
            zout.writestr(name, file_data)

    print(f"OK:{output_path}")


if __name__ == '__main__':
    main()
