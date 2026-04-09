"""
actualizar_pagina.py
====================
Lee "detalle y stock.xlsx" y actualiza index.html automaticamente.

Uso:
    python actualizar_pagina.py

Opciones:
    python actualizar_pagina.py --excel ruta/al/archivo.xlsx
    python actualizar_pagina.py --preview   (muestra cambios sin aplicar)
"""

import openpyxl
import re
import os
import sys
import shutil
from datetime import datetime

# --- Configuracion ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(os.path.expanduser('~'), 'Downloads', 'detalle y stock.xlsx')
HTML_PATH = os.path.join(SCRIPT_DIR, 'index.html')
BACKUP_DIR = os.path.join(SCRIPT_DIR, 'backups')

# Titulos de seccion para el HTML
SECTION_TITLES = {
    'sahumerios': 'Sahumerios',
    'bombas': 'Bombas y Defumacion',
    'sahumadores': 'Sahumadores',
    'hornitos': 'Hornitos',
    'portasahumerios': 'Portasahumerios',
    'aromatizadores': 'Aromatizadores',
    'kits': 'Kits',
    'llaveros': 'Llaveros',
    'velas': 'Velas',
    'difusores': 'Difusores',
    'lamparas': 'Lamparas',
    'sales': 'Sales',
    'garrapinada': 'Garrapinada',
}

# Orden de secciones en la pagina
SECTION_ORDER = list(SECTION_TITLES.keys())


def read_excel(path):
    """Lee el Excel y devuelve lista de productos.

    Columnas del Excel:
        A = Articulo (nombre)
        B = Descripcion
        C = Precio
        D = Cantidad (stock)
        E = Seccion (categoria)
        F = Marca
        G = Imagen
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    products = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 7:
            continue

        nombre, descripcion, precio, stock, seccion, marca, imagen = row[:7]

        # Saltar filas vacias
        if not nombre or not seccion:
            continue

        seccion = str(seccion).strip().lower()
        nombre = str(nombre).strip()
        descripcion = str(descripcion).strip() if descripcion else ''
        precio = int(precio) if precio else 0
        stock = int(stock) if stock else 0
        marca = str(marca).strip() if marca else 'Artesanal'
        imagen = str(imagen).strip() if imagen else ''

        products.append({
            'seccion': seccion,
            'nombre': nombre,
            'descripcion': descripcion,
            'precio': precio,
            'stock': stock,
            'marca': marca,
            'imagen': imagen,
        })

    return products


def generate_card_html(p):
    """Genera el HTML de una card de producto."""
    price_display = f'${p["precio"]:,.0f}'.replace(',', '.') if p['precio'] > 0 else 'Consultar'
    return (
        f'      <div class="card" data-name="{p["nombre"]}" data-price="{p["precio"]}" '
        f'data-brand="{p["marca"]}" data-stock="{p["stock"]}" '
        f'data-img="{p["imagen"]}" data-desc="{p["descripcion"]}">\n'
        f'        <div class="card-img-wrap"><img src="{p["imagen"]}" alt="{p["nombre"]}" loading="lazy" /></div>\n'
        f'        <div class="card-body">\n'
        f'          <div class="card-category">{SECTION_TITLES.get(p["seccion"], p["seccion"].title())}</div>\n'
        f'          <div class="card-name">{p["nombre"]}</div>\n'
        f'          <div class="card-footer">\n'
        f'            <div class="card-price">{price_display}</div>\n'
        f'            <button class="add-btn" onclick="addToCart(this)">+</button>\n'
        f'          </div>\n'
        f'        </div>\n'
        f'      </div>\n'
    )


def generate_section_html(seccion, products):
    """Genera el HTML completo de una seccion."""
    title = SECTION_TITLES.get(seccion, seccion.title())
    cards = ''.join(generate_card_html(p) for p in products)
    return (
        f'\n  <!-- ===== {title.upper()} ===== -->\n'
        f'  <div class="product-section" data-cat="{seccion}">\n'
        f'    <h2 class="section-title">{title} <span></span></h2>\n'
        f'    <div class="grid">\n\n'
        f'{cards}\n'
        f'    </div>\n'
        f'  </div>\n'
    )


def generate_brand_checkboxes(products):
    """Genera los checkboxes de marca para el sidebar."""
    brands = {}
    for p in products:
        marca = p['marca']
        brands[marca] = brands.get(marca, 0) + 1

    # Orden: Sagrada Madre primero, luego alfabetico, Artesanal al final
    def brand_sort_key(item):
        name = item[0]
        if name == 'Sagrada Madre':
            return (0, name)
        if name == 'Artesanal':
            return (2, name)
        return (1, name)

    sorted_brands = sorted(brands.items(), key=brand_sort_key)

    lines = []
    for marca, count in sorted_brands:
        lines.append(
            f'      <label class="filter-option">'
            f'<input type="checkbox" value="{marca}" onchange="applyFilters()"> '
            f'{marca} <span class="count" data-brand-count="{marca}"></span></label>'
        )
    return '\n'.join(lines)


def generate_category_checkboxes(products):
    """Genera los checkboxes de categoria para el sidebar."""
    cats_seen = []
    for p in products:
        if p['seccion'] not in cats_seen:
            cats_seen.append(p['seccion'])

    lines = []
    for cat in cats_seen:
        label = SECTION_TITLES.get(cat, cat.title())
        lines.append(
            f'      <label class="filter-option">'
            f'<input type="checkbox" value="{cat}" onchange="applyFilters()"> '
            f'{label} <span class="count" data-cat-count="{cat}"></span></label>'
        )
    return '\n'.join(lines)


def update_html(html, products, preview=False):
    """Actualiza el HTML con los productos del Excel."""

    # 1. Reemplazar secciones de productos
    # Encontrar desde <main class="shop"> hasta </main>
    main_match = re.search(
        r'(<main class="shop">)\s*(.*?)\s*(</main>)',
        html, re.DOTALL
    )
    if not main_match:
        print('ERROR: No se encontro <main class="shop"> en el HTML')
        sys.exit(1)

    # Agrupar productos por seccion manteniendo el orden
    sections = {}
    for p in products:
        sec = p['seccion']
        if sec not in sections:
            sections[sec] = []
        sections[sec].append(p)

    # Generar HTML de todas las secciones
    all_sections_html = ''
    # Usar el orden definido, mas cualquier seccion nueva
    ordered_sections = [s for s in SECTION_ORDER if s in sections]
    new_sections = [s for s in sections if s not in SECTION_ORDER]
    for seccion in ordered_sections + new_sections:
        all_sections_html += generate_section_html(seccion, sections[seccion])

    # Agregar el div de no-results
    all_sections_html += (
        '\n  <div class="no-results" id="noResults">\n'
        '    <div class="no-results-icon">&#128269;</div>\n'
        '    <p>No encontramos productos con ese nombre.<br/>'
        'Proba con otra palabra o <strong>escribinos por WhatsApp</strong> para ayudarte.</p>\n'
        '  </div>\n'
    )

    new_main = f'{main_match.group(1)}\n{all_sections_html}\n{main_match.group(3)}'
    html = html[:main_match.start()] + new_main + html[main_match.end():]

    # 2. Actualizar checkboxes de categorias en sidebar
    cat_match = re.search(
        r'(<div class="sidebar-body" id="filterCats">)\s*(.*?)\s*(</div>)',
        html, re.DOTALL
    )
    if cat_match:
        new_cats = generate_category_checkboxes(products)
        html = (html[:cat_match.start(1)] +
                cat_match.group(1) + '\n' + new_cats + '\n    ' + cat_match.group(3) +
                html[cat_match.end(3):])

    # 3. Actualizar checkboxes de marcas en sidebar
    brand_match = re.search(
        r'(<div class="sidebar-body" id="filterBrands">)\s*(.*?)\s*(</div>)',
        html, re.DOTALL
    )
    if brand_match:
        new_brands = generate_brand_checkboxes(products)
        html = (html[:brand_match.start(1)] +
                brand_match.group(1) + '\n' + new_brands + '\n    ' + brand_match.group(3) +
                html[brand_match.end(3):])

    return html


def create_backup(html_path):
    """Crea un backup del HTML antes de modificar."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'index_backup_{timestamp}.html')
    shutil.copy2(html_path, backup_path)
    return backup_path


def main():
    # Parsear argumentos
    preview = '--preview' in sys.argv
    excel_path = EXCEL_PATH

    for i, arg in enumerate(sys.argv):
        if arg == '--excel' and i + 1 < len(sys.argv):
            excel_path = sys.argv[i + 1]

    # Verificar archivos
    if not os.path.exists(excel_path):
        print(f'ERROR: No se encontro el Excel en: {excel_path}')
        sys.exit(1)

    if not os.path.exists(HTML_PATH):
        print(f'ERROR: No se encontro index.html en: {HTML_PATH}')
        sys.exit(1)

    # Leer datos
    print(f'Leyendo Excel: {excel_path}')
    products = read_excel(excel_path)
    print(f'  Productos encontrados: {len(products)}')

    # Estadisticas
    sections = {}
    brands = {}
    for p in products:
        sections[p['seccion']] = sections.get(p['seccion'], 0) + 1
        brands[p['marca']] = brands.get(p['marca'], 0) + 1

    print(f'\n  Secciones:')
    for sec, count in sections.items():
        print(f'    {SECTION_TITLES.get(sec, sec)}: {count} productos')

    print(f'\n  Marcas:')
    for marca, count in sorted(brands.items()):
        print(f'    {marca}: {count} productos')

    # Validaciones
    errors = []
    for i, p in enumerate(products, 2):
        if p['seccion'] not in SECTION_TITLES:
            errors.append(f'  Fila {i}: seccion "{p["seccion"]}" no es valida. Opciones: {", ".join(SECTION_TITLES.keys())}')
        if not p['imagen']:
            errors.append(f'  Fila {i}: "{p["nombre"]}" no tiene imagen')
        elif not os.path.exists(os.path.join(SCRIPT_DIR, p['imagen'])):
            errors.append(f'  Fila {i}: imagen "{p["imagen"]}" no existe')

    if errors:
        print(f'\n  ADVERTENCIAS ({len(errors)}):')
        for e in errors:
            print(e)

    # Leer y actualizar HTML
    print(f'\nLeyendo HTML: {HTML_PATH}')
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    new_html = update_html(html, products, preview)

    if preview:
        print('\n--- MODO PREVIEW: no se guardo nada ---')
        print('Ejecuta sin --preview para aplicar los cambios.')
        return

    # Backup
    backup_path = create_backup(HTML_PATH)
    print(f'  Backup creado: {backup_path}')

    # Guardar
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(new_html)

    print(f'\n  LISTO! Pagina actualizada con {len(products)} productos.')
    print(f'  Abri index.html en el navegador para verificar.')


if __name__ == '__main__':
    main()
