# Importar bibliotecas necesarias
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import math
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from dotenv import load_dotenv
import os

# Cargar las variables de entorno desde .env
load_dotenv()

# Credenciales de inicio de sesión
login_url = "https://reventa.biomac.com.ar/wp-login.php"
categories = [
    "https://reventa.biomac.com.ar/categoria-producto/vegetales/",
    "https://reventa.biomac.com.ar/categoria-producto/frutas/",
    "https://reventa.biomac.com.ar/categoria-producto/helados/",
]

# Cargar credenciales desde variables de entorno
credentials = {
    "log": os.getenv("LOG"),  # Usuario desde el archivo .env
    "pwd": os.getenv("PWD"),  # Contraseña desde el archivo .env
    "redirect_to": login_url,
    "testcookie": "1"
}

# Crear una sesión para mantener las cookies
session = requests.Session()

# Realizar el login
response = session.post(login_url, data=credentials)

# Verificar si el login fue exitoso
if response.status_code == 200 and "dashboard" not in response.url:
    print("Inicio de sesión exitoso.")

    all_products = []  # Lista para almacenar todos los productos de todas las categorías

    for category_url in categories:
        # Acceder a la página protegida con la sesión autenticada
        response = session.get(category_url)

        # Verificamos si la solicitud fue exitosa
        if response.status_code == 200:
            print(f"Acceso exitoso a la página de productos: {category_url}")
            soup = BeautifulSoup(response.content, "html.parser")

            # Extraemos las cards de productos
            cards = soup.find_all(class_="item-producto-bio")
            
            # Obtenemos la categoría a partir de la URL
            category_name = re.search(r"categoria-producto/([^/]+)/", category_url)
            category_name = category_name.group(1) if category_name else "Desconocido"

            # Procesamos cada card
            for card in cards:
                # Título del producto
                title = card.find(class_="woocommerce-loop-product__title").text.strip()

                # Descuento (si existe)
                discount_span = card.find("span", class_="onsale off")
                try:
                    discount = (
                        int(re.search(r"(\d+)", discount_span.text).group(1)) if discount_span else 0
                    )
                except ValueError:
                    print(f"Error al convertir descuento: {discount_span.text if discount_span else 'N/A'}")
                    discount = 0

                # Verificar si no hay stock
                out_of_stock_span = card.find("span", class_="out_of_stock")
                if out_of_stock_span:
                    current_price = "SIN STOCK"
                else:
                    # Precio actual (manejar tanto con descuento como sin descuento)
                    price_span = card.find("span", class_="price")
                    current_price = None
                    if price_span:
                        # Caso: Producto con descuento
                        price_ins = price_span.find("ins", attrs={"aria-hidden": "true"})
                        if price_ins:
                            current_price = (
                                price_ins.find("bdi").text.strip()
                                .replace("$", "")
                                .replace(".", "")
                                .replace(",", ".")
                            )
                        else:
                            # Caso: Producto sin descuento
                            price_regular = price_span.find("span", class_="woocommerce-Price-amount")
                            if price_regular:
                                current_price = (
                                    price_regular.find("bdi").text.strip()
                                    .replace("$", "")
                                    .replace(".", "")
                                    .replace(",", ".")
                                )

                # Limpieza y validación del precio base
                try:
                    current_price = float(current_price) if current_price and current_price != "SIN STOCK" else None
                except ValueError:
                    print(f"Error al convertir precio: {current_price}")
                    current_price = None

                # Peso (extraemos en base a regex)
                def extract_weight(title):
                    match = re.search(r"(\d+(?:,\d+)?)(?=\s?(kg|gr|g))", title, re.IGNORECASE)
                    if match:
                        weight = match.group(1).replace(",", ".")  # Convertimos a formato decimal
                        unit = match.group(2).lower()
                        if unit in ["g", "gr"]:
                            return float(weight) / 1000  # Convertimos gramos a kilogramos
                        return float(weight)
                    return None

                # Limpieza del título
                def clean_title(title):
                    return re.sub(r"\s?por\s?.*", "", title, flags=re.IGNORECASE).strip()

                # Mínimo de compra
                quantity_div = card.find("div", class_="quantity")
                min_purchase = "N/A"
                if quantity_div:
                    min_input = quantity_div.find("input", class_="input-text qty text")
                    if min_input:
                        min_value = min_input.get("min", "1")  # Tomar el valor del atributo 'min' (si no, usar 1)
                        value = min_input.get("value", "1")  # Tomar el valor del atributo 'value' como fallback
                        min_purchase = int(min_value) if min_value and int(min_value) > 1 else int(value)

                # Porcentaje de ganancia predeterminado (en Excel será editable)
                profit_margin = 30

                # Precio final con ganancia
                def calculate_final_price(price, margin):
                    if price is not None:
                        return float(price) * (1 + margin / 100)
                    return None

                # Redondeo hacia arriba
                def round_up_price(price):
                    if price is not None:
                        return int(math.ceil(price / 100.0) * 100)
                    return None

                # Calculamos el precio final y el redondeado
                final_price = calculate_final_price(current_price, profit_margin)
                rounded_price = round_up_price(final_price)

                # Agregamos los datos procesados
                all_products.append({
                    "Producto": clean_title(title),
                    "Precio Base ($)": current_price,
                    "Peso (kg)": extract_weight(title),
                    "Descuento (%)": discount,
                    "Mínimo de Compra": min_purchase,
                    "Categoría": category_name,
                    "Porcentaje de Ganancia (%)": profit_margin,
                    "Precio Final ($)": final_price,
                    "Precio Redondeado ($)": rounded_price
                })

        else:
            print(f"Error al acceder a la página protegida: {response.status_code} - {category_url}")
    
    # Guardamos los datos en un DataFrame
    df = pd.DataFrame(all_products)

    # Crear un archivo Excel con estilos
    output_file = "productos_biomac.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(df.columns.tolist())  # Agregar encabezados

    # Estilo de encabezados
    header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for col_idx, header in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border

    # Aplicar datos y bordes a las filas
    category_colors = {
        "vegetales": "00913f",  # Verde
        "frutas": "ffa127",     # Naranja
        "helados": "24afff",    # Celeste
    }

    for index, row in df.iterrows():
        color = category_colors.get(row["Categoría"], "FFFFFF")  # Default blanco
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        font = Font(color="FFFFFF" if row["Categoría"] in category_colors else "000000")

        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=index + 2, column=col_idx)
            cell.value = value
            cell.fill = fill
            cell.font = font
            cell.border = border

    # Agregar filtro en encabezados
    ws.auto_filter.ref = f"A1:I{len(df) + 1}"  # Ajustar el rango según la cantidad de filas

    # Guardar el archivo Excel
    wb.save(output_file)
    print(f"Datos extraídos y guardados en: {output_file}")
else:
    print("Error en el inicio de sesión. Verifica tus credenciales.")
