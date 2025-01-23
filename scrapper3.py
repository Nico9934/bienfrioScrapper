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
from datetime import datetime

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

                # Verificar si no hay stock
                out_of_stock_span = card.find("span", class_="out_of_stock")
                if out_of_stock_span:
                    current_price = "SIN STOCK"
                    base_price = None
                else:
                    # Precio base y precio con descuento
                    price_span = card.find("span", class_="price")
                    base_price = None
                    discount_price = None

                    if price_span:
                        # Precio base (tachado)
                        base_del = price_span.find("del", attrs={"aria-hidden": "true"})
                        if base_del:
                            base_price = (
                                base_del.find("bdi").text.strip()
                                .replace("$", "")
                                .replace(".", "")
                                .replace(",", ".")
                            )

                        # Precio con descuento (ins)
                        price_ins = price_span.find("ins", attrs={"aria-hidden": "true"})
                        if price_ins:
                            discount_price = (
                                price_ins.find("bdi").text.strip()
                                .replace("$", "")
                                .replace(".", "")
                                .replace(",", ".")
                            )
                        else:
                            # Caso: Sin descuento, usar precio regular
                            price_regular = price_span.find("span", class_="woocommerce-Price-amount")
                            if price_regular:
                                discount_price = (
                                    price_regular.find("bdi").text.strip()
                                    .replace("$", "")
                                    .replace(".", "")
                                    .replace(",", ".")
                                )

                # Limpieza y validación de los precios
                try:
                    base_price = float(base_price) if base_price else None
                except ValueError:
                    print(f"Error al convertir precio base: {base_price}")
                    base_price = None

                try:
                    discount_price = float(discount_price) if discount_price and discount_price != "SIN STOCK" else None
                except ValueError:
                    print(f"Error al convertir precio con descuento: {discount_price}")
                    discount_price = None

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
                final_price = calculate_final_price(discount_price, profit_margin)
                rounded_price = round_up_price(final_price)

                # Agregamos los datos procesados
                all_products.append({
                    "Producto": clean_title(title),
                    "Precio Base ($)": base_price,
                    "Precio con Descuento ($)": discount_price,
                    "Peso (kg)": extract_weight(title),
                    "Categoría": category_name.capitalize(),
                    "Mínimo de Compra": min_purchase,
                    "Precio Final ($)": final_price,
                    "Precio Redondeado ($)": rounded_price
                })

    # Convertimos los datos a un DataFrame
    df = pd.DataFrame(all_products)

    # Guardamos los datos en un archivo Excel
    excel_file = f"productos_biomac_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(excel_file, index=False, engine='openpyxl')
    print(f"Archivo Excel guardado como {excel_file}.")
else:
    print("Error al iniciar sesión. Verifica tus credenciales.")
