import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import math

# Credenciales de inicio de sesión
login_url = "https://reventa.biomac.com.ar/wp-login.php"
categories = [
    "https://reventa.biomac.com.ar/categoria-producto/vegetales/",
    "https://reventa.biomac.com.ar/categoria-producto/frutas/",
    "https://reventa.biomac.com.ar/categoria-producto/helados/",
]

# Reemplaza con tus credenciales
credentials = {
    "log": "",  # Cambia esto por tu nombre de usuario
    "pwd": "",  # Cambia esto por tu contraseña
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
                discount = re.search(r"(\d+)", discount_span.text).group(1) if discount_span else "0"

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

                # Mínimo de compra (obtenido del atributo "min")
                quantity_div = card.find("div", class_="quantity")
                min_purchase = "N/A"
                if quantity_div:
                    min_input = quantity_div.find("input", class_="input-text qty text")
                    if min_input and min_input.get("min"):
                        min_purchase = min_input["min"]

                # Porcentaje de ganancia predeterminado (en Excel será editable)
                profit_margin = 30

                # Precio final con ganancia
                def calculate_final_price(price, margin):
                    if price and price.isnumeric():
                        return float(price) * (1 + margin / 100)
                    return None

                # Redondeo hacia arriba
                def round_up_price(price):
                    if price:
                        return int(math.ceil(price / 100.0) * 100)
                    return None

                # Calculamos el precio final y el redondeado
                final_price = calculate_final_price(current_price, profit_margin)
                rounded_price = round_up_price(final_price)

                # Agregamos los datos procesados
                all_products.append({
                    "Producto": clean_title(title),
                    "Precio Base ($)": current_price.replace(".", ","),  # Reemplazamos el punto por coma
                    "Peso (kg)": extract_weight(title),
                    "Descuento": discount,
                    "Mínimo de Compra": min_purchase,
                    "Categoría": category_name,
                    "Porcentaje de Ganancia (%)": profit_margin,
                    "Precio Final ($)": str(final_price).replace(".", ","),  # Formato para Excel
                    "Precio Redondeado ($)": rounded_price
                })

        else:
            print(f"Error al acceder a la página protegida: {response.status_code} - {category_url}")
    
    # Guardamos los datos en un DataFrame
    df = pd.DataFrame(all_products)
    
    # Guardamos el DataFrame en un archivo Excel
    output_file = "productos_biomac.xlsx"
    df.to_excel(output_file, index=False, engine="openpyxl")
    
    print(f"Datos extraídos y guardados en: {output_file}")
else:
    print("Error en el inicio de sesión. Verifica tus credenciales.")
