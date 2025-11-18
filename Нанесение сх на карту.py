import requests
import urllib3
import pandas as pd
from shapely.geometry import shape
from shapely.ops import transform
from pyproj import Transformer
import os

# Отключаем предупреждения о сертификатах
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Функция для запроса к API НСПД
def get_cadastral_data(cadastral_number):
    url = f"https://nspd.gov.ru/api/geoportal/v2/search/geoportal?thematicSearchId=1&query={cadastral_number}"
    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Referer': 'https://nspd.gov.ru/',
        'Accept': 'application/json'
    }

    try:
        response = requests.get(url, headers=headers, verify=False, timeout=10, proxies={})
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка при получении {cadastral_number}: {e}")
        return None

# Обработка одного участка
def extract_plot_info(feature):
    props = feature.get('properties', {})
    geometry = feature.get('geometry')

    if geometry is None:
        print(f"⚠️ Пропущен участок — отсутствует геометрия: {props.get('label', 'Без названия')}")
        return None

    try:
        geom = shape(geometry)
        # Центроид в EPSG:3857
        centroid_3857 = geom.centroid

        # Преобразуем в WGS84 (широта, долгота)
        transformer = Transformer.from_crs("EPSG:3857", "EPSG:4326", always_xy=True)
        lon, lat = transformer.transform(centroid_3857.x, centroid_3857.y)

        return {
            "Кадастровый номер": props.get("label", ""),
            "Адрес": props.get("options", {}).get("readable_address", ""),
            "Категория": props.get("options", {}).get("land_record_category_type", ""),
            "Разрешённое использование": props.get("options", {}).get("permitted_use_established_by_document", ""),
            "Площадь, м²": props.get("options", {}).get("specified_area", ""),
            "Долгота": lon,
            "Широта": lat
        }
    except Exception as e:
        print(f"❌ Ошибка при обработке геометрии участка {props.get('label', '')}: {e}")
        return None

# 📋 Тестовые кадастровые номера (можно заменить на свои)
cadastral_list = [
    "50:04:0010209:57",
    "22:61:053901:569",
    "22:61:000000:1360",
    "22:61:053801:217",
    "22:61:053601:150",
    "22:61:000000:1449",
    "22:61:053501:56",
    "22:61:000000:93",
    "22:61:052601:26",
    "22:61:052501:127",
    "22:61:000000:113",
    "22:61:000000:630",
]

# 📦 Сбор данных
results = []
for cad_num in cadastral_list:
    response = get_cadastral_data(cad_num)
    if response and response.get("data", {}).get("features"):
        feature = response["data"]["features"][0]
        land_type = feature["properties"]["options"].get("land_record_category_type", "Не указана")
        print(f"✔️ {cad_num} — категория: {land_type}")
        info = extract_plot_info(feature)
        if info:
            results.append(info)
    else:
        print(f"⚠️ Не найден участок: {cad_num}")

# 💾 Экспорт в Excel
if results:
    output_path = os.path.join(r"C:\Users\nkazakov\Downloads", "Участки_НСПД обл.xlsx")
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False)
    print(f"\n✅ Данные сохранены в файл: {output_path}")
else:
    print("\n⚠️ Не удалось получить данные ни по одному участку. Проверьте номера.")
