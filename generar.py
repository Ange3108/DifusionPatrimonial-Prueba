"""Genera datos.json, imágenes de relleno y un QR por registro.

Uso:
    python generar.py URL_DEL_SITIO [ARCHIVO.xlsx|ARCHIVO.csv]

Ejemplo:
    python generar.py https://ange3108.github.io/DifusionPatrimonial-Prueba/ registros.xlsx

Acepta el archivo exportado desde Microsoft Lists (Excel o CSV).
Si una fila ya tiene su imagen en imagenes/ID.(jpg|jpeg|png|webp|svg) se usa esa;
si no, se crea una imagen de relleno.

Requiere: pip install pandas openpyxl "qrcode[pil]"
"""
import json, os, sys
import pandas as pd
import qrcode

BASE_URL = sys.argv[1] if len(sys.argv) > 1 else "https://ange3108.github.io/DifusionPatrimonial-Prueba/"
if not BASE_URL.endswith("/"):
    BASE_URL += "/"
ARCHIVO = sys.argv[2] if len(sys.argv) > 2 else "lista_sharepoint_prueba.xlsx"

os.makedirs("imagenes", exist_ok=True)
os.makedirs("qr", exist_ok=True)

if ARCHIVO.lower().endswith(".csv"):
    df = pd.read_csv(ARCHIVO, encoding="utf-8-sig")
else:
    df = pd.read_excel(ARCHIVO)


def texto(r, col):
    v = r.get(col, "")
    return "" if pd.isna(v) else str(v)


def fecha(r, col):
    v = r.get(col)
    if v is None or pd.isna(v):
        return ""
    return pd.to_datetime(v, dayfirst=True).strftime("%Y-%m-%d")


def avance(r):
    v = r.get("Avance", 0)
    if isinstance(v, str):
        v = float(v.replace("%", "").replace(",", ".") or 0)
        v = v / 100 if v > 1 else v
    v = 0.0 if pd.isna(v) else float(v)
    return v / 100 if v > 1 else v


def imagen_para(n):
    for ext in ("jpg", "jpeg", "png", "webp", "svg"):
        ruta = f"imagenes/{n}.{ext}"
        if os.path.exists(ruta):
            return ruta
    # Relleno vintage; reemplácelo con la foto real usando el mismo nombre.
    paleta = ["#8a2f29", "#3d5a4f", "#c4912c", "#4a5f8a", "#6b5540"]
    c = paleta[n % len(paleta)]
    ruta = f"imagenes/{n}.svg"
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(
            '<svg xmlns="http://www.w3.org/2000/svg" width="600" height="320" viewBox="0 0 600 320">'
            f'<rect width="600" height="320" fill="{c}"/>'
            '<rect x="14" y="14" width="572" height="292" fill="none" stroke="#f1e6cc" stroke-width="2" opacity=".7"/>'
            '<rect x="24" y="24" width="552" height="272" fill="none" stroke="#f1e6cc" stroke-width="1" opacity=".5"/>'
            '<text x="300" y="120" font-family="Georgia,serif" font-size="20" fill="#f1e6cc" '
            'text-anchor="middle" letter-spacing="6">REGISTRO</text>'
            f'<text x="300" y="215" font-family="Georgia,serif" font-weight="bold" font-size="84" fill="#f1e6cc" '
            f'text-anchor="middle">{n:03d}</text></svg>'
        )
    return ruta


registros = []
for i, r in df.iterrows():
    n = i + 1
    registros.append({
        "id": n,
        "titulo": texto(r, "Título"),
        "descripcion": texto(r, "Descripción"),
        "departamento": texto(r, "Departamento"),
        "estado": texto(r, "Estado"),
        "prioridad": texto(r, "Prioridad"),
        "responsable": texto(r, "Responsable"),
        "fecha_inicio": fecha(r, "Fecha de inicio"),
        "fecha_entrega": fecha(r, "Fecha de entrega"),
        "avance": avance(r),
        "imagen": imagen_para(n),
    })
    qrcode.make(f"{BASE_URL}detalle.html?id={n}").save(f"qr/{n:02d}.png")

with open("datos.json", "w", encoding="utf-8") as f:
    json.dump(registros, f, ensure_ascii=False, indent=2)

print(f"{len(registros)} registros desde {ARCHIVO}; QR apuntando a {BASE_URL}")
