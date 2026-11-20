"""Regenera data/validaciones.json y data/ordenes.json desde los Excel de PNCIPROY2."""
from pathlib import Path
import json
import unicodedata
from openpyxl import load_workbook

DESKTOP = Path.home() / "Desktop"
CANDIDATOS = [
    DESKTOP,
    Path.home() / "Downloads",
    Path(r"C:\Users\Procesos\Desktop"),
    DESKTOP / "Registro No Conforme - pagina corregida",
    DESKTOP / "PROYECTO" / "PNCIPROY2",
    DESKTOP / "PROYECTO" / "PNCIPROY",
    DESKTOP / "PNCIPROY",
    Path("/home/ubuntu/.cursor/projects/agent-repos-registro-no-conforme/uploads"),
]


def primer_existente(nombre):
    for carpeta in CANDIDATOS:
        ruta = carpeta / nombre
        if ruta.exists():
            return ruta
    if "SEGUIMIENTO" in nombre.upper():
        for carpeta in CANDIDATOS:
            if not carpeta.exists():
                continue
            hits = sorted(carpeta.glob("SEGUIMIENTO*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
            if hits:
                return hits[0]
    raise FileNotFoundError("No se encontró " + nombre + " en: " + ", ".join(str(c) for c in CANDIDATOS))


OUT = Path(__file__).resolve().parents[1] / "data"


def texto(v):
    if v is None:
        return ""
    return str(v).strip()


def col_values(ws, col_idx):
    seen = set()
    vals = []
    for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, values_only=True):
        s = texto(row[0])
        if s and s not in seen:
            seen.add(s)
            vals.append(s)
    return vals


def extraer_validaciones(xlsx=None):
    ruta = Path(xlsx) if xlsx else primer_existente("VALIDACIONES.xlsx")
    wb = load_workbook(ruta, data_only=True)
    ws = wb.active
    data = {
        "mes": col_values(ws, 1),
        "proceso_detecta": col_values(ws, 4),
        "maquina_detecta": col_values(ws, 7),
        "operador_detecta": col_values(ws, 10),
        "proceso_origina": col_values(ws, 13),
        "maquina": col_values(ws, 16),
        "operador": col_values(ws, 19),
        "defecto": col_values(ws, 22),
        "reporta": col_values(ws, 25),
        "autorizo": col_values(ws, 28),
    }
    wb.close()
    return data, ruta


def normalizar_encabezado(v):
    s = texto(v)
    for enc, dec in (("latin-1", "utf-8"), ("cp1252", "utf-8")):
        try:
            s2 = s.encode(enc).decode(dec)
            if "Ã" not in s2:
                s = s2
                break
        except Exception:
            pass
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.lower().split())


def indice_encabezado(normas, *claves):
    for clave in claves:
        c = normalizar_encabezado(clave)
        for i, n in enumerate(normas):
            if n == c:
                return i
    return None


def indices_ordenes(headers):
    normas = [normalizar_encabezado(h) for h in headers]
    sku = indice_encabezado(normas, "SKU")
    id_art = None
    id_prod = None
    desc = None
    for i, n in enumerate(normas):
        if "id producto" in n:
            id_prod = i
        elif n.startswith("id art"):
            id_art = i
        elif n.startswith("descrip"):
            desc = i
    orden = indice_encabezado(normas, "ID Orden de Venta")
    oc = indice_encabezado(normas, "Orden Compra")
    cliente = indice_encabezado(normas, "Nombre Cliente")
    # Fallback a las columnas históricas si faltan encabezados.
    if orden is None:
        orden = 1
    if oc is None:
        oc = 2
    if cliente is None:
        cliente = 5
    if sku is None:
        sku = 10
    if id_art is None:
        id_art = 11
    if id_prod is not None:
        producto = desc if desc is not None else id_prod + 1
    else:
        producto = desc if desc is not None else 12
    return {
        "orden": orden,
        "oc": oc,
        "cliente": cliente,
        "sku": sku,
        "id_art": id_art,
        "id_prod": id_prod,
        "producto": producto,
    }


def celda(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return texto(row[idx])


def extraer_ordenes(xlsx=None):
    ruta = Path(xlsx) if xlsx else primer_existente("SEGUIMIENTO DE OT1 (1).xlsx")
    wb = load_workbook(ruta, read_only=True, data_only=True)
    hoja = "rep-ov-espec" if "rep-ov-espec" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[hoja]
    headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    idx = indices_ordenes(headers)
    registros = []
    vacias = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        orden = celda(row, idx["orden"])
        oc = celda(row, idx["oc"])
        if orden == "" and oc == "":
            vacias += 1
            if vacias >= 20:
                break
            continue
        vacias = 0
        registros.append({
            "orden": orden,
            "ordenCompra": oc,
            "cliente": celda(row, idx["cliente"]),
            "sku": celda(row, idx["sku"]),
            "idArticulo": celda(row, idx["id_art"]),
            "idProducto": celda(row, idx["id_prod"]),
            "producto": celda(row, idx["producto"]),
        })
    wb.close()
    return registros, ruta, idx


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    try:
        validaciones, ruta_val = extraer_validaciones()
        (OUT / "validaciones.json").write_text(
            json.dumps(validaciones, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print("VALIDACIONES:", ruta_val)
        print("validaciones.json:", {k: len(v) for k, v in validaciones.items()})
    except FileNotFoundError as err:
        print("validaciones.json: se omite (" + str(err) + ")")
    ordenes, ruta_ot, idx = extraer_ordenes()
    (OUT / "ordenes.json").write_text(
        json.dumps(ordenes, ensure_ascii=False), encoding="utf-8"
    )
    print("SEGUIMIENTO:", ruta_ot)
    print("columnas:", idx)
    print("ordenes.json:", len(ordenes), "registros")
    con_id = sum(1 for r in ordenes if r.get("idProducto"))
    print("con ID Producto:", con_id)


if __name__ == "__main__":
    main()
