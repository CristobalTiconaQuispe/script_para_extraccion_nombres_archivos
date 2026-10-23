import csv
import io
import os
import shutil
import sys
from datetime import datetime


CARPETA_ORIGEN = r"C:\Users\Cristobal\Desktop\origen"
CARPETA_DESTINO = r"C:\Users\Cristobal\Desktop\destino"
ARCHIVO_TABLA = os.path.join(CARPETA_DESTINO, "registro_archivos.csv")

COLUMNAS = ["Nombre del Archivo", "Estado", "Fecha Registro", "Fecha Movimiento"]

ESTADOS_VERIFICADOS = {"VERIFICADO", "VERIFICADA", "SI", "SÍ", "OK", "X"}

ARCHIVOS_IGNORADOS = {"desktop.ini", "thumbs.db"}


def ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def configurar_consola():
    for flujo in (sys.stdout, sys.stderr):
        if hasattr(flujo, "reconfigure"):
            try:
                flujo.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def normalizar_estado(valor):
    return valor.strip().upper()


def leer_texto(ruta):
    for codificacion in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(ruta, "r", encoding=codificacion, newline="") as archivo:
                return archivo.read()
        except UnicodeDecodeError:
            continue
    raise OSError(f"No se pudo leer el archivo: {ruta}")


def detectar_delimitador(texto):
    primera_linea = texto.splitlines()[0] if texto.splitlines() else ""
    if ";" in primera_linea:
        return ";"
    if "\t" in primera_linea:
        return "\t"
    if "," in primera_linea:
        return ","
    return ";"


def leer_tabla():
    if not os.path.exists(ARCHIVO_TABLA):
        return [], True

    texto = leer_texto(ARCHIVO_TABLA)

    lineas = texto.splitlines()
    if lineas and lineas[0].strip().lower().startswith("sep="):
        texto = "\n".join(lineas[1:])

    delimitador = detectar_delimitador(texto)

    filas_crudas = [
        fila
        for fila in csv.reader(io.StringIO(texto), delimiter=delimitador)
        if any(celda.strip() for celda in fila)
    ]

    if not filas_crudas:
        return [], True

    encabezado = [celda.strip().lower() for celda in filas_crudas[0]]

    if "estado" in encabezado and "nombre del archivo" in encabezado:
        indices = {
            columna: encabezado.index(columna.lower())
            for columna in COLUMNAS
            if columna.lower() in encabezado
        }
        filas = []
        for cruda in filas_crudas[1:]:
            fila = {}
            for columna in COLUMNAS:
                indice = indices.get(columna)
                if indice is not None and indice < len(cruda):
                    fila[columna] = cruda[indice]
                else:
                    fila[columna] = ""
            if fila["Nombre del Archivo"]:
                filas.append(fila)
        return filas, False

    inicio = 1 if encabezado and encabezado[0] == "nombre del archivo" else 0
    filas = []
    for cruda in filas_crudas[inicio:]:
        nombre = cruda[0] if cruda else ""
        if not nombre:
            continue
        estado = "MOVIDO" if os.path.exists(os.path.join(CARPETA_DESTINO, nombre)) else "PENDIENTE"
        filas.append(
            {
                "Nombre del Archivo": nombre,
                "Estado": estado,
                "Fecha Registro": "",
                "Fecha Movimiento": "",
            }
        )
    return filas, True


def guardar_tabla(filas):
    os.makedirs(CARPETA_DESTINO, exist_ok=True)
    temporal = ARCHIVO_TABLA + ".tmp"
    try:
        with open(temporal, "w", encoding="utf-8-sig", newline="") as archivo:
            escritor = csv.DictWriter(archivo, fieldnames=COLUMNAS, delimiter=";")
            escritor.writeheader()
            for fila in filas:
                escritor.writerow({columna: fila.get(columna, "") for columna in COLUMNAS})
        os.replace(temporal, ARCHIVO_TABLA)
    except OSError:
        if os.path.exists(temporal):
            try:
                os.remove(temporal)
            except OSError:
                pass
        raise


def guardar_con_aviso(filas):
    try:
        guardar_tabla(filas)
        return True
    except OSError as error:
        print(f"[ERROR] No se pudo guardar la tabla: {error}")
        print("Cierra el archivo en Excel y vuelve a ejecutar el comando.")
        return False


def tabla_escribible():
    if not os.path.exists(ARCHIVO_TABLA):
        return True
    try:
        descriptor = os.open(ARCHIVO_TABLA, os.O_RDWR)
        os.close(descriptor)
        return True
    except OSError:
        return False


def obtener_archivos_origen():
    if not os.path.isdir(CARPETA_ORIGEN):
        raise FileNotFoundError(CARPETA_ORIGEN)

    nombre_tabla = os.path.basename(ARCHIVO_TABLA).casefold()
    nombres = []
    with os.scandir(CARPETA_ORIGEN) as entradas:
        for entrada in entradas:
            if not entrada.is_file():
                continue
            nombre = entrada.name
            if nombre.casefold() == nombre_tabla:
                continue
            if nombre.startswith("~$") or nombre.casefold() in ARCHIVOS_IGNORADOS:
                continue
            nombres.append(nombre)
    nombres.sort(key=str.casefold)
    return nombres


def nombre_valido(nombre):
    return (
        bool(nombre)
        and nombre not in {".", ".."}
        and not os.path.isabs(nombre)
        and os.path.basename(nombre) == nombre
    )


def nueva_fila(nombre, estado):
    return {
        "Nombre del Archivo": nombre,
        "Estado": estado,
        "Fecha Registro": ahora(),
        "Fecha Movimiento": "",
    }


def mapa_por_nombre(filas):
    mapa = {}
    for fila in filas:
        nombre = fila.get("Nombre del Archivo", "")
        if nombre:
            mapa.setdefault(nombre.casefold(), []).append(fila)
    return mapa


def fila_activa(coincidencias):
    for fila in coincidencias:
        estado = normalizar_estado(fila.get("Estado", ""))
        if fila.get("Fecha Movimiento", "").strip():
            continue
        if estado == "MOVIDO":
            continue
        return fila
    return None


def tiene_movida(coincidencias):
    for fila in coincidencias:
        if fila.get("Fecha Movimiento", "").strip():
            return True
        if normalizar_estado(fila.get("Estado", "")) == "MOVIDO":
            return True
    return False


def intentar_mover(fila):
    nombre = fila.get("Nombre del Archivo", "")
    ruta_origen = os.path.join(CARPETA_ORIGEN, nombre)
    ruta_destino = os.path.join(CARPETA_DESTINO, nombre)

    if not os.path.isfile(ruta_origen):
        fila["Estado"] = "NO ENCONTRADO"
        print(f"[AVISO] No existe en origen: {nombre}")
        return "no_encontrado"

    if os.path.exists(ruta_destino):
        fila["Estado"] = "CONFLICTO"
        print(f"[CONFLICTO] Ya existe en destino, no se movio: {nombre}")
        return "conflicto"

    try:
        shutil.move(ruta_origen, ruta_destino)
    except OSError as error:
        print(f"[ERROR] No se pudo mover {nombre}: {error}")
        return "error"

    fila["Estado"] = "VERIFICADO"
    fila["Fecha Movimiento"] = ahora()
    print(f"[OK] Movido: {nombre}")
    return "movido"


def listar():
    try:
        nombres = obtener_archivos_origen()
    except FileNotFoundError:
        print(f"[ERROR] La carpeta origen no existe: {CARPETA_ORIGEN}")
        return 1

    os.makedirs(CARPETA_DESTINO, exist_ok=True)
    filas, migrada = leer_tabla()

    registrados = {
        fila["Nombre del Archivo"].casefold()
        for fila in filas
        if fila["Nombre del Archivo"]
    }
    nuevos = []

    for nombre in nombres:
        if nombre.casefold() in registrados:
            continue
        filas.append(
            {
                "Nombre del Archivo": nombre,
                "Estado": "PENDIENTE",
                "Fecha Registro": ahora(),
                "Fecha Movimiento": "",
            }
        )
        registrados.add(nombre.casefold())
        nuevos.append(nombre)

    if (nuevos or migrada) and not guardar_con_aviso(filas):
        return 1

    print("=" * 60)
    if nuevos:
        print(f"Nombres registrados como PENDIENTE ({len(nuevos)}):")
        for nombre in nuevos:
            print(f"  - {nombre}")
    else:
        print("No hay nombres nuevos para registrar.")
    if migrada:
        print("La tabla fue creada o convertida al formato de 4 columnas.")

    pendientes = sum(
        1 for fila in filas if normalizar_estado(fila["Estado"]) == "PENDIENTE"
    )
    print(f"Total de nombres en la tabla: {len(filas)}")
    print(f"Pendientes de verificar: {pendientes}")
    print("=" * 60)
    print(f"Tabla: {ARCHIVO_TABLA}")
    print("Siguiente paso:")
    print("  1. Abre la tabla en Excel y cambia Estado de PENDIENTE a VERIFICADO.")
    print("  2. Guarda el archivo como CSV UTF-8.")
    print("  3. Ejecuta: python extra.py mover")
    return 0


def verificar():
    if not os.path.isdir(CARPETA_ORIGEN):
        print(f"[ERROR] La carpeta origen no existe: {CARPETA_ORIGEN}")
        return 1

    if not tabla_escribible():
        print("[ERROR] La tabla esta abierta en Excel o no se puede escribir.")
        print("Cierra el archivo y vuelve a ejecutar el comando.")
        return 1

    os.makedirs(CARPETA_DESTINO, exist_ok=True)

    filas, migrada = leer_tabla()
    if migrada and not guardar_con_aviso(filas):
        return 1

    por_nombre = mapa_por_nombre(filas)

    movidos = 0
    nuevos = 0
    duplicados = 0
    conflictos = 0
    no_encontrados = 0
    errores = 0

    for nombre in obtener_archivos_origen():
        if not nombre_valido(nombre):
            errores += 1
            print(f"[ERROR] Nombre no valido: {nombre!r}")
            continue

        clave = nombre.casefold()
        coincidencias = por_nombre.get(clave, [])

        fila = fila_activa(coincidencias)

        existe_en_destino = os.path.exists(os.path.join(CARPETA_DESTINO, nombre))

        if fila is None:
            if existe_en_destino or tiene_movida(coincidencias):
                fila = nueva_fila(nombre, "PENDIENTE")
                filas.append(fila)
                por_nombre.setdefault(clave, []).append(fila)
                duplicados += 1
                print(f"[DUPLICADO] Queda PENDIENTE: {nombre}")
            else:
                fila = nueva_fila(nombre, "VERIFICADO")
                filas.append(fila)
                por_nombre.setdefault(clave, []).append(fila)
                nuevos += 1
                resultado = intentar_mover(fila)
                if resultado == "movido":
                    movidos += 1
                elif resultado == "conflicto":
                    conflictos += 1
                elif resultado == "no_encontrado":
                    no_encontrados += 1
                elif resultado == "error":
                    errores += 1
            if not guardar_con_aviso(filas):
                return 1
            continue

        estado = normalizar_estado(fila.get("Estado", ""))

        if existe_en_destino:
            if estado in ("", "PENDIENTE"):
                duplicados += 1
                print(f"[DUPLICADO] Sigue PENDIENTE: {nombre}")
            elif estado in ESTADOS_VERIFICADOS:
                fila["Estado"] = "CONFLICTO"
                conflictos += 1
                print(f"[CONFLICTO] Ya existe en destino: {nombre}")
                if not guardar_con_aviso(filas):
                    return 1
            else:
                duplicados += 1
                print(f"[DUPLICADO] Estado {estado or 'VACIO'}, no se movio: {nombre}")
            continue

        if estado in ESTADOS_VERIFICADOS or estado in (
            "",
            "PENDIENTE",
            "CONFLICTO",
            "NO ENCONTRADO",
            "ERROR",
        ):
            fila["Estado"] = "VERIFICADO"
            if not fila.get("Fecha Registro", "").strip():
                fila["Fecha Registro"] = ahora()
            resultado = intentar_mover(fila)
            if resultado == "movido":
                movidos += 1
            elif resultado == "conflicto":
                conflictos += 1
            elif resultado == "no_encontrado":
                no_encontrados += 1
            elif resultado == "error":
                errores += 1
            if not guardar_con_aviso(filas):
                return 1
            continue

        duplicados += 1
        print(f"[DUPLICADO] Estado {estado or 'VACIO'}, no se movio: {nombre}")

    print("=" * 60)
    print(f"Verificados y movidos automaticamente: {movidos}")
    print(f"Duplicados que quedan PENDIENTE: {duplicados}")
    print(f"Conflictos: {conflictos}")
    print(f"No encontrados: {no_encontrados}")
    print(f"Errores: {errores}")
    print(f"Nombres nuevos registrados: {nuevos}")
    print("=" * 60)
    return 1 if errores else 0


def mover():
    if not os.path.isdir(CARPETA_ORIGEN):
        print(f"[ERROR] La carpeta origen no existe: {CARPETA_ORIGEN}")
        return 1

    if not os.path.exists(ARCHIVO_TABLA):
        print("[ERROR] No existe la tabla de registro.")
        print("Ejecuta primero: python extra.py listar")
        return 1

    if not tabla_escribible():
        print("[ERROR] La tabla esta abierta en Excel o no se puede escribir.")
        print("Cierra el archivo y vuelve a ejecutar el comando.")
        return 1

    os.makedirs(CARPETA_DESTINO, exist_ok=True)
    filas, migrada = leer_tabla()

    if migrada and not guardar_con_aviso(filas):
        return 1

    movidos = 0
    no_encontrados = 0
    conflictos = 0
    errores = 0

    for fila in filas:
        if normalizar_estado(fila.get("Estado", "")) not in ESTADOS_VERIFICADOS:
            continue
        if fila.get("Fecha Movimiento", "").strip():
            continue

        nombre = fila.get("Nombre del Archivo", "")

        if not nombre_valido(nombre):
            fila["Estado"] = "ERROR"
            errores += 1
            print(f"[ERROR] Nombre no valido en la tabla: {nombre!r}")
            if not guardar_con_aviso(filas):
                return 1
            continue

        resultado = intentar_mover(fila)
        if resultado == "movido":
            movidos += 1
        elif resultado == "conflicto":
            conflictos += 1
        elif resultado == "no_encontrado":
            no_encontrados += 1
        elif resultado == "error":
            errores += 1
        if not guardar_con_aviso(filas):
            return 1

    print("=" * 60)
    print(f"Movidos: {movidos}")
    print(f"Conflictos: {conflictos}")
    print(f"No encontrados: {no_encontrados}")
    print(f"Errores: {errores}")
    print("=" * 60)
    return 1 if errores else 0


def simular():
    try:
        nombres = obtener_archivos_origen()
    except FileNotFoundError:
        print(f"[ERROR] La carpeta origen no existe: {CARPETA_ORIGEN}")
        return 1

    filas, migrada = leer_tabla()
    por_nombre = mapa_por_nombre(filas)

    moverian = 0
    pendientes = 0
    conflictos = 0

    print("=" * 60)
    print("SIMULACION: no se escribe ni se mueve ningun archivo")
    print("=" * 60)
    if migrada:
        print("La tabla se convertiria al formato de 4 columnas.")

    for nombre in nombres:
        if not nombre_valido(nombre):
            print(f"  [ERROR] Nombre no valido: {nombre!r}")
            continue

        coincidencias = por_nombre.get(nombre.casefold(), [])
        fila = fila_activa(coincidencias)
        existe_en_destino = os.path.exists(os.path.join(CARPETA_DESTINO, nombre))

        if fila is None:
            if existe_en_destino or tiene_movida(coincidencias):
                pendientes += 1
                print(f"  [DUPLICADO -> PENDIENTE] {nombre}")
            else:
                moverian += 1
                print(f"  [VERIFICAR Y MOVER] {nombre}")
            continue

        estado = normalizar_estado(fila.get("Estado", ""))

        if existe_en_destino:
            if estado in ("", "PENDIENTE"):
                pendientes += 1
                print(f"  [DUPLICADO -> PENDIENTE] {nombre}")
            else:
                conflictos += 1
                print(f"  [CONFLICTO] {nombre}")
        elif estado in ESTADOS_VERIFICADOS or estado in (
            "",
            "PENDIENTE",
            "CONFLICTO",
            "NO ENCONTRADO",
            "ERROR",
        ):
            moverian += 1
            print(f"  [VERIFICAR Y MOVER] {nombre}")
        else:
            pendientes += 1
            print(f"  [DUPLICADO] {nombre}")

    if not nombres:
        print("  (no hay archivos en origen)")

    print("=" * 60)
    print(f"Se verificarian y moverian: {moverian}")
    print(f"Quedarian PENDIENTE: {pendientes}")
    print(f"Conflictos: {conflictos}")
    print("=" * 60)
    return 0


def mostrar_ayuda():
    print("Uso:")
    print("  python extra.py verificar  Verifica y mueve automatico; duplicados quedan PENDIENTE")
    print("  python extra.py listar     Solo registra nombres nuevos como PENDIENTE")
    print("  python extra.py mover      Mueve los VERIFICADO que aun no se han movido")
    print("  python extra.py simular    Muestra que haria verificar sin escribir ni mover")
    print("  python extra.py ayuda      Muestra esta ayuda")


def main():
    configurar_consola()
    accion = sys.argv[1].strip().lower() if len(sys.argv) > 1 else "verificar"

    if accion in ("ayuda", "help", "-h", "--help"):
        mostrar_ayuda()
        return 0

    acciones = {
        "verificar": verificar,
        "procesar": verificar,
        "automatico": verificar,
        "listar": listar,
        "registrar": listar,
        "mover": mover,
        "simular": simular,
        "simulacion": simular,
    }

    funcion = acciones.get(accion)
    if funcion is None:
        print(f"Accion no reconocida: {accion}")
        mostrar_ayuda()
        return 1

    return funcion()


if __name__ == "__main__":
    sys.exit(main())
