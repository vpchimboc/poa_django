import csv
import json
from functools import lru_cache

from django.conf import settings
from django.http import HttpResponse, JsonResponse
from django.shortcuts import render

CATEGORIAS = ["Unidad Administrativa", "Unidad Académica", "Carrera"]


@lru_cache(maxsize=1)
def cargar_poa():
    with open(settings.POA_DATA_FILE, encoding="utf-8") as fh:
        return json.load(fh)


def resumen(unidades):
    actividades = [a for u in unidades for a in u["actividades"]]
    con_avance = [a for a in actividades if a.get("avance") is not None]
    promedio = (sum(u["avance"] for u in unidades) / len(unidades)) if unidades else 0
    return {
        "unidades": len(unidades),
        "actividades": len(actividades),
        "completadas": len([a for a in con_avance if a["avance"] >= 1]),
        "sin_iniciar": len([a for a in con_avance if a["avance"] == 0]),
        "observaciones": len([a for a in actividades if a.get("observacion")]),
        "promedio": promedio,
    }


def index(request):
    poa = cargar_poa()
    return render(
        request,
        "dashboard/index.html",
        {
            "poa_json": json.dumps(poa, ensure_ascii=False),
            "anio": poa["anio"],
            "categorias": CATEGORIAS,
            "resumen": resumen(poa["unidades"]),
        },
    )


def api_poa(request):
    """API JSON: soporta ?q= (búsqueda) y ?categoria=."""
    poa = cargar_poa()
    q = request.GET.get("q", "").strip().lower()
    cat = request.GET.get("categoria", "").strip()

    unidades = poa["unidades"]
    if cat and cat != "Todas":
        unidades = [u for u in unidades if u["categoria"] == cat]
    if q:
        unidades = [
            u
            for u in unidades
            if q
            in " ".join(
                str(u.get(k) or "") for k in ("nombre", "titulo", "responsable", "categoria")
            ).lower()
        ]

    return JsonResponse(
        {"anio": poa["anio"], "resumen": resumen(unidades), "unidades": unidades},
        json_dumps_params={"ensure_ascii": False},
    )


def descargar_csv(request):
    """Descarga todas las actividades del POA en CSV (Excel-friendly)."""
    poa = cargar_poa()
    response = HttpResponse(content_type="text/csv; charset=utf-8")
    response["Content-Disposition"] = f'attachment; filename="poa_{poa["anio"]}_actividades.csv"'
    response.write("\ufeff")  # BOM para que Excel reconozca los acentos

    writer = csv.writer(response, delimiter=";")
    writer.writerow(
        [
            "Categoría",
            "Unidad",
            "Título",
            "Responsable unidad",
            "Avance unidad (%)",
            "Objetivo",
            "Tarea",
            "Prioridad",
            "Fecha de entrega",
            "Responsable actividad",
            "Entregable",
            "Avance actividad (%)",
            "Observación",
        ]
    )
    for u in poa["unidades"]:
        avance_u = round((u["avance"] or 0) * 100, 1)
        for a in u["actividades"]:
            avance_a = "" if a.get("avance") is None else round(a["avance"] * 100, 1)
            writer.writerow(
                [
                    u["categoria"],
                    u["nombre"],
                    u.get("titulo") or "",
                    u.get("responsable") or "",
                    avance_u,
                    a.get("objetivo") or "",
                    a.get("tarea") or "",
                    a.get("prioridad") or "",
                    a.get("fecha") or "",
                    a.get("responsable") or "",
                    a.get("entregable") or "",
                    avance_a,
                    a.get("observacion") or "",
                ]
            )
    return response


def descargar_json(request):
    poa = cargar_poa()
    payload = json.dumps(poa, ensure_ascii=False, indent=2)
    response = HttpResponse(payload, content_type="application/json; charset=utf-8")
    response["Content-Disposition"] = f'attachment; filename="poa_{poa["anio"]}.json"'
    return response


# ---------------------------------------------------------------------------
# Panel de administración: subir el Excel del POA y regenerar el JSON
# ---------------------------------------------------------------------------

import datetime
import shutil

from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import user_passes_test
from django.shortcuts import redirect

from .importer import ExcelPOAError, poa_desde_excel

es_admin = user_passes_test(lambda u: u.is_active and u.is_staff, login_url="/panel/login/")


def panel_login(request):
    if request.user.is_authenticated and request.user.is_staff:
        return redirect("panel")
    error = None
    if request.method == "POST":
        usuario = authenticate(
            request,
            username=request.POST.get("username", "").strip(),
            password=request.POST.get("password", ""),
        )
        if usuario and usuario.is_staff:
            login(request, usuario)
            return redirect("panel")
        error = "Usuario o contraseña incorrectos, o la cuenta no es administradora."
    return render(request, "dashboard/panel_login.html", {"error": error})


def panel_logout(request):
    logout(request)
    return redirect("panel_login")


def _historial():
    carpeta = settings.POA_DATA_FILE.parent / "respaldos"
    if not carpeta.exists():
        return []
    archivos = sorted(carpeta.glob("*.json"), reverse=True)
    return [
        {
            "nombre": a.name,
            "fecha": datetime.datetime.fromtimestamp(a.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
            "kb": round(a.stat().st_size / 1024, 1),
        }
        for a in archivos[:10]
    ]


@es_admin
def panel(request):
    poa = cargar_poa()
    contexto = {
        "anio": poa["anio"],
        "resumen": resumen(poa["unidades"]),
        "proyectos": len(poa.get("proyectos", [])),
        "historial": _historial(),
        "archivo": settings.POA_DATA_FILE.name,
    }

    if request.method == "POST":
        excel = request.FILES.get("excel")
        anio = request.POST.get("anio") or poa["anio"]
        if not excel:
            messages.error(request, "Selecciona un archivo de Excel (.xlsx).")
            return render(request, "dashboard/panel.html", contexto)
        if not excel.name.lower().endswith((".xlsx", ".xlsm")):
            messages.error(request, "El archivo debe ser .xlsx o .xlsm.")
            return render(request, "dashboard/panel.html", contexto)

        try:
            nuevo = poa_desde_excel(excel, int(anio))
        except ExcelPOAError as exc:
            messages.error(request, str(exc))
            return render(request, "dashboard/panel.html", contexto)
        except Exception as exc:  # noqa: BLE001
            messages.error(request, f"No se pudo procesar el Excel: {exc}")
            return render(request, "dashboard/panel.html", contexto)

        omitidas = nuevo.pop("_omitidas", [])

        # Respaldo del JSON anterior antes de sobrescribirlo
        destino = settings.POA_DATA_FILE
        if destino.exists():
            carpeta = destino.parent / "respaldos"
            carpeta.mkdir(parents=True, exist_ok=True)
            marca = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            shutil.copy2(destino, carpeta / f"{destino.stem}-{marca}.json")

        destino.parent.mkdir(parents=True, exist_ok=True)
        with open(destino, "w", encoding="utf-8") as fh:
            json.dump(nuevo, fh, ensure_ascii=False, indent=1)

        cargar_poa.cache_clear()

        actividades = sum(len(u["actividades"]) for u in nuevo["unidades"])
        messages.success(
            request,
            f"Datos actualizados: {len(nuevo['unidades'])} unidades, {actividades} actividades "
            f"y {len(nuevo['proyectos'])} proyectos.",
        )
        if omitidas:
            messages.warning(
                request,
                "Hojas omitidas (sin encabezados reconocibles): " + ", ".join(omitidas),
            )
        return redirect("panel")

    return render(request, "dashboard/panel.html", contexto)


@es_admin
def restaurar_respaldo(request, nombre):
    origen = settings.POA_DATA_FILE.parent / "respaldos" / nombre
    if not origen.exists() or origen.suffix != ".json":
        messages.error(request, "El respaldo indicado no existe.")
        return redirect("panel")
    shutil.copy2(origen, settings.POA_DATA_FILE)
    cargar_poa.cache_clear()
    messages.success(request, f"Se restauró el respaldo {nombre}.")
    return redirect("panel")
