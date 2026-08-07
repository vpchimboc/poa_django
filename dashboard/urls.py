from django.urls import path

from . import views

urlpatterns = [
    path("", views.index, name="index"),
    path("api/poa/", views.api_poa, name="api_poa"),
    path("descargar/csv/", views.descargar_csv, name="descargar_csv"),
    path("descargar/json/", views.descargar_json, name="descargar_json"),
    path("panel/", views.panel, name="panel"),
    path("panel/login/", views.panel_login, name="panel_login"),
    path("panel/salir/", views.panel_logout, name="panel_logout"),
    path("panel/restaurar/<str:nombre>/", views.restaurar_respaldo, name="restaurar_respaldo"),
]
