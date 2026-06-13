from pathlib import Path
import json
import random
import threading
import http.server
import socketserver
import socket
import webbrowser
import base64
import mimetypes
import time
import pandas as pd


class OSMExplorer:
    GIPHY_RANDOM = [
        "https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExeWR5ZmlxMjN0NjlpbzdvMmc4MnM1dXFpYmd3bjRwejlpeGc1NGo5MyZlcD12MV9naWZzX3NlYXJjaCZjdD1n/cKQksH9JmUus0/200.webp",
        "https://media1.giphy.com/media/v1.Y2lkPTc5MGI3NjExYzF3aTg2eDJyazZ3aG45ZnZicDBmNWlvZjNkaWpzMTJreGowdG14dyZlcD12MV9naWZzX3NlYXJjaCZjdD1n/26gJz0lJ04Y7O7rG4/giphy.gif",
        "https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExaWJxNnR0amprcDg4dHNuM3BqNWp6NjkxNmtkMGVyMnF1d3c3ZXA0eiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/3ohc1h5TqJN2TZGf5G/giphy.gif",
    ]

    def __init__(self, output_dir=None, data_dir=None):
        self.package_dir = Path(__file__).parent
        self.output_dir = Path(output_dir) if output_dir else self.package_dir.parent
        self.data_dir = Path(data_dir) if data_dir else (self.package_dir / "data")

    @staticmethod
    def _clean_columns(df):
        df = df.copy()
        df.columns = [str(c).strip().lower() for c in df.columns]
        return df

    @staticmethod
    def _cell(row, name, default=""):
        value = row.get(name, default)
        if pd.isna(value):
            return default
        text = str(value).strip()
        if text.lower() == "nan":
            return default
        return text

    @staticmethod
    def _num(value):
        if pd.isna(value):
            raise ValueError("coordenada vacía")
        text = str(value).strip().replace(",", ".")
        if not text:
            raise ValueError("coordenada vacía")
        return float(text)

    def _logo_src(self):
        """Busca logo_puce.png/jpg en la carpeta principal o en repo_openstreetmap."""
        candidates = [
            self.output_dir / "logo_puce.png",
            self.output_dir / "logo_puce.jpg",
            self.output_dir / "Logo_PUCESD.png",
            self.package_dir / "logo_puce.png",
            self.package_dir / "logo_puce.jpg",
            self.package_dir / "Logo_PUCESD.png",
        ]
        for path in candidates:
            if path.exists():
                mime = mimetypes.guess_type(str(path))[0] or "image/png"
                data = base64.b64encode(path.read_bytes()).decode("ascii")
                return f"data:{mime};base64,{data}"
        return ""

    @staticmethod
    def _format_excel_value(value, label=""):
        """Convierte cualquier celda del Excel en texto seguro para mostrar en el popup."""
        if pd.isna(value):
            return ""

        # Fechas reales leídas por pandas.
        if isinstance(value, pd.Timestamp):
            return value.strftime("%Y-%m-%d")

        # Fechas guardadas como número serial de Excel.
        label_text = str(label).lower()
        if "fecha" in label_text:
            try:
                number = float(str(value).replace(",", "."))
                if 25000 <= number <= 60000:
                    date = pd.to_datetime(number, unit="D", origin="1899-12-30")
                    return date.strftime("%Y-%m-%d")
            except Exception:
                pass

        # Evita mostrar 10.0 cuando en Excel se ve 10.
        try:
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
        except Exception:
            pass

        text = str(value).strip()
        if text.lower() == "nan":
            return ""
        return text

    def _load_from_xlsx(self):
        master_path = self.data_dir / "coordenadas.xlsx"
        if not master_path.exists():
            raise FileNotFoundError(f"No existe el archivo: {master_path}")

        # Se lee el Excel original para conservar los nombres de columnas tal como están escritos.
        raw_master = pd.read_excel(master_path, engine="openpyxl")
        display_names = {str(c).strip().lower(): str(c).strip() for c in raw_master.columns}

        master = self._clean_columns(raw_master)

        required = ["lat", "lon"]
        missing = [c for c in required if c not in master.columns]
        if missing:
            raise ValueError(f"Faltan columnas obligatorias en coordenadas.xlsx: {missing}")

        # name es recomendable, pero si no existe se usa codigo o Punto N.
        if "name" not in master.columns:
            print("Aviso: no existe columna 'name'. Se usará 'codigo' o 'Punto N'.")

        # Estas columnas sirven para ubicar/controlar el punto. Las demás se muestran automáticamente en el popup.
        hidden_cols = {"lat", "lon", "image", "data_file"}

        coords = []
        skipped = 0
        missing_data_files = set()
        for idx, row in master.iterrows():
            name = self._cell(row, "name")
            codigo = self._cell(row, "codigo")
            lat_raw = row.get("lat", "")
            lon_raw = row.get("lon", "")

            # Saltar filas completamente vacías.
            if not name and not codigo and pd.isna(lat_raw) and pd.isna(lon_raw):
                skipped += 1
                continue

            try:
                lat = self._num(lat_raw)
                lon = self._num(lon_raw)
            except Exception as e:
                raise ValueError(f"Error en fila {idx + 2} de coordenadas.xlsx: lat/lon inválidas. Detalle: {e}")

            link = self._cell(row, "link")
            if not link:
                link = self._cell(row, "image")

            c = {
                "name": name or codigo or f"Punto {idx + 1}",
                "codigo": codigo,
                "lat": lat,
                "lon": lon,
                "link": link,
                "image": link,
                "extra_fields": [],
                "info": [],
                "roads": [],
            }

            # Agrega automáticamente todas las columnas nuevas/eliminadas del Excel al popup.
            # Si mañana agregas "observación", "parroquia", "estado", etc., aparece sin tocar el código.
            for col in master.columns:
                if col in hidden_cols:
                    continue
                if col == "link":
                    # Link se muestra como botón/enlace al final del popup.
                    continue
                label = display_names.get(col, col)
                value = self._format_excel_value(row.get(col, ""), label)
                c["extra_fields"].append({"label": label, "value": value})

            data_file = self._cell(row, "data_file")
            if data_file:
                info_path = self.data_dir / data_file
                if info_path.exists():
                    df = pd.read_excel(info_path, engine="openpyxl")
                    df = self._clean_columns(df)

                    # Formato inventario vial.
                    if "via" in df.columns:
                        for _, ir in df.iterrows():
                            via = self._cell(ir, "via")
                            if not via:
                                continue
                            def fnum(v):
                                try:
                                    return float(str(v).replace(",", "."))
                                except Exception:
                                    return 0.0
                            def inum(v):
                                try:
                                    return int(float(str(v).replace(",", ".")))
                                except Exception:
                                    return 0
                            c["roads"].append({
                                "via": via,
                                "tipo": self._cell(ir, "tipo"),
                                "superficie": self._cell(ir, "superficie"),
                                "sentido": self._cell(ir, "sentido_unico"),
                                "carriles": self._cell(ir, "carriles"),
                                "vel_max": self._cell(ir, "velocidad_max"),
                                "longitud": fnum(ir.get("longitud_m", 0)),
                                "id_osm": inum(ir.get("id_osm", 0)),
                            })

                    # Formato resumen antiguo.
                    elif "categoría" in df.columns or "categoria" in df.columns:
                        cat_col = "categoría" if "categoría" in df.columns else "categoria"
                        for _, ir in df.iterrows():
                            c["info"].append({
                                "cat": self._cell(ir, cat_col),
                                "qty": self._cell(ir, "cantidad"),
                                "desc": self._cell(ir, "descripción") or self._cell(ir, "descripcion"),
                            })
                else:
                    missing_data_files.add(data_file)

            coords.append(c)

        if not coords:
            raise ValueError("coordenadas.xlsx no tiene puntos válidos.")

        print(f"Puntos cargados desde coordenadas.xlsx: {len(coords)}")
        print("Columnas detectadas:", ", ".join(display_names.get(c, c) for c in master.columns))
        if missing_data_files:
            print("Aviso: estos archivos data_file no están en la carpeta data:", ", ".join(sorted(missing_data_files)))
        if skipped:
            print(f"Filas vacías omitidas: {skipped}")
        return coords

    HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PUCE — Mapa de Puntos Críticos Zona 4 Loja</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>window.CESIUM_BASE_URL = 'https://unpkg.com/cesium@1.120.0/Build/Cesium/';</script>
<script src="https://unpkg.com/cesium@1.120.0/Build/Cesium/Cesium.js"></script>
<link href="https://unpkg.com/cesium@1.120.0/Build/Cesium/Widgets/widgets.css" rel="stylesheet" />
<style>
* { margin:0; padding:0; box-sizing:border-box; font-family:'Segoe UI', Tahoma, sans-serif; }
body { display:flex; height:100vh; overflow:hidden; background:#f0f2f5; }
#sidebar { width:400px; min-width:400px; background:#fff; color:#333; display:flex; flex-direction:column; border-right:1px solid #e0e0e0; z-index:1000; box-shadow:2px 0 8px rgba(0,0,0,0.06); }
#sidebar-header { padding:16px 20px; background:#fff; border-bottom:1px solid #e0e0e0; }
#brand { display:flex; justify-content:center; margin-bottom:12px; }
#brand-logo { display:flex; align-items:center; justify-content:center; gap:14px; }
#brand-logo img { width:58px; height:58px; object-fit:contain; display:$LOGO_DISPLAY; }
#brand-logo span { font-size:34px; font-weight:700; letter-spacing:5px; color:#444; }
#author { font-size:13px; color:#555; font-weight:500; margin-bottom:3px; text-align:left; }
#disclaimer { font-size:11px; color:#999; font-style:italic; text-align:left; }
#coord-list { flex:1; overflow-y:auto; padding:10px; }
.coord-item { display:flex; align-items:center; gap:10px; padding:10px 12px; margin-bottom:6px; background:#fff; border-radius:8px; cursor:pointer; transition:all .2s; border:1px solid #e8e8e8; box-shadow:0 1px 3px rgba(0,0,0,.04); }
.coord-item:hover { background:#fff8f0; border-color:#f57c00; box-shadow:0 2px 8px rgba(245,124,0,.12); }
.coord-item .dot { width:10px; height:10px; border-radius:50%; background:#f57c00; flex-shrink:0; }
.coord-item .info { flex:1; min-width:0; }
.coord-item .name { font-size:13px; font-weight:500; color:#1a1a2e; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.coord-item .coords { font-size:11px; color:#999; }
#sidebar-footer { padding:10px 16px; border-top:1px solid #e8e8e8; font-size:11px; color:#999; text-align:center; background:#fafafa; }
#sidebar-footer a { color:#f57c00; text-decoration:none; }
#btn-home { padding:12px 16px; border:none; background:#fff; color:#666; cursor:pointer; font-size:18px; font-weight:600; border-radius:8px; box-shadow:0 2px 12px rgba(0,0,0,.15); }
#btn-home:hover { background:#fff8f0; color:#f57c00; }
#btn-sidebar-home { display:none; width:calc(100% - 20px); margin:8px auto 0; padding:8px 12px; border:none; background:#f57c00; color:#fff; cursor:pointer; font-size:13px; font-weight:600; border-radius:8px; }
.view-3d-active #btn-sidebar-home { display:block; }
#btn-info { position:absolute; top:150px; right:20px; z-index:999; width:40px; height:40px; border:none; background:#fff; color:#888; cursor:pointer; font-size:16px; font-weight:700; border-radius:50%; box-shadow:0 2px 10px rgba(0,0,0,.1); display:flex; align-items:center; justify-content:center; }
#info-panel { position:absolute; top:0; right:-420px; width:400px; height:100%; z-index:998; background:#fff; box-shadow:-4px 0 30px rgba(0,0,0,.08); transition:right .35s ease; display:flex; flex-direction:column; overflow-y:auto; }
#info-panel.open { right:0; }
#info-panel-header { padding:28px 28px 16px; border-bottom:1px solid #eee; display:flex; align-items:center; justify-content:space-between; }
#info-panel-header h3 { font-size:15px; color:#222; font-weight:600; margin:0; letter-spacing:2px; text-transform:uppercase; }
#info-panel-close { background:none; border:none; font-size:18px; color:#bbb; cursor:pointer; }
#info-panel-body { padding:20px 28px 28px; font-size:12.5px; color:#555; line-height:1.8; }
#info-panel-body h4 { font-size:10.5px; font-weight:700; color:#999; margin:22px 0 6px; letter-spacing:1.5px; text-transform:uppercase; border-bottom:1px solid #f0f0f0; padding-bottom:4px; }
#info-panel-body .disclaimer { background:#fafafa; border-left:2px solid #ccc; padding:14px 16px; font-size:11.5px; color:#777; margin-top:20px; line-height:1.7; }
#view-container { flex:1; position:relative; }
.view { position:absolute; top:0; left:0; width:100%; height:100%; display:none; }
.view.active { display:block; }
#view-toggle { position:absolute; top:80px; right:20px; z-index:1000; display:flex; gap:4px; transition:right .35s ease; }
#view-toggle.shifted { right:420px; }
#view-toggle button { padding:12px 22px; border:none; background:#fff; color:#666; cursor:pointer; font-size:18px; font-weight:600; border-radius:8px; box-shadow:0 2px 12px rgba(0,0,0,.15); }
#view-toggle button.active { background:#f57c00; color:#fff; }
.info-table { width:100%; border-collapse:collapse; margin-top:8px; font-size:11px; }
.info-table th { background:#f57c00; color:#fff; padding:4px 6px; text-align:left; font-weight:600; white-space:nowrap; }
.info-table td { padding:3px 6px; border-bottom:1px solid #e8e8e8; color:#444; white-space:nowrap; }
.info-table tr:nth-child(even) td { background:#fffaf2; }
.popup-content { max-height:340px; overflow-y:auto; }
.popup-field { font-size:11px; color:#444; margin-top:8px; line-height:1.6; }
#popup-3d { display:none; position:absolute; bottom:30px; left:50%; transform:translateX(-50%); z-index:1002; background:#fff; border-radius:12px; box-shadow:0 4px 24px rgba(0,0,0,.2); max-width:480px; padding:8px; overflow:hidden; }
#music-btn { background:none; border:none; cursor:pointer; font-size:14px; color:#999; padding:2px 4px; border-radius:4px; }
#menu-btn { display:none; position:fixed; top:12px; left:12px; z-index:3000; width:40px; height:40px; border:none; background:#fff; color:#333; cursor:pointer; font-size:22px; border-radius:50%; box-shadow:0 2px 10px rgba(0,0,0,.15); align-items:center; justify-content:center; }
#sidebar-overlay { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,.4); z-index:1999; }
#sidebar-overlay.open { display:block; }
@media (max-width:768px) {
body { flex-direction:column; }
#sidebar { position:fixed; top:0; left:0; width:85vw; max-width:360px; height:100%; z-index:2000; transform:translateX(-100%); transition:transform .3s ease; box-shadow:4px 0 20px rgba(0,0,0,.15); }
#sidebar.open { transform:translateX(0); }
#menu-btn { display:flex; }
#view-toggle { top:12px; right:12px; }
#view-toggle button { padding:8px 14px; font-size:15px; }
#btn-info { top:70px; right:12px; width:36px; height:36px; }
#info-panel { width:100%; right:-100%; }
#view-toggle.shifted { right:12px; }
#popup-3d { max-width:90vw; bottom:12px; }
#brand-logo img { width:46px; height:46px; }
#brand-logo span { font-size:28px; }
}
</style>
</head>
<body>
<div id="sidebar-overlay" onclick="toggleSidebar()"></div>
<button id="menu-btn" onclick="toggleSidebar()">☰</button>
<div id="sidebar" class="view-3d-active">
<div id="sidebar-header">
    <div id="brand">
        <div id="brand-logo">
            <img src="$PUCE_LOGO" alt="Logo PUCE">
            <span>PUCE</span>
        </div>
    </div>
    <div id="author">Samantha Mantilla y Mateo Valverde</div>
    <div id="disclaimer">Mapa de Puntos Críticos Zona 4 Loja</div>
</div>
    <button id="btn-sidebar-home" onclick="homeView()">🏠 Vista Global</button>
    <div id="coord-list"></div>
    <div id="sidebar-footer">
        <div style="display:flex;align-items:center;justify-content:center;gap:6px;flex-wrap:wrap;">
            <span>OpenStreetMap 2D / 3D &bull; Datos &copy; <a href="https://openstreetmap.org/copyright" target="_blank">OpenStreetMap</a></span>
            <span style="color:#bbb;">&bull;</span>
            <span>🎵 <a href="https://www.scottbuckley.com.au/library/filaments/" target="_blank" style="color:#999;text-decoration:none;">Scott Buckley</a></span>
            <button id="music-btn" onclick="toggleMusic()" title="Música de fondo">🔇</button>
        </div>
        <div style="margin-top:6px; color:#999;">Código desarrollado por: <strong>Ing. Carlos Celi</strong></div>
    </div>
</div>
<div id="view-container">
    <div id="view-toggle">
        <button id="btn-2d" onclick="switchView('2d')">🗺️ 2D</button>
        <button id="btn-3d" class="active" onclick="switchView('3d')">🌐 3D</button>
        <button id="btn-home" onclick="homeView()">🏠</button>
    </div>
    <div id="map-2d" class="view"></div>
    <div id="map-3d" class="view active"></div>
    <div id="popup-3d"></div>
    <button id="btn-info" onclick="toggleInfo()">ℹ️</button>
    <div id="info-panel">
        <div id="info-panel-header">
            <h3>Información</h3>
            <button id="info-panel-close" onclick="toggleInfo()">✕</button>
        </div>
        <div id="info-panel-body">
            <div style="text-align:center;margin-bottom:16px;">
                <img src="$PUCE_LOGO" style="height:58px;width:auto;display:$LOGO_DISPLAY;" alt="Logo PUCE">
            </div>
            <h4>Objetivo</h4>
            <p>Visualizar los puntos críticos viales de la Zona 4 del cantón Loja mediante un mapa interactivo 2D y 3D.</p>
            <h4>Autores</h4>
            <p><strong>Samantha Mantilla y Mateo Valverde</strong></p>
            <h4>Uso</h4>
            <p>Proyecto académico para la visualización de puntos críticos, coordenadas y atributos viales levantados en campo.</p>
            <h4>Datos visualizados</h4>
            <p>Código, coordenadas, capa de rodadura, severidad, drenaje y enlaces asociados a cada punto.</p>
            <div class="disclaimer"><strong>PUCE</strong><br>Mapa de Puntos Críticos Zona 4 Loja.</div>
        </div>
    </div>
</div>
<script>
var COORDS_STORAGE_KEY = 'osm_coords';
var COORDS_HASH_KEY = 'osm_coords_hash';
var APP_VERSION = '$APP_VERSION';
var VERSION_URL = 'osm_explorer_version.json';
var defaultCoords = $DEFAULT_COORDS;
var savedCoords = [];
var markers = [];
var circles = [];
var currentView = '3d';
if (!defaultCoords || defaultCoords.length === 0) {
    document.body.innerHTML = '<div style="padding:20px;font-family:Arial;">No hay puntos válidos en coordenadas.xlsx</div>';
    throw new Error('defaultCoords vacío');
}
var map2d = L.map('map-2d', { zoomControl:true }).setView([defaultCoords[0].lat, defaultCoords[0].lon], 15);
L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', { attribution:'&copy; <a href="https://openstreetmap.org/copyright">OpenStreetMap</a>', maxZoom:19 }).addTo(map2d);
var viewer = null;
var viewerReady = false;
var popup3D = document.getElementById('popup-3d');
var PIN_SVG = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24"><circle cx="12" cy="12" r="9" fill="#f57c00" stroke="#fff" stroke-width="2.5"/></svg>');
function initViewer() {
    if (viewerReady) return;
    viewerReady = true;
    if (typeof Cesium === 'undefined') { console.error('Cesium no cargó. Revisa internet.'); return; }
    try {
        if (typeof Cesium.buildModuleUrl !== 'undefined' && typeof Cesium.buildModuleUrl.setBaseUrl === 'function') { Cesium.buildModuleUrl.setBaseUrl('https://unpkg.com/cesium@1.120.0/Build/Cesium/'); }
        viewer = new Cesium.Viewer('map-3d', { terrainProvider:new Cesium.EllipsoidTerrainProvider(), baseLayerPicker:false, geocoder:false, homeButton:false, sceneModePicker:false, timeline:false, animation:false, navigationHelpButton:false, infoBox:false, fullscreenButton:false, selectionIndicator:false });
        viewer.scene.skyAtmosphere.show = true;
        viewer.scene.skyAtmosphere.brightnessShift = 2.0;
        viewer.scene.skyAtmosphere.saturationShift = 0.8;
        viewer.scene.globe.showGroundAtmosphere = true;
        viewer.scene.globe.enableLighting = true;
        viewer.scene.globe.dynamicAtmosphereLighting = true;
        viewer.scene.globe.dynamicAtmosphereLightingFromSun = true;
        viewer.scene.globe.atmosphereLightIntensity = 6.0;
        viewer.scene.globe.atmosphereBrightnessShift = 1.2;
        viewer.scene.fog.enabled = true;
        viewer.scene.fog.density = 3.0e-5;
        viewer.scene.postProcessStages.fxaa.enabled = true;
        viewer.imageryLayers.removeAll();
        var nightLayer = viewer.imageryLayers.addImageryProvider(new Cesium.UrlTemplateImageryProvider({ url:'https://gibs.earthdata.nasa.gov/wmts/epsg3857/best/VIIRS_CityLights_2012/default/GoogleMapsCompatible_Level8/{z}/{y}/{x}.jpeg', tilingScheme:new Cesium.WebMercatorTilingScheme(), maximumLevel:8, credit:'NASA VIIRS City Lights 2012' }));
        var satLayer = viewer.imageryLayers.addImageryProvider(new Cesium.UrlTemplateImageryProvider({ url:'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', credit:'Esri World Imagery' }));
        var osmLayer = viewer.imageryLayers.addImageryProvider(new Cesium.OpenStreetMapImageryProvider({ url:'https://tile.openstreetmap.org/' }));
        satLayer.nightAlpha = 0.0;
        osmLayer.nightAlpha = 0.0;
        function blendLayers() {
            var h = Cesium.Cartographic.fromCartesian(viewer.camera.position).height;
            var t;
            if (h > 500000) { t = 0.75; }
            else if (h < 100000) { t = 0.0; }
            else { t = 0.75 * (1 - (500000 - h) / 400000); }
            satLayer.alpha = t;
            osmLayer.alpha = 1 - t;
            nightLayer.alpha = 1;
            var close = h < 150000;
            viewer.scene.skyAtmosphere.show = !close;
            viewer.scene.globe.showGroundAtmosphere = !close;
            viewer.scene.globe.enableLighting = !close;
            viewer.scene.fog.enabled = !close;
        }
        viewer.camera.moveEnd.addEventListener(blendLayers);
        blendLayers();
        viewer.camera.flyTo({ destination: Cesium.Cartesian3.fromDegrees(defaultCoords[0].lon, defaultCoords[0].lat, 8000000), duration:0 });
        viewer.screenSpaceEventHandler.setInputAction(function(click) { var picked = viewer.scene.pick(click.position); if (Cesium.defined(picked) && Cesium.defined(picked.id) && picked.id.coordIndex !== undefined) { goToCoord(savedCoords[picked.id.coordIndex]); } }, Cesium.ScreenSpaceEventType.LEFT_CLICK);
        updateCesiumMarkers();
    } catch(e) { console.error('[Cesium] Error al inicializar:', e.message, e.stack); viewer = null; viewerReady = false; }
}
function updateCesiumMarkers() {
    if (!viewer) return;
    viewer.entities.removeAll();
    var INFLUENCIA = 200;
    savedCoords.forEach(function(c, i) {
        var pos = Cesium.Cartesian3.fromDegrees(c.lon, c.lat);
        viewer.entities.add({ position:pos, ellipse:{ semiMinorAxis:INFLUENCIA, semiMajorAxis:INFLUENCIA, material:Cesium.Color.fromCssColorString('#1976d2').withAlpha(.25), outline:true, outlineColor:Cesium.Color.BLACK.withAlpha(.7), outlineWidth:2 } });
        viewer.entities.add({ position:pos, billboard:{ image:PIN_SVG, width:24, height:24, verticalOrigin:Cesium.VerticalOrigin.CENTER }, name:c.name, coordIndex:i });
    });
}
function switchView(view) {
    currentView = view;
    document.querySelectorAll('.view').forEach(function(v){ v.classList.remove('active'); });
    document.querySelectorAll('#view-toggle button').forEach(function(b){ b.classList.remove('active'); });
    document.getElementById('map-' + view).classList.add('active');
    document.getElementById('btn-' + view).classList.add('active');
    document.getElementById('sidebar').classList.toggle('view-3d-active', view === '3d');
    if (view === '3d') { initViewer(); if (viewer) viewer.resize(); }
    if (view === '2d') { map2d.invalidateSize(); }
    popup3D.style.display = 'none';
}
function toggleSidebar() { document.getElementById('sidebar').classList.toggle('open'); document.getElementById('sidebar-overlay').classList.toggle('open'); }
function homeView() { if (viewer && currentView === '3d') { viewer.camera.flyTo({ destination:Cesium.Cartesian3.fromDegrees(defaultCoords[0].lon, defaultCoords[0].lat, 8000000), duration:1.5 }); } else { map2d.setView([defaultCoords[0].lat, defaultCoords[0].lon], 11); } }
function coordsHash() { return JSON.stringify(defaultCoords); }
function loadCoords() {
    var stored = localStorage.getItem(COORDS_STORAGE_KEY);
    var storedHash = localStorage.getItem(COORDS_HASH_KEY);
    var currentHash = coordsHash();
    if (stored && storedHash === currentHash) { savedCoords = JSON.parse(stored); }
    else { savedCoords = JSON.parse(JSON.stringify(defaultCoords)); localStorage.setItem(COORDS_STORAGE_KEY, JSON.stringify(savedCoords)); localStorage.setItem(COORDS_HASH_KEY, currentHash); }
    renderCoords(); updateMarkers();
}
function saveCoords() { localStorage.setItem(COORDS_STORAGE_KEY, JSON.stringify(savedCoords)); }
function checkForUpdates() {
    fetch(VERSION_URL + '?t=' + Date.now(), { cache:'no-store' })
        .then(function(r){ return r.ok ? r.json() : null; })
        .then(function(data){
            if (data && data.version && data.version !== APP_VERSION) {
                localStorage.removeItem(COORDS_STORAGE_KEY);
                localStorage.removeItem(COORDS_HASH_KEY);
                location.reload();
            }
        })
        .catch(function(){});
}
setInterval(checkForUpdates, 3000);
function buildPopupHtml(c) {
    var html = '<div class="popup-content">';
    html += '<div style="text-align:center;margin-bottom:6px;"><b style="color:#f57c00;font-size:14px;">' + escapeHtml(c.name || '') + '</b></div>';
    html += '<div style="text-align:center;color:#888;font-size:11px;margin-bottom:4px;">' + Number(c.lat).toFixed(5) + ', ' + Number(c.lon).toFixed(5) + '</div>';
    html += '<div class="popup-field">';

    if (c.extra_fields && c.extra_fields.length > 0) {
        for (var ef=0; ef<c.extra_fields.length; ef++) {
            var field = c.extra_fields[ef];
            var value = field.value;
            if (value === null || value === undefined || String(value).trim() === '') { value = 'Sin dato'; }
            html += '<b>' + escapeHtml(field.label || '') + ':</b> ' + escapeHtml(value) + '<br>';
        }
    } else {
        // Compatibilidad con mapas antiguos.
        html += '<b>Código:</b> ' + escapeHtml(c.codigo || 'Sin dato') + '<br>';
        html += '<b>Capa de rodadura:</b> ' + escapeHtml(c.capa_rod || 'Sin dato') + '<br>';
        html += '<b>Severidad:</b> ' + escapeHtml(c.severidad || 'Sin dato') + '<br>';
        html += '<b>Drenaje:</b> ' + escapeHtml(c.drenaje || 'Sin dato') + '<br>';
    }

    if (c.link && c.link !== 'nan') { html += '<b>Link:</b> <a href="' + encodeURI(c.link) + '" target="_blank">Abrir enlace</a>'; }
    html += '</div>';
    if (c.roads && c.roads.length > 0) {
        html += '<div style="margin-top:8px;font-size:11px;font-weight:700;color:#f57c00;">Inventario Vial (200m)</div>';
        html += '<table class="info-table"><tr><th>Vía</th><th>Tipo</th><th>Superficie</th><th>Carriles</th><th>Vel</th><th>Long</th></tr>';
        for (var i=0; i<c.roads.length; i++) { var r = c.roads[i]; html += '<tr><td>' + escapeHtml(r.via) + '</td><td>' + escapeHtml(r.tipo) + '</td><td>' + escapeHtml(r.superficie) + '</td><td>' + escapeHtml(r.carriles) + '</td><td>' + escapeHtml(r.vel_max) + '</td><td>' + Number(r.longitud || 0).toFixed(0) + 'm</td></tr>'; }
        html += '</table>';
    } else if (c.info && c.info.length > 0) {
        html += '<table class="info-table"><tr><th>Categoría</th><th>Cant</th><th>Descripción</th></tr>';
        for (var j=0; j<c.info.length; j++) { var q = c.info[j]; html += '<tr><td>' + escapeHtml(q.cat) + '</td><td>' + escapeHtml(q.qty) + '</td><td>' + escapeHtml(q.desc) + '</td></tr>'; }
        html += '</table>';
    }
    html += '</div>';
    return html;
}
function escapeHtml(v) { return String(v == null ? '' : v).replace(/[&<>'"]/g, function(ch){ return {'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]; }); }
function goToCoord(c) {
    if (window.innerWidth <= 768) { toggleSidebar(); }
    map2d.setView([c.lat, c.lon], 16);
    if (viewerReady && viewer) { viewer.camera.flyTo({ destination:Cesium.Cartesian3.fromDegrees(c.lon, c.lat, 1000), duration:1.5, orientation:{ heading:Cesium.Math.toRadians(0), pitch:Cesium.Math.toRadians(-90), roll:0 } }); }
    var popupHtml = buildPopupHtml(c);
    L.popup({ maxWidth:420, className:'custom-popup' }).setLatLng([c.lat, c.lon]).setContent(popupHtml).openOn(map2d);
    if (currentView === '3d') showPopup3D(popupHtml);
}
function showPopup3D(html) {
    popup3D.innerHTML = '<button onclick="this.parentElement.style.display=\'none\'" style="position:absolute;top:6px;right:8px;border:none;background:none;font-size:18px;cursor:pointer;color:#999;z-index:10;">×</button>' + html;
    popup3D.style.display = 'block';
}
function renderCoords() {
    var list = document.getElementById('coord-list'); list.innerHTML = '';
    savedCoords.forEach(function(c, i) {
        var item = document.createElement('div'); item.className = 'coord-item';
        item.innerHTML = '<div class="dot"></div><div class="info"><div class="name">' + escapeHtml(c.name || '') + '</div><div class="coords">' + Number(c.lat).toFixed(5) + ', ' + Number(c.lon).toFixed(5) + '</div></div>';
        item.addEventListener('click', function(){ goToCoord(c); }); list.appendChild(item);
    });
}
function updateMarkers() {
    markers.forEach(function(m){ map2d.removeLayer(m); }); circles.forEach(function(c){ map2d.removeLayer(c); }); markers = []; circles = [];
    savedCoords.forEach(function(c) {
        var popupHtml = buildPopupHtml(c);
        var m = L.marker([c.lat, c.lon]).addTo(map2d).bindPopup(popupHtml, { maxWidth:420, className:'custom-popup' }); markers.push(m);
        var circle = L.circle([c.lat, c.lon], { radius:200, color:'#1565c0', fillColor:'#1976d2', fillOpacity:.25, weight:2.5, opacity:.9, dashArray:'8, 5' }).addTo(map2d); circles.push(circle);
    });
    if (viewerReady) updateCesiumMarkers();
}
loadCoords(); initViewer(); if (viewer) viewer.resize();
var audio = new Audio('https://www.scottbuckley.com.au/library/wp-content/uploads/2019/10/sb_filaments.mp3'); audio.loop = true; audio.volume = 0.06; var musicPlaying = false;
function toggleInfo() { document.getElementById('info-panel').classList.toggle('open'); document.getElementById('view-toggle').classList.toggle('shifted'); }
function toggleMusic() { if (musicPlaying) { audio.pause(); document.getElementById('music-btn').textContent = '🔇'; } else { audio.play().catch(function(){}); document.getElementById('music-btn').textContent = '🔊'; } musicPlaying = !musicPlaying; }
audio.play().then(function(){ musicPlaying = true; document.getElementById('music-btn').textContent = '🔊'; }).catch(function(){ document.addEventListener('click', function firstClick(){ audio.play().then(function(){ musicPlaying = true; document.getElementById('music-btn').textContent = '🔊'; }).catch(function(){}); document.removeEventListener('click', firstClick); }, { once:true }); });
</script>
</body>
</html>'''

    def _coords_to_js(self, coords):
        return json.dumps(coords, ensure_ascii=False, indent=2)

    def _data_exists(self):
        return (self.data_dir / "coordenadas.xlsx").exists()

    def _version_path(self):
        return self.output_dir / "osm_explorer_version.json"

    def _write_version(self, version):
        self._version_path().write_text(json.dumps({"version": version}, ensure_ascii=False), encoding="utf-8")

    def _watched_files(self):
        """Archivos que se vigilan para actualizar el mapa automáticamente."""
        files = []
        master_path = self.data_dir / "coordenadas.xlsx"
        if master_path.exists():
            files.append(master_path)
            try:
                df = pd.read_excel(master_path, engine="openpyxl")
                df = self._clean_columns(df)
                if "data_file" in df.columns:
                    for value in df["data_file"].dropna():
                        name = str(value).strip()
                        if name:
                            extra = self.data_dir / name
                            if extra.exists():
                                files.append(extra)
            except Exception:
                pass
        return files

    def _snapshot_mtimes(self):
        return {str(p): p.stat().st_mtime_ns for p in self._watched_files() if p.exists()}

    def _watch_excel_changes(self, filename="osm_explorer.html", interval=2):
        """Regenera el HTML cuando cambia coordenadas.xlsx o un archivo data_file."""
        previous = self._snapshot_mtimes()
        print("Vigilando cambios en data/coordenadas.xlsx...")
        while True:
            time.sleep(interval)
            current = self._snapshot_mtimes()
            if current != previous:
                previous = current
                try:
                    print("Cambio detectado en el Excel. Actualizando mapa...")
                    self.generate_html(filename)
                    print("Mapa regenerado. La página se recargará sola.")
                except Exception as e:
                    print(f"No se pudo actualizar el mapa: {e}")

    def generate_html(self, filename="osm_explorer.html"):
        version = str(int(time.time() * 1000))
        coords = self._load_from_xlsx() if self._data_exists() else [
            {"name":"Vilcabamba", "codigo":"", "lat":-4.30822, "lon":-79.24635, "link":"", "image":"", "capa_rod":"", "severidad":"", "drenaje":"", "info":[], "roads":[]}
        ]
        for c in coords:
            if not c.get("image"):
                c["image"] = c.get("link", "") or random.choice(self.GIPHY_RANDOM)
        logo_src = self._logo_src()
        html = (self.HTML_TEMPLATE
                .replace("$PUCE_LOGO", logo_src)
                .replace("$LOGO_DISPLAY", "inline-block" if logo_src else "none")
                .replace("$APP_VERSION", version)
                .replace("$DEFAULT_COORDS", self._coords_to_js(coords)))
        path = self.output_dir / filename
        path.write_text(html, encoding="utf-8")
        self._write_version(version)
        print(f"Mapa actualizado en: {path}")
        return path

    def open(self, filename="osm_explorer.html", port=8765):
        path = self.generate_html(filename)
        server_dir = str(self.output_dir)

        class ReusableTCPServer(socketserver.TCPServer):
            allow_reuse_address = True

        class Handler(http.server.SimpleHTTPRequestHandler):
            def end_headers(self):
                self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
                self.send_header("Pragma", "no-cache")
                self.send_header("Expires", "0")
                super().end_headers()

            def __init__(self, *args, **kwargs):
                super().__init__(*args, directory=server_dir, **kwargs)

        # Buscar puerto libre si 8765 está ocupado.
        selected_port = port
        for p in [port, 8000, 8001, 8002, 8010]:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            available = sock.connect_ex(("127.0.0.1", p)) != 0
            sock.close()
            if available:
                selected_port = p
                break

        url = f"http://127.0.0.1:{selected_port}/{filename}"
        print(f"Abriendo mapa en: {url}")
        print("Deja esta terminal abierta. Para cerrar el servidor presiona Ctrl + C.")

        watcher = threading.Thread(target=self._watch_excel_changes, args=(filename,), daemon=True)
        watcher.start()

        with ReusableTCPServer(("127.0.0.1", selected_port), Handler) as httpd:
            threading.Timer(1.0, lambda: webbrowser.open(url)).start()
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("Servidor cerrado.")

        return path


if __name__ == "__main__":
    app = OSMExplorer()
    app.open()
