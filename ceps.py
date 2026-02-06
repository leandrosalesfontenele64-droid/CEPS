import os
import pandas as pd
import folium
from folium.plugins import Draw, FastMarkerCluster
import webbrowser
import json

# ===== CONFIGURAÇÃO DE CAMINHOS =====
PASTA_EXCEL = r"C:\Users\JT-120\OneDrive - Speed Rabbit Express Ltda\Área de Trabalho\ceps brasil - feishu"
ARQUIVO_MAPA = "mapa_bairros_premium.html"


def gerar_mapa():
    # 1. Localizar arquivo Excel
    if not os.path.exists(PASTA_EXCEL):
        print(f"❌ Erro: Pasta não encontrada: {PASTA_EXCEL}")
        return

    arquivos = [os.path.join(PASTA_EXCEL, f) for f in os.listdir(PASTA_EXCEL) if f.lower().endswith(".xlsx")]
    if not arquivos:
        print(f"❌ Nenhum arquivo Excel encontrado em: {PASTA_EXCEL}")
        return

    arquivo_excel = arquivos[0]
    print(f"📥 Lendo arquivo: {arquivo_excel}")

    # 2. Carregar e limpar dados
    try:
        df = pd.read_excel(arquivo_excel)
    except Exception as e:
        print(f"❌ Erro ao ler o Excel: {e}")
        return

    df = df.dropna(subset=["latitude", "longitude"])
    df["latitude"] = df["latitude"].astype(float)
    df["longitude"] = df["longitude"].astype(float)

    # Padronizar colunas de filtro
    col_estado = next((c for c in df.columns if c.lower() in ['estado', 'uf']), 'estado')
    col_cidade = next((c for c in df.columns if c.lower() in ['cidade', 'municipio', 'localidade', 'city']), 'cidade')

    if col_estado not in df.columns: df[col_estado] = 'Não Informado'
    if col_cidade not in df.columns: df[col_cidade] = 'Não Informado'

    df = df.rename(columns={col_estado: 'estado_filtro', col_cidade: 'cidade_filtro'})

    # 3. Criar o mapa base
    mapa = folium.Map(location=[-14.2350, -51.9253], zoom_start=4, tiles="CartoDB positron")

    # 4. Adicionar ferramentas de desenho (Leaflet.Draw)
    draw = Draw(
        export=False,
        position='topleft',
        draw_options={
            'polyline': False,
            'rectangle': True,
            'circle': True,
            'polygon': True,
            'marker': False,
            'circlemarker': False
        },
        edit_options={'edit': True, 'remove': True}
    )
    draw.add_to(mapa)

    # 5. Preparar dados
    data_json = df.to_json(orient='records')
    estados = sorted(df['estado_filtro'].unique().tolist())
    cidades = sorted(df['cidade_filtro'].unique().tolist())

    # 6. Injetar Lógica Premium (Cores, Ícones e Exportação)
    custom_ui = f"""
    <style>
        #filter-panel {{
            position: fixed; top: 10px; right: 10px; z-index: 1000;
            background: white; padding: 15px; border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15); width: 260px;
            max-height: 85vh; overflow-y: auto; font-family: 'Segoe UI', sans-serif;
        }}
        .section-title {{
            font-weight: bold; margin: 12px 0 5px 0; padding-bottom: 5px;
            border-bottom: 2px solid #f8f9fa; display: flex; justify-content: space-between;
        }}
        .filter-list {{ max-height: 140px; overflow-y: auto; border: 1px solid #f1f1f1; padding: 5px; margin-bottom: 10px; border-radius: 6px; }}
        .filter-item {{ display: flex; align-items: center; gap: 8px; font-size: 0.85em; padding: 4px 0; }}
        .btn-small {{ background: #f1f3f5; border: none; padding: 3px 8px; border-radius: 4px; font-size: 10px; cursor: pointer; }}
        .btn-small:hover {{ background: #e9ecef; }}

        #color-picker-panel {{
            position: fixed; top: 10px; left: 50px; z-index: 1000;
            background: white; padding: 10px; border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1); display: flex; align-items: center; gap: 10px;
            font-family: 'Segoe UI', sans-serif; font-size: 13px;
        }}

        .export-btn {{
            position: fixed; bottom: 25px; right: 25px; z-index: 1000;
            padding: 16px 30px; background: linear-gradient(135deg, #28a745 0%, #218838 100%);
            color: white; border: none; border-radius: 50px; font-weight: bold;
            cursor: pointer; box-shadow: 0 6px 20px rgba(40, 167, 69, 0.4);
            transition: all 0.3s; font-size: 15px;
        }}
        .export-btn:hover {{ transform: translateY(-3px); box-shadow: 0 8px 25px rgba(40, 167, 69, 0.5); }}

        /* Estilo dos Clusters */
        .marker-cluster-small {{ background-color: rgba(181, 226, 140, 0.6); }}
        .marker-cluster-small div {{ background-color: rgba(110, 204, 57, 0.6); }}
        .marker-cluster-medium {{ background-color: rgba(241, 211, 87, 0.6); }}
        .marker-cluster-medium div {{ background-color: rgba(240, 194, 12, 0.6); }}
    </style>

    <div id="color-picker-panel">
        <b>Cor do Desenho:</b>
        <input type="color" id="draw-color" value="#3388ff" onchange="updateDrawStyle()">
    </div>

    <div id="filter-panel">
        <div style="font-weight: bold; text-align: center; margin-bottom: 12px; color: #2c3e50; font-size: 16px;">Painel de Controle</div>

        <div class="section-title">Estados <div><button class="btn-small" onclick="toggleGroup('state-check', true)">Todos</button> <button class="btn-small" onclick="toggleGroup('state-check', false)">Nenhum</button></div></div>
        <div class="filter-list">
            {"".join([f'<div class="filter-item"><input type="checkbox" class="state-check" value="{e}" checked onchange="updateMap()">{e}</div>' for e in estados])}
        </div>

        <div class="section-title">Cidades <div><button class="btn-small" onclick="toggleGroup('city-check', true)">Todas</button> <button class="btn-small" onclick="toggleGroup('city-check', false)">Nenhuma</button></div></div>
        <div class="filter-list" style="max-height: 220px;">
            {"".join([f'<div class="filter-item"><input type="checkbox" class="city-check" value="{c}" checked onchange="updateMap()">{c}</div>' for c in cidades])}
        </div>
    </div>
    """

    script_js = f"""
    {custom_ui}
    <script src="https://cdn.jsdelivr.net/npm/@turf/turf@6/turf.min.js"></script>
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.4.1/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.4.1/dist/MarkerCluster.Default.css" />
    <script src="https://unpkg.com/leaflet.markercluster@1.4.1/dist/leaflet.markercluster.js"></script>

    <script>
    var points_data = {data_json};
    var current_cluster = null;
    var mapObject = null;
    var drawnItems = new L.FeatureGroup();

    function initMap() {{
        for (var key in window) {{
            if (key.startsWith('map_') && window[key] instanceof L.Map) {{
                mapObject = window[key];
                break;
            }}
        }}

        if (mapObject) {{
            mapObject.addLayer(drawnItems);

            // Capturar desenhos e aplicar a cor selecionada
            mapObject.on(L.Draw.Event.CREATED, function (event) {{
                var layer = event.layer;
                var color = document.getElementById('draw-color').value;

                if (layer.setStyle) {{
                    layer.setStyle({{ color: color, fillColor: color, fillOpacity: 0.4 }});
                }}
                layer.options.color_name = color; // Salvar a cor no objeto
                drawnItems.addLayer(layer);
            }});

            updateMap();
        }} else {{
            setTimeout(initMap, 100);
        }}
    }}

    window.updateDrawStyle = function() {{
        // Atualiza a cor para os próximos desenhos
        var color = document.getElementById('draw-color').value;
        // O Leaflet Draw não tem uma API simples para mudar a cor global em tempo real, 
        // mas o evento CREATED acima resolve para os novos.
    }};

    window.toggleGroup = function(className, val) {{
        document.querySelectorAll('.' + className).forEach(el => el.checked = val);
        updateMap();
    }};

    window.updateMap = function() {{
        if (!mapObject) return;
        if (current_cluster) mapObject.removeLayer(current_cluster);

        var selStates = Array.from(document.querySelectorAll('.state-check:checked')).map(el => el.value);
        var selCities = Array.from(document.querySelectorAll('.city-check:checked')).map(el => el.value);

        var filtered = points_data.filter(p => selStates.includes(p.estado_filtro) && selCities.includes(p.cidade_filtro));

        if (filtered.length > 0) {{
            current_cluster = L.markerClusterGroup({{
                maxClusterRadius: 50,
                spiderfyOnMaxZoom: true,
                showCoverageOnHover: false,
                zoomToBoundsOnClick: true
            }});

            filtered.forEach(p => {{
                var m = L.marker([p.latitude, p.longitude]);
                var popup = '<div style="min-width: 150px; font-family: sans-serif;">';
                popup += `<h4 style="margin:0 0 5px 0; color:#007bff;">${{p.bairro || 'Informações'}}</h4>`;
                for (var key in p) {{
                    if (!['estado_filtro', 'cidade_filtro', 'latitude', 'longitude'].includes(key)) {{
                        popup += `<b>${{key}}:</b> ${{p[key]}}<br>`;
                    }}
                }}
                popup += '</div>';
                m.bindPopup(popup);
                current_cluster.addLayer(m);
            }});
            mapObject.addLayer(current_cluster);
        }}
    }};

    document.addEventListener('DOMContentLoaded', function() {{
        initMap();

        var btn = document.createElement('button');
        btn.className = 'export-btn';
        btn.innerHTML = '🚀 Exportar Dados com Cores';
        document.body.appendChild(btn);

        btn.onclick = function() {{
            var all_exported_data = [];

            if (drawnItems.getLayers().length === 0) {{
                alert('Desenhe pelo menos um polígono ou círculo para exportar!');
                return;
            }}

            drawnItems.eachLayer(function(layer) {{
                var color = layer.options.color_name || '#3388ff';
                var area_geojson = layer.toGeoJSON();

                var points_in_area = [];

                if (layer instanceof L.Circle) {{
                    var center = layer.getLatLng();
                    var radius = layer.getRadius();
                    points_in_area = points_data.filter(p => {{
                        var dist = mapObject.distance([p.latitude, p.longitude], center);
                        return dist <= radius;
                    }});
                }} else {{
                    points_in_area = points_data.filter(p => {{
                        var point = turf.point([p.longitude, p.latitude]);
                        return turf.booleanPointInPolygon(point, area_geojson);
                    }});
                }}

                points_in_area.forEach(p => {{
                    var row = Object.assign({{}}, p);
                    row['COR_POLIGONO'] = color;
                    all_exported_data.push(row);
                }});
            }});

            if (all_exported_data.length === 0) {{
                alert('Nenhum ponto encontrado dentro das áreas desenhadas.');
                return;
            }}

            exportToCsv('dados_com_cores.csv', all_exported_data);
        }};

        function exportToCsv(filename, rows) {{
            var processRow = function (row) {{
                return Object.values(row).map(v => {{
                    var s = v === null ? '' : v.toString();
                    return '"' + s.replace(/"/g, '""') + '"';
                }}).join(';');
            }};

            var csvFile = '\\uFEFF';
            var headers = Object.keys(rows[0]).filter(h => !['estado_filtro', 'cidade_filtro'].includes(h));
            csvFile += headers.join(';') + '\\n';

            rows.forEach(r => {{
                var cleanRow = {{}};
                headers.forEach(h => cleanRow[h] = r[h]);
                csvFile += processRow(cleanRow) + '\\n';
            }});

            var blob = new Blob([csvFile], {{ type: 'text/csv;charset=utf-8;' }});
            var link = document.createElement("a");
            link.setAttribute("href", URL.createObjectURL(blob));
            link.setAttribute("download", filename);
            link.click();
        }}
    }});
    </script>
    """
    mapa.get_root().html.add_child(folium.Element(script_js))

    caminho_completo = os.path.abspath(ARQUIVO_MAPA)
    mapa.save(caminho_completo)
    print(f"✅ Mapa Premium Gerado: {caminho_completo}")

    try:
        chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe %s'
        webbrowser.get(chrome_path).open(caminho_completo)
    except:
        webbrowser.open(caminho_completo)


if __name__ == "__main__":
    gerar_mapa()
