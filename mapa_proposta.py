import os
import re
import json
import glob
import base64
import pandas as pd
import folium

def parse_coordinate(coord_str, is_latitude=True):
    """
    Converte uma string representando uma coordenada em DMS ou decimal para float.
    Suporta:
      - Formato decimal (ex.: "-21.4755" ou "21.4755")
      - Formato DMS (ex.: "21° 28′ 32″ Sul" ou "21°28'32\" S")
      - Hemisférios informados tanto por letra quanto por extenso (NORTE, SUL, LESTE, OESTE)

    Se não encontrar indicador de hemisfério e o número for positivo, assume-se o hemisfério do Brasil:
      - Latitude: S (negativo)
      - Longitude: W (negativo)
    """
    if not isinstance(coord_str, str):
        try:
            return float(coord_str)
        except:
            return None

    c = coord_str.strip().replace(',', '.').upper()

    hemisphere = None
    if "NORTE" in c:
        hemisphere = "N"
        c = c.replace("NORTE", "")
    elif "SUL" in c:
        hemisphere = "S"
        c = c.replace("SUL", "")
    elif "LESTE" in c:
        hemisphere = "E"
        c = c.replace("LESTE", "")
    elif "OESTE" in c:
        hemisphere = "W"
        c = c.replace("OESTE", "")
    else:
        match = re.search(r'\b[NSWE]\b', c)
        if match:
            hemisphere = match.group(0)
            c = re.sub(r'\b[NSWE]\b', '', c)

    c = c.strip()

    dms_pattern = re.compile(
        r'^-?\d+(?:\.\d+)?[°:\s]+\d+(?:\.\d+)?[\'′\s]+\d+(?:\.\d+)?["″]?$'
    )

    if dms_pattern.match(c):
        pattern = r'^(-?\d+(?:\.\d+)?)[°:\s]+(\d+(?:\.\d+)?)[\'′\s]+(\d+(?:\.\d+)?)(?:["″])?$'
        m = re.match(pattern, c)
        if not m:
            return None
        deg = float(m.group(1))
        mn = float(m.group(2))
        sc = float(m.group(3))
        dec = abs(deg) + mn/60 + sc/3600
        sign_from_number = -1 if deg < 0 else 1
    else:
        try:
            dec = float(c)
        except ValueError:
            return None
        sign_from_number = -1 if dec < 0 else 1
        dec = abs(dec)

    if hemisphere == 'S' or hemisphere == 'W':
        dec = -abs(dec)
    elif hemisphere == 'N' or hemisphere == 'E':
        dec = abs(dec)
    else:
        dec = sign_from_number * dec
        if dec > 0:
            dec = -dec

    if is_latitude and not (-90 <= dec <= 90):
        return None
    if not is_latitude and not (-180 <= dec <= 180):
        return None

    return dec

def carregar_dados(file_path, sheet_name):
    """
    Lê as colunas do Excel:
      - Coluna C (índice 2): NUMERO_PI
      - Coluna U (20): CULTURA
      - Coluna AC (28): Municipio
      - Coluna AD (29): UF
      - Coluna AE (30): Latitude
      - Coluna AF (31): Longitude
    Converte as coordenadas usando parse_coordinate e filtra linhas inválidas.
    Mostra em debug as linhas removidas para facilitar a identificação de erros.
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0,
                       usecols=[2, 20, 28, 29, 30, 31])
    df.columns = ["NUMERO_PI", "CULTURA", "Municipio", "UF", "Latitude", "Longitude"]

    original_count = len(df)

    df["CULTURA"] = df["CULTURA"].astype(str).str.strip().str.capitalize()
    df["UF"] = df["UF"].astype(str).str.strip().str.upper()

    df["Latitude"] = df["Latitude"].apply(lambda x: parse_coordinate(str(x), True))
    df["Longitude"] = df["Longitude"].apply(lambda x: parse_coordinate(str(x), False))

    df_before_filter = df.copy()

    df = df[df["Latitude"].notna() & df["Longitude"].notna()]

    df = df[
        (df["Latitude"] >= -34) & (df["Latitude"] <= 5) &
        (df["Longitude"] >= -74) & (df["Longitude"] <= -34)
    ]

    final_count = len(df)

    if final_count < original_count:
        print(f"Linhas iniciais: {original_count}")
        print(f"Linhas finais:  {final_count}")
        print("As linhas abaixo foram removidas (coordenadas inválidas ou fora do Brasil):")
        removed = df_before_filter[~df_before_filter.index.isin(df.index)]
        print(removed[["NUMERO_PI", "Latitude", "Longitude"]].to_string(index=False))
        print("Verifique se há erro de digitação ou se está fora da faixa do Brasil.")

    correcoes = {
        2520025827: (-21.4755, -56.1495),  # 21° 28′ 32″ Sul, 56° 8′ 58″ Oeste
        2520052513: (-22.2765, -53.1680),
    }
    for pi, (lat, lon) in correcoes.items():
        df.loc[df["NUMERO_PI"] == pi, ["Latitude", "Longitude"]] = (lat, lon)

    df["Estado"] = "BR" + df["UF"]

    return df

def get_logo_base64(path_logo):
    with open(path_logo, 'rb') as image_file:
        logo_data = image_file.read()
    return base64.b64encode(logo_data).decode('utf-8')

def get_marker_color(row):
    cultura = str(row["CULTURA"]).strip().lower()
    if "soja" in cultura:
        return "DeepSkyBlue"
    elif "milho" in cultura:
        return "GoldenRod"
    elif "trigo" in cultura:
        return "Purple"
    else:
        return "black"

def adicionar_legenda(mapa, df):
    culturas = sorted(df["CULTURA"].unique())
    legend_lines = ""
    for cultura in culturas:
        cor = get_marker_color({"CULTURA": cultura})
        legend_lines += (
            f'<p><span style="background-color: {cor}; display: inline-block; '
            f'width: 15px; height: 15px; margin-right: 5px; border-radius: 50%;"></span> {cultura}</p>'
        )
    legenda_html = f'''
    <div style="
        position: fixed;
        bottom: 20px;
        left: 10px;
        z-index: 9999;
        font-size: 12px;
        background-color: white;
        padding: 10px;
        border: 2px solid black;
        border-radius: 5px;
    ">
        <h4 style="margin: 0 0 10px 0; text-align: center;">Legenda</h4>
        {legend_lines}
    </div>
    '''
    mapa.get_root().html.add_child(folium.Element(legenda_html))

def adicionar_cabecalho(mapa, total_points, logo_path="layout_set_logo (1).png", milho_count=0, trigo_count=0):
    logo_base64 = get_logo_base64(logo_path)
    logo_img = f'<img src="data:image/png;base64,{logo_base64}" alt="Sompo Logo" style="width:120px; height:auto; margin-bottom:10px;" />'
    cabecalho_html = f'''
    <div style="
        position: fixed;
        top: 20px;
        right: 10px;
        z-index: 9999;
        font-size: 14px;
        background-color: white;
        padding: 10px;
        border: 2px solid black;
        border-radius: 5px;
        font-family: 'Poppins', sans-serif;
        min-width: 220px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        text-align: center;
    ">
        {logo_img}
        <h4 style="margin: 0 0 5px 0; font-weight: 600;">Sompo Agro - Proposta</h4>
        <p style="margin: 2px 0;">Março de 2025</p>
        <p style="margin: 2px 0;"><b>Total: {total_points}</b></p>
        <p style="margin: 2px 0;">Milho: {milho_count} | Trigo: {trigo_count}</p>
    </div>
    '''
    mapa.get_root().html.add_child(folium.Element(cabecalho_html))

def criar_mapa_com_camadas(df, geojson_data, output_file, logo_path="layout_set_logo (1).png"):
    total_points = len(df)
    if total_points == 0:
        print("Nenhum dado de pontos encontrado. Gerando mapa base sem marcadores.")

    # Calcula quantas propostas possuem "milho" e "trigo"
    milho_count = df[df["CULTURA"].str.contains("milho", case=False, na=False)].shape[0]
    trigo_count = df[df["CULTURA"].str.contains("trigo", case=False, na=False)].shape[0]

    breakdown_group = df.groupby(['Estado', 'CULTURA']).size().reset_index(name='Count')
    breakdown_dict = {}
    for state in breakdown_group['Estado'].unique():
        lines = []
        state_data = breakdown_group[breakdown_group['Estado'] == state]
        for _, r in state_data.iterrows():
            lines.append(f"{r['CULTURA']}: {r['Count']}")
        breakdown_dict[state] = lines

    state_names = {
        'BRAC': 'Acre', 'BRAL': 'Alagoas', 'BRAP': 'Amapá', 'BRAM': 'Amazonas',
        'BRBA': 'Bahia', 'BRCE': 'Ceará', 'BRDF': 'Distrito Federal', 'BRES': 'Espírito Santo',
        'BRGO': 'Goiás', 'BRMA': 'Maranhão', 'BRMT': 'Mato Grosso', 'BRMS': 'Mato Grosso do Sul',
        'BRMG': 'Minas Gerais', 'BRPA': 'Pará', 'BRPB': 'Paraíba', 'BRPR': 'Paraná',
        'BRPE': 'Pernambuco', 'BRPI': 'Piauí', 'BRRJ': 'Rio de Janeiro', 'BRRN': 'Rio Grande do Norte',
        'BRRS': 'Rio Grande do Sul', 'BRRO': 'Rondônia', 'BRRR': 'Roraima', 'BRSC': 'Santa Catarina',
        'BRSP': 'São Paulo', 'BRSE': 'Sergipe', 'BRTO': 'Tocantins'
    }
    for feature in geojson_data["features"]:
        estado_id = feature["properties"]["id"]
        nome_estado = state_names.get(estado_id, estado_id)
        if estado_id in breakdown_dict:
            total = sum(int(line.split(": ")[1]) for line in breakdown_dict[estado_id])
            lines = breakdown_dict[estado_id]
            table_html = (
                f"<table border='1' style='border-collapse: collapse; font-size:12px;'>"
                f"<tr><th>{nome_estado} - TOTAL: {total}</th></tr>"
            )
            for line in lines:
                table_html += f"<tr><td style='padding:2px 5px;'>{line}</td></tr>"
            table_html += "</table>"
            feature["properties"]["observacoes"] = table_html
        else:
            feature["properties"]["observacoes"] = (
                f"<table border='1' style='border-collapse: collapse; font-size:12px;'>"
                f"<tr><td>{nome_estado} - Sem dados</td></tr></table>"
            )

    mapa = folium.Map(location=[-15.8267, -47.9218], zoom_start=4.0, tiles=None)
    folium.TileLayer(
        tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        attr='<a href="https://sompoagricola.com.br">Sompo Agro</a> | Criado por Jean Lima',
        name='Sompo Agro - Relatório'
    ).add_to(mapa)

    css = """
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600&display=swap" rel="stylesheet">
    <style>
    .leaflet-tooltip {
        font-family: 'Poppins', sans-serif;
        font-size: 14px;
    }
    </style>
    """
    mapa.get_root().html.add_child(folium.Element(css))

    folium.GeoJson(
        geojson_data,
        name="Estados",
        style_function=lambda feature: {
            "fillColor": "lightblue",
            "color": "green",
            "weight": 2,
            "dashArray": "5, 5",
        },
        highlight_function=lambda feature: {
            "fillColor": "yellow",
            "color": "white",
            "weight": 3,
            "dashArray": "",
        },
        tooltip=folium.GeoJsonTooltip(
            fields=["observacoes"],
            aliases=["Culturas"],
            labels=True,
            localize=True,
            sticky=False,
            parse_html=True
        )
    ).add_to(mapa)

    # Criação dos marcadores com CircleMarker conforme solicitado:
    for _, row in df.iterrows():
        cor = get_marker_color(row)
        popup_content = (
            f"<b>NUMERO_PI:</b> {row['NUMERO_PI']}<br>"
            f"<b>Cultura:</b> {row['CULTURA']}<br>"
            f"<b>Município:</b> {row['Municipio']}<br>"
            f"<b>UF:</b> {row['UF']}"
        )
        folium.CircleMarker(
            location=[row['Latitude'], row['Longitude']],
            radius=2.5,           # Círculo um pouco maior
            color="black",        # Borda escura
            weight=2,             # Espessura da borda
            fill=True,
            fill_color=cor,
            fill_opacity=0.8,     # Centro quase opaco
            opacity=0.5,          # Borda com opacidade ajustada
            popup=folium.Popup(popup_content, max_width=250)
        ).add_to(mapa)

    adicionar_cabecalho(mapa, total_points, logo_path, milho_count, trigo_count)
    adicionar_legenda(mapa, df)
    mapa.save(output_file)
    print(f"Mapa salvo em: {output_file}")

def encontrar_planilha(diretorio):
    padroes = [
        "Relatorio Gerencial*.xls",
        "Relatorio Gerencial*.xlsx",
        "Relatório Gerencial*.xls",
        "Relatório Gerencial*.xlsx"
    ]
    arquivos_encontrados = []
    for padrao in padroes:
        arquivos = glob.glob(os.path.join(diretorio, padrao))
        arquivos_encontrados.extend(arquivos)
    if arquivos_encontrados:
        return max(arquivos_encontrados, key=os.path.getctime)
    else:
        return None

if __name__ == "__main__":
    diretorio_planilhas = r"C:\Users\artsj\OneDrive\Área de Trabalho\Python" #Mudar o caminho
    file_path = encontrar_planilha(diretorio_planilhas)
    if not file_path:
        print("Nenhuma planilha encontrada no diretório especificado.")
    else:
        sheet_name = 'Sheet1'
        output_file = os.path.join(diretorio_planilhas, "mapa_proposta.html")
        geojson_path = os.path.join(diretorio_planilhas, "br.json")
        logo_path = os.path.join(diretorio_planilhas, "layout_set_logo (1).png")

        df = carregar_dados(file_path, sheet_name)
        with open(geojson_path, encoding="utf-8") as f:
            geojson_data = json.load(f)

        criar_mapa_com_camadas(df, geojson_data, output_file, logo_path)
