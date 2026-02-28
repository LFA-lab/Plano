"""
Module de génération du tableau de bord HTML global à partir de fichiers Excel Réserves.
Utilise les onglets "Réserves" de plusieurs .xlsx, les fusionne et produit des graphiques Plotly.
Rendu type Power BI : cartes blanches sur fond gris clair.

Dépendances (à inclure via --hidden-import pour PyInstaller) : pandas, plotly, openpyxl.
Pandas est requis pour pd.read_excel() et les agrégations ; le format .xlsx ne peut pas être lu par le module csv.
"""
import os
import glob
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


SHEET_RESERVES = "Réserves"
HEADER_ROW = 2  # Ligne 3 dans Excel (index 2) = en-têtes
COL_STATUT = "Statut"
COL_TYPE = "Type Réserve"
COL_GRAVITE = "Gravité"
OUTPUT_HTML = "dashboard_global.html"

# Couleurs gravité (alignées avec sitemark)
GRAVITY_COLORS = {"1": "#FF5252", "2": "#FF9800", "3": "#FFC107"}


def _read_reserves_from_excel(path: str, progress_callback=None, status_callback=None) -> pd.DataFrame:
    """Lit l'onglet Réserves d'un fichier Excel. Retourne un DataFrame ou None si erreur."""
    try:
        df = pd.read_excel(path, sheet_name=SHEET_RESERVES, header=HEADER_ROW)
        if df.empty or len(df.columns) < 2:
            return None
        return df
    except Exception:
        return None


def generate_global_dashboard(folder_path: str, progress_callback=None, status_callback=None):
    """
    Génère un tableau de bord HTML global à partir de tous les .xlsx du dossier.

    - Liste les .xlsx, extrait l'onglet "Réserves" de chacun.
    - Ajoute une colonne Site dérivée du nom du fichier.
    - Produit 3 graphiques : Statuts, Top 10 Types, Gravité par Site.
    - Sauvegarde dashboard_global.html (design cartes sur fond gris) et l'ouvre.

    progress_callback(pct: int, message: str) et status_callback(message: str) sont optionnels.
    """
    if status_callback:
        status_callback("Recherche des fichiers Excel...")
    pattern = os.path.join(folder_path, "*.xlsx")
    files = sorted(glob.glob(pattern))
    if not files:
        if status_callback:
            status_callback("Aucun fichier .xlsx trouvé")
        raise FileNotFoundError(f"Aucun fichier .xlsx dans le dossier : {folder_path}")

    total = len(files)
    frames = []
    for i, path in enumerate(files):
        if progress_callback:
            pct = int(20 * (i + 1) / total)
            progress_callback(pct, f"Lecture : {os.path.basename(path)}")
        if status_callback:
            status_callback(f"Lecture {i + 1}/{total} : {os.path.basename(path)}")
        df = _read_reserves_from_excel(path, progress_callback, status_callback)
        if df is not None and not df.empty:
            site = os.path.splitext(os.path.basename(path))[0]
            df = df.copy()
            df["Site"] = site
            frames.append(df)

    if not frames:
        if status_callback:
            status_callback("Aucune donnée Réserves trouvée")
        raise ValueError("Aucun onglet 'Réserves' valide trouvé dans les fichiers Excel.")

    if progress_callback:
        progress_callback(25, "Fusion des données...")
    if status_callback:
        status_callback("Fusion des données...")
    merged = pd.concat(frames, ignore_index=True)

    # Colonnes possibles (Excel peut avoir libellé ou clé)
    statut_col = COL_STATUT if COL_STATUT in merged.columns else (
        "statut" if "statut" in merged.columns else None
    )
    type_col = COL_TYPE if COL_TYPE in merged.columns else (
        "type_reserve" if "type_reserve" in merged.columns else None
    )
    gravite_col = COL_GRAVITE if COL_GRAVITE in merged.columns else (
        "gravite" if "gravite" in merged.columns else None
    )

    # Gravité en texte pour légende et color mapping robustes
    if gravite_col and gravite_col in merged.columns:
        merged = merged.copy()
        merged["Gravité"] = merged[gravite_col].fillna("").astype(str).str.strip()
        merged.loc[merged["Gravité"] == "", "Gravité"] = "Non renseigné"
    else:
        merged = merged.copy()
        merged["Gravité"] = "Non renseigné"

    if progress_callback:
        progress_callback(40, "Génération des graphiques...")
    if status_callback:
        status_callback("Génération des graphiques...")

    # 1) Répartition globale des Statuts
    if statut_col and statut_col in merged.columns:
        statuts = merged[statut_col].fillna("Non renseigné").astype(str)
        fig_statuts = px.pie(
            values=statuts.value_counts().values,
            names=statuts.value_counts().index,
            title="Répartition globale des Statuts",
        )
    else:
        fig_statuts = go.Figure().add_annotation(
            text="Colonne Statut manquante",
            xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False
        )
        fig_statuts.update_layout(title="Répartition globale des Statuts")

    # 2) Top 10 des Types de réserves
    if type_col and type_col in merged.columns:
        top_types = merged[type_col].fillna("Non renseigné").astype(str).value_counts().head(10)
        fig_types = px.bar(
            x=top_types.values,
            y=top_types.index,
            orientation="h",
            title="Top 10 des Types de réserves",
            labels={"x": "Nombre", "y": "Type"},
        )
        fig_types.update_layout(yaxis={"categoryorder": "total ascending"})
    else:
        fig_types = go.Figure().add_annotation(
            text="Colonne Type Réserve manquante",
            xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False
        )
        fig_types.update_layout(title="Top 10 des Types de réserves")

    # 3) Gravité par Site — barres empilées, sites triés par volume total décroissant
    site_grav = merged.groupby(["Site", "Gravité"], dropna=False).size().reset_index(name="Nombre")
    site_totals = merged.groupby("Site").size().sort_values(ascending=False)
    site_order = site_totals.index.tolist()
    site_grav["Site"] = pd.Categorical(site_grav["Site"], categories=site_order, ordered=True)
    site_grav = site_grav.sort_values("Site")

    if site_grav.empty or site_grav["Nombre"].sum() == 0:
        fig_gravite = go.Figure().add_annotation(
            text="Aucune donnée Gravité par Site",
            xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False
        )
        fig_gravite.update_layout(title="Gravité par Site")
    else:
        fig_gravite = px.bar(
            site_grav,
            x="Site",
            y="Nombre",
            color="Gravité",
            title="Gravité par Site",
            barmode="stack",
            color_discrete_map=GRAVITY_COLORS,
        )
        fig_gravite.update_layout(
            xaxis={"categoryorder": "array", "categoryarray": site_order},
            xaxis_tickangle=-45,
        )

    # Design HTML : cartes blanches sur fond gris (sans subplots)
    html_cards_style = """
    <style>
      body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #e8e8e8; margin: 0; padding: 24px; }
      .dashboard-title { color: #333; margin-bottom: 20px; font-size: 24px; font-weight: 600; }
      .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); padding: 24px; margin-bottom: 24px; }
    </style>
    """
    html_head = f'<!DOCTYPE html><html><head><meta charset="utf-8"><title>Tableau de bord global — Réserves</title>{html_cards_style}</head><body>'
    html_title = '<div class="dashboard-title">Tableau de bord global — Réserves</div>'
    html_foot = "</body></html>"

    part1 = fig_statuts.to_html(full_html=False, include_plotlyjs="cdn")
    part2 = fig_types.to_html(full_html=False, include_plotlyjs=False)
    part3 = fig_gravite.to_html(full_html=False, include_plotlyjs=False)

    full_html = html_head + html_title + '<div class="card">' + part1 + "</div>" + '<div class="card">' + part2 + "</div>" + '<div class="card">' + part3 + "</div>" + html_foot

    out_path = os.path.join(folder_path, OUTPUT_HTML)
    if progress_callback:
        progress_callback(85, "Sauvegarde du dashboard...")
    if status_callback:
        status_callback("Sauvegarde du dashboard...")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(full_html)

    if progress_callback:
        progress_callback(100, "Terminé")
    if status_callback:
        status_callback("Terminé")

    try:
        os.startfile(out_path)
    except AttributeError:
        import subprocess
        import platform
        if platform.system() == "Darwin":
            subprocess.run(["open", out_path], check=False)
        else:
            subprocess.run(["xdg-open", out_path], check=False)

    return out_path
