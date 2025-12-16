# app.py
import os
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from datetime import datetime
import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go

# ----------------------------
# 1) Chargement du dataset
# ----------------------------
xlsx_path = Path("data_dashboard_large.xlsx")
if not xlsx_path.exists():
    sys.exit("Erreur : le fichier 'data_dashboard_large.xlsx' est introuvable dans le dossier courant. "
             "Placez votre fichier à côté de ce script et relancez.")

# Lecture — laissez pandas détecter les types ; parse dates
df = pd.read_excel(str(xlsx_path), engine="openpyxl", parse_dates=["Date_Transaction"])

# Vérification colonnes strictes
expected_cols = {"ID_Client", "Date_Transaction", "Montant", "Magasin", "Categorie_Produit",
                 "Quantite", "Mode_Paiement", "Satisfaction_Client"}
missing = expected_cols - set(df.columns)
if missing:
    sys.exit(f"Erreur : colonnes manquantes dans le fichier : {missing}. "
             f"Colonnes attendues EXACTES : {sorted(expected_cols)}")

# Nettoyage & conversions sûres
df['Montant'] = pd.to_numeric(df['Montant'], errors='coerce').fillna(0.0)
df['ID_Client'] = df['ID_Client'].astype(str).fillna('')
df['Magasin'] = df['Magasin'].astype(str).fillna('')
df['Categorie_Produit'] = df['Categorie_Produit'].astype(str).fillna('')
df['Mode_Paiement'] = df['Mode_Paiement'].astype(str).fillna('')
df['Quantite'] = pd.to_numeric(df['Quantite'], errors='coerce').fillna(0).astype(int)
df['Satisfaction_Client'] = pd.to_numeric(df['Satisfaction_Client'], errors='coerce').fillna(0).astype(int)
if not pd.api.types.is_datetime64_any_dtype(df['Date_Transaction']):
    df['Date_Transaction'] = pd.to_datetime(df['Date_Transaction'], errors='coerce')

# Ajout colonne date_only pour groupby journalier
df['Date_only'] = df['Date_Transaction'].dt.date

# ----------------------------
# 2) App Dash + layout
# ----------------------------
px.defaults.template = "plotly_white"
external_stylesheets = ["https://stackpath.bootstrapcdn.com/bootswatch/4.5.2/flatly/bootstrap.min.css"]
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)
server = app.server

# Filtres options (initial)
min_date = df['Date_Transaction'].min().date() if not df['Date_Transaction'].isna().all() else datetime.today().date()
max_date = df['Date_Transaction'].max().date() if not df['Date_Transaction'].isna().all() else datetime.today().date()
store_options = [{'label': s, 'value': s} for s in sorted(df['Magasin'].unique())]
cat_options = [{'label': c, 'value': c} for c in sorted(df['Categorie_Produit'].unique())]
pay_options = [{'label': p, 'value': p} for p in sorted(df['Mode_Paiement'].unique())]

def kpi_card(title, value, subtitle="", fmt="{:,.2f}"):
    try:
        display = fmt.format(value)
    except Exception:
        display = str(value)
    return html.Div(className="card m-2 p-3 shadow-sm", style={"width":"18rem","display":"inline-block"}, children=[
        html.Div(title, style={"fontWeight":"600"}),
        html.Div(display, style={"fontSize":"1.4rem","fontWeight":"700"}),
        html.Small(subtitle, className="text-muted")
    ])

app.layout = html.Div([
    html.Div([ html.H2("TP2 — Dashboard KPI Interactif"), html.P("Données : data_dashboard_large.xlsx") ],
             className="m-3", style={"textAlign":"center"}),

    # Filters
    html.Div(className="card m-3 p-3", children=[
        html.Div(className="row", children=[
            html.Div(className="col-md-3", children=[
                html.Label("Période"),
                dcc.DatePickerRange(
                    id='date-range',
                    start_date=min_date,
                    end_date=max_date,
                    display_format='YYYY-MM-DD'
                )
            ]),
            html.Div(className="col-md-3", children=[
                html.Label("Magasin"),
                dcc.Dropdown(id='filter-store', options=store_options, multi=True, placeholder="Tous magasins")
            ]),
            html.Div(className="col-md-3", children=[
                html.Label("Catégorie"),
                dcc.Dropdown(id='filter-cat', options=cat_options, multi=True, placeholder="Toutes catégories")
            ]),
            html.Div(className="col-md-3", children=[
                html.Label("Mode de paiement"),
                dcc.Dropdown(id='filter-pay', options=pay_options, multi=True, placeholder="Tous modes")
            ]),
        ])
    ]),

    # KPI row (dynamic)
    html.Div(id='kpi-row', className="m-3"),

    # Vue d'ensemble: daily sales + category pie
    html.Div(className="row m-3", children=[
        html.Div(className="col-md-6", children=[dcc.Graph(id='daily-sales-trend')]),
        html.Div(className="col-md-6", children=[dcc.Graph(id='cat-sales-pie')])
    ]),

    html.Hr(),

    # Analyse par magasin
    html.Div(className="m-3", children=[
        html.H3("Analyse par magasin"),
        html.Div(className="row", children=[
            html.Div(className="col-md-6", children=[dcc.Graph(id='store-sales-pie')]),
            html.Div(className="col-md-6", children=[dcc.Graph(id='store-avg-bar')])
        ]),
        html.Br(),
        html.H5("Ventes totales et nombre de transactions par magasin"),
        dash_table.DataTable(id='store-table',
                             page_size=10)
    ]),

    html.Hr(),

    # Analyse catégories
    html.Div(className="m-3", children=[
        html.H3("Analyse des catégories de produits"),
        html.Div(className="row", children=[
            html.Div(className="col-md-6", children=[dcc.Graph(id='qty-by-category')]),
            html.Div(className="col-md-6", children=[dcc.Graph(id='stacked-sales-cat-store')])
        ]),
        html.Br(),
        html.H5("Top 5 produits par catégorie (si colonne 'Produit' présente)"),
        html.Div(id='top-products-area')
    ]),

    html.Hr(),

    # Modes de paiement
    html.Div(className="m-3", children=[
        html.H3("Modes de paiement"),
        html.Div(className="row", children=[
            html.Div(className="col-md-6", children=[dcc.Graph(id='payment-pie')]),
            html.Div(className="col-md-6", children=[html.Div(id='top-payment-kpi')])
        ])
    ]),

    html.Hr(),

    # Satisfaction client
    html.Div(className="m-3", children=[
        html.H3("Satisfaction client"),
        html.Div(className="row", children=[
            html.Div(className="col-md-6", children=[dcc.Graph(id='satisfaction-by-store')]),
            html.Div(className="col-md-6", children=[dcc.Graph(id='satisfaction-by-category')])
        ]),
        html.Br(),
        html.H5("Distribution des scores de satisfaction"),
        dash_table.DataTable(id='satisfaction-table', page_size=10)
    ]),

    html.Div(className="m-3 text-muted", children=[
        html.P("Remarque : pour afficher le Top 5 produits par catégorie, ajoutez une colonne 'Produit' dans votre fichier Excel.")
    ])
], style={"fontFamily":"Segoe UI, Roboto, Arial, sans-serif"})

# ----------------------------
# 3) Callback : met à jour tout en fonction des filtres
# ----------------------------
@app.callback(
    Output('kpi-row', 'children'),
    Output('daily-sales-trend', 'figure'),
    Output('cat-sales-pie', 'figure'),
    Output('store-sales-pie', 'figure'),
    Output('store-avg-bar', 'figure'),
    Output('store-table', 'data'),
    Output('store-table', 'columns'),
    Output('qty-by-category', 'figure'),
    Output('stacked-sales-cat-store', 'figure'),
    Output('top-products-area', 'children'),
    Output('payment-pie', 'figure'),
    Output('top-payment-kpi', 'children'),
    Output('satisfaction-by-store', 'figure'),
    Output('satisfaction-by-category', 'figure'),
    Output('satisfaction-table', 'data'),
    Input('date-range', 'start_date'),
    Input('date-range', 'end_date'),
    Input('filter-store', 'value'),
    Input('filter-cat', 'value'),
    Input('filter-pay', 'value')
)
def update_all(start_date, end_date, stores, cats, pays):
    # Filtrage
    dff = df.copy()
    if start_date:
        dff = dff[dff['Date_Transaction'] >= pd.to_datetime(start_date)]
    if end_date:
        dff = dff[dff['Date_Transaction'] <= pd.to_datetime(end_date) + pd.Timedelta(days=1)]
    if stores:
        dff = dff[dff['Magasin'].isin(stores)]
    if cats:
        dff = dff[dff['Categorie_Produit'].isin(cats)]
    if pays:
        dff = dff[dff['Mode_Paiement'].isin(pays)]

    # KPIs
    total_ventes = dff['Montant'].sum()
    nb_transactions = len(dff)
    vente_moyenne = dff['Montant'].mean() if nb_transactions>0 else 0.0
    satisfaction_moyenne = dff['Satisfaction_Client'].mean() if nb_transactions>0 else 0.0

    kpi_children = [
        kpi_card("Total des ventes (EUR)", total_ventes, "Somme des ventes"),
        kpi_card("Nombre total de transactions", nb_transactions, "Nombre de lignes"),
        kpi_card("Montant moyen / transaction (EUR)", vente_moyenne, "Moyenne des montants"),
        kpi_card("Satisfaction moyenne (1-5)", satisfaction_moyenne, "Moyenne des scores")
    ]

    # Daily sales trend
    daily = dff.groupby(dff['Date_Transaction'].dt.date)['Montant'].sum().reset_index(name='Revenue')
    if daily.empty:
        fig_daily = go.Figure()
        fig_daily.update_layout(title="Revenu journalier (aucune donnée)")
    else:
        fig_daily = px.line(daily, x='Date_Transaction', y='Revenue', markers=True, title='Revenu journalier')
        fig_daily.update_xaxes(title='Date')
        fig_daily.update_yaxes(title='Revenue (€)')

    # Category sales pie (CA share)
    sales_by_cat = dff.groupby('Categorie_Produit')['Montant'].sum().sort_values(ascending=False)
    if sales_by_cat.empty:
        fig_cat = go.Figure()
    else:
        fig_cat = px.pie(values=sales_by_cat.values, names=sales_by_cat.index, title='Part du CA par catégorie', hole=0.35)

    # Sales by store (pie) and avg by store (bar)
    sales_store = dff.groupby('Magasin')['Montant'].sum().sort_values(ascending=False)
    if sales_store.empty:
        fig_store_pie = go.Figure()
    else:
        fig_store_pie = px.pie(values=sales_store.values, names=sales_store.index, title='Répartition du CA par magasin', hole=0.35)

    avg_store = dff.groupby('Magasin')['Montant'].mean().sort_values(ascending=False)
    if avg_store.empty:
        fig_avg_store = go.Figure()
    else:
        fig_avg_store = px.bar(x=avg_store.index, y=avg_store.values, title='Montant moyen par transaction par magasin', labels={'x':'Magasin','y':'Montant moyen (€)'})

    # Store table: total sales and count
    store_table_df = dff.groupby('Magasin').agg(Ventes_Totales=('Montant','sum'), Nb_Transactions=('Montant','count')).reset_index()
    store_table_data = store_table_df.to_dict('records')
    store_table_columns = [{"name":c, "id":c} for c in store_table_df.columns]

    # Quantity per category
    qty_cat = dff.groupby('Categorie_Produit')['Quantite'].sum().sort_values(ascending=False)
    if qty_cat.empty:
        fig_qty_cat = go.Figure()
    else:
        fig_qty_cat = px.bar(x=qty_cat.index, y=qty_cat.values, title='Quantité totale vendue par catégorie', labels={'x':'Catégorie','y':'Quantité'})

    # Stacked sales by category and store
    stacked = dff.groupby(['Categorie_Produit','Magasin'])['Montant'].sum().reset_index()
    if stacked.empty:
        fig_stacked = go.Figure()
    else:
        fig_stacked = px.bar(stacked, x='Categorie_Produit', y='Montant', color='Magasin', title='Montants des ventes par catégorie et magasin', barmode='stack')

    # Top 5 products per category - only if 'Produit' exists
    top_products_children = []
    if 'Produit' in dff.columns:
        for cat in sorted(dff['Categorie_Produit'].unique()):
            sub = dff[dff['Categorie_Produit'] == cat]
            top_prod = sub.groupby('Produit')['Quantite'].sum().sort_values(ascending=False).head(5).reset_index()
            if top_prod.empty:
                continue
            top_prod['Rank'] = range(1, len(top_prod)+1)
            table = dash_table.DataTable(
                columns=[{"name":"Rang","id":"Rank"},{"name":"Produit","id":"Produit"},{"name":"Quantité vendue","id":"Quantite"}],
                data=top_prod.to_dict('records'),
                page_size=5,
                style_cell={'textAlign':'left','padding':'5px'},
                style_header={'fontWeight':'bold'}
            )
            top_products_children.append(html.Div(className="m-2", children=[html.H6(f"Catégorie: {cat}"), table], style={"display":"inline-block","verticalAlign":"top","width":"32%"}))
        if not top_products_children:
            top_products_children = [html.Div("Aucune donnée produit disponible pour le Top 5.")]
    else:
        top_products_children = [html.Div("Colonne 'Produit' absente : impossible de générer le Top 5 des produits par catégorie.")]

    # Payment pie and top payment mode KPI
    pay_counts = dff['Mode_Paiement'].value_counts()
    if pay_counts.empty:
        fig_pay = go.Figure()
        top_payment_html = html.Div("Aucune donnée de paiement")
    else:
        fig_pay = px.pie(values=pay_counts.values, names=pay_counts.index, title='Répartition des transactions par mode de paiement', hole=0.35)
        top_mode = pay_counts.idxmax()
        top_mode_pct = (pay_counts.max() / pay_counts.sum()) * 100
        top_payment_html = kpi_card("Mode de paiement le plus utilisé", top_mode, f"Part ~ {top_mode_pct:.1f}%")

    # Satisfaction plots
    sat_store = dff.groupby('Magasin')['Satisfaction_Client'].mean().sort_values(ascending=False)
    if sat_store.empty:
        fig_sat_store = go.Figure()
    else:
        fig_sat_store = px.bar(x=sat_store.index, y=sat_store.values, title='Satisfaction moyenne par magasin', labels={'x':'Magasin','y':'Satisfaction moyenne (1-5)'})

    sat_cat = dff.groupby('Categorie_Produit')['Satisfaction_Client'].mean().sort_values(ascending=False)
    if sat_cat.empty:
        fig_sat_cat = go.Figure()
    else:
        fig_sat_cat = px.bar(x=sat_cat.index, y=sat_cat.values, title='Satisfaction moyenne par catégorie', labels={'x':'Catégorie','y':'Satisfaction moyenne (1-5)'})

    # Satisfaction distribution table
    sat_dist = dff['Satisfaction_Client'].value_counts().sort_index().reset_index()
    sat_dist.columns = ['score','count']
    total = sat_dist['count'].sum() if not sat_dist.empty else 1
    sat_dist['pct'] = (sat_dist['count'] / total * 100).round(2)
    sat_table_data = sat_dist.to_dict('records')

    return (
        kpi_children,
        fig_daily,
        fig_cat,
        fig_store_pie,
        fig_avg_store,
        store_table_data,
        store_table_columns,
        fig_qty_cat,
        fig_stacked,
        top_products_children,
        fig_pay,
        top_payment_html,
        fig_sat_store,
        fig_sat_cat,
        sat_table_data
    )

# 4) Lancement
if __name__ == "__main__":
    app.run(debug=True, port=8051)
