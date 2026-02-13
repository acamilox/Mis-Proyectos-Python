# -*- coding: utf-8 -*-
"""
Factory de gráficos para el Dashboard UCI
Crea visualizaciones Plotly de forma modular
"""

import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from typing import List, Optional, Dict, Any
from dash import html
import dash_bootstrap_components as dbc

from .config import COLORS, MESES, GRAPH_CONFIG, CHART_COLORS


class ChartFactory:
    """Factory para crear gráficos del Dashboard"""
    
    def __init__(self):
        self.colors = COLORS
        self.chart_colors = CHART_COLORS
        self.graph_config = GRAPH_CONFIG
        self.meses = MESES
    
    def create_bar_chart(
        self, 
        data: pd.Series, 
        title: str, 
        color: str = None,
        decimals: int = 0
    ) -> go.Figure:
        """
        Crea un gráfico de barras
        
        Args:
            data: Serie de datos
            title: Título del gráfico
            color: Color de las barras
            decimals: Decimales a mostrar
        """
        bar_color = color or self.colors['primary']
        
        if decimals == 0:
            hover_format = '<b>%{x}</b><br>Valor: %{y:.0f}<extra></extra>'
            text_template = '%{y:.0f}'
        else:
            hover_format = f'<b>%{{x}}</b><br>Valor: %{{y:.{decimals}f}}<extra></extra>'
            text_template = f'%{{y:.{decimals}f}}'
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=self.meses,
            y=data.values,
            marker_color=bar_color,
            marker_line_color='rgba(0,0,0,0.1)',
            marker_line_width=1,
            hovertemplate=hover_format,
            text=data.values,
            texttemplate=text_template,
            textposition='outside',
            textangle=-90,
            textfont=dict(color='white', size=11, family='Inter', weight=600)
        ))
        
        fig.update_layout(
            showlegend=False,
            plot_bgcolor=self.colors['bg_secondary'],
            paper_bgcolor=self.colors['bg_secondary'],
            title={
                'text': title,
                'font': {'size': 14, 'color': self.colors['text_primary'], 'family': 'Inter'}, 
                'x': 0.5, 
                'xanchor': 'center'
            },
            xaxis={
                'gridcolor': 'rgba(148, 163, 184, 0.1)', 
                'showline': True, 
                'linecolor': 'rgba(148, 163, 184, 0.2)', 
                'color': '#cbd5e1', 
                'tickfont': {'size': 11, 'color': '#cbd5e1'}, 
                'title': ''
            },
            yaxis={
                'gridcolor': 'rgba(148, 163, 184, 0.1)', 
                'showline': True, 
                'linecolor': 'rgba(148, 163, 184, 0.2)', 
                'color': '#94a3b8', 
                'tickformat': 'd', 
                'tickfont': {'size': 11, 'color': '#94a3b8'}, 
                'title': ''
            },
            margin=dict(l=50, r=30, t=50, b=50),
            height=350,
            hoverlabel=dict(bgcolor=self.colors['bg_primary'], font_size=13, font_color='#ffffff'),
            bargap=0.3
        )
        
        return fig
    
    def create_line_chart(
        self, 
        data: pd.Series, 
        title: str, 
        color: str = None,
        decimals: int = 1
    ) -> go.Figure:
        """
        Crea un gráfico de líneas
        
        Args:
            data: Serie de datos
            title: Título del gráfico
            color: Color de la línea
            decimals: Decimales a mostrar
        """
        line_color = color or self.colors['secondary']
        
        if decimals == 0:
            text_values = [f'{int(v)}' if pd.notna(v) else '' for v in data]
        else:
            text_values = [f'{v:.{decimals}f}' if pd.notna(v) else '' for v in data]
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=self.meses,
            y=data.values,
            mode='lines+markers+text',
            line=dict(color=line_color, width=2),
            marker=dict(size=8, color=line_color, line=dict(width=2, color=self.colors['bg_secondary'])),
            text=text_values,
            textposition='top center',
            textfont=dict(color=self.colors['text_primary'], size=10),
            hovertemplate=f'<b>%{{x}}</b><br>Valor: %{{y:.{decimals}f}}<extra></extra>'
        ))
        
        yaxis_config = {
            'gridcolor': 'rgba(148, 163, 184, 0.1)', 
            'showline': True, 
            'linecolor': 'rgba(148, 163, 184, 0.2)', 
            'color': '#94a3b8', 
            'tickfont': {'size': 11, 'color': '#94a3b8'}
        }
        
        if isinstance(data, pd.Series):
            max_val = data.max() if not data.empty else 0
            if pd.notna(max_val) and max_val <= 5:
                yaxis_config['dtick'] = 1
        
        fig.update_layout(
            title={
                'text': title, 
                'font': {'size': 14, 'color': self.colors['text_primary'], 'family': 'Inter'}, 
                'x': 0.5, 
                'xanchor': 'center'
            },
            showlegend=False,
            plot_bgcolor=self.colors['bg_secondary'],
            paper_bgcolor=self.colors['bg_secondary'],
            xaxis={
                'gridcolor': 'rgba(148, 163, 184, 0.1)', 
                'showline': True, 
                'linecolor': 'rgba(148, 163, 184, 0.2)', 
                'color': '#cbd5e1', 
                'tickfont': {'size': 11, 'color': '#cbd5e1'}, 
                'title': ''
            },
            yaxis=yaxis_config,
            margin=dict(l=50, r=30, t=50, b=50),
            height=350,
            hoverlabel=dict(bgcolor=self.colors['bg_primary'], font_size=13, font_color='#ffffff')
        )
        
        return fig
    
    def create_grouped_bar_chart(
        self, 
        row_names: List[str], 
        labels: List[str], 
        title: str,
        processor = None,
        meses: List[str] = None,
        decimals: int = 0
    ) -> go.Figure:
        """
        Crea un gráfico de barras agrupadas
        
        Args:
            row_names: Nombres de las filas a graficar
            labels: Etiquetas para cada serie
            title: Título del gráfico
            processor: DataProcessor con datos cargados
            meses: Lista de meses
            decimals: Decimales a mostrar
        """
        if meses is None:
            meses = self.meses
        
        colors = [self.colors['primary'], self.colors['secondary'], self.colors['accent'], self.colors['success']]
        
        fig = go.Figure()
        
        for i, (row_name, label) in enumerate(zip(row_names, labels)):
            if processor is not None:
                data = processor.safe_get_row(row_name, meses)
            else:
                data = pd.Series([0] * len(meses), index=meses)
            
            if decimals == 0:
                hover_template = f'<b>%{{x}}</b><br>{label}: %{{y:.0f}}<extra></extra>'
                text_template = '%{y:.0f}'
            else:
                hover_template = f'<b>%{{x}}</b><br>{label}: %{{y:.{decimals}f}}<extra></extra>'
                text_template = f'%{{y:.{decimals}f}}'
            
            fig.add_trace(go.Bar(
                name=label,
                x=meses,
                y=data.values,
                marker_color=colors[i % len(colors)],
                marker_line_color='rgba(0,0,0,0.1)',
                marker_line_width=1,
                hovertemplate=hover_template,
                text=data.values,
                texttemplate=text_template,
                textposition='outside',
                textangle=-90,
                textfont=dict(color='white', size=10, family='Inter')
            ))
        
        fig.update_layout(
            title={
                'text': title, 
                'font': {'size': 14, 'color': self.colors['text_primary'], 'family': 'Inter'}, 
                'x': 0.5, 
                'xanchor': 'center'
            },
            barmode='group',
            plot_bgcolor=self.colors['bg_secondary'],
            paper_bgcolor=self.colors['bg_secondary'],
            xaxis={
                'gridcolor': 'rgba(148, 163, 184, 0.1)', 
                'showline': True, 
                'linecolor': 'rgba(148, 163, 184, 0.2)', 
                'color': self.colors['muted'], 
                'tickfont': {'size': 11, 'color': '#94a3b8'}, 
                'title': ''
            },
            yaxis={
                'gridcolor': 'rgba(148, 163, 184, 0.1)', 
                'showline': True, 
                'linecolor': 'rgba(148, 163, 184, 0.2)', 
                'color': self.colors['muted'], 
                'tickformat': 'd', 
                'tickfont': {'size': 11, 'color': '#94a3b8'}, 
                'title': ''
            },
            margin=dict(l=50, r=30, t=50, b=70),
            height=380,
            legend=dict(
                orientation="h", 
                yanchor="top", 
                y=-0.15, 
                xanchor="center", 
                x=0.5,
                bgcolor='rgba(30, 41, 59, 0.8)', 
                font={'color': self.colors['text_primary'], 'size': 11}, 
                bordercolor='rgba(148, 163, 184, 0.2)', 
                borderwidth=1
            ),
            hoverlabel=dict(bgcolor=self.colors['bg_primary'], font_size=13, font_color='#ffffff'),
            bargap=0.25
        )
        
        return fig
    
    def create_donut_chart(
        self, 
        labels: List[str], 
        values: List[float], 
        title: str,
        colors: List[str] = None,
        auto_sort: bool = True
    ) -> go.Figure:
        """
        Crea un gráfico de dona
        
        Args:
            labels: Etiquetas
            values: Valores
            title: Título
            colors: Lista de colores
            auto_sort: Si True, ordena de mayor a menor
        """
        if colors is None:
            colors = self.chart_colors
        
        if auto_sort:
            datos = sorted(zip(values, labels, colors[:len(labels)]), reverse=True)
            values_final = [d[0] for d in datos]
            labels_final = [d[1] for d in datos]
            colors_final = [d[2] for d in datos]
        else:
            values_final = list(values)
            labels_final = list(labels)
            colors_final = list(colors[:len(labels)])
        
        total = sum(values_final)
        
        # Calcular rotación para centrar el mayor
        if total > 0:
            primer_pct = values_final[0] / total
            primer_angulo = primer_pct * 360
            rotation_ajustado = -primer_angulo / 2
        else:
            rotation_ajustado = 0
        
        fig = go.Figure(data=[go.Pie(
            labels=labels_final,
            values=values_final,
            hole=0.5,
            marker=dict(colors=colors_final, line=dict(color=self.colors['bg_secondary'], width=2)),
            textinfo='percent',
            textposition='inside',
            textfont=dict(size=11, color='white', family='Inter', weight='bold'),
            hovertemplate='<b>%{label}</b><br>Cantidad: %{value}<br>Porcentaje: %{percent}<extra></extra>',
            direction='clockwise',
            sort=False,
            rotation=rotation_ajustado
        )])
        
        # Total en el centro
        fig.add_annotation(
            text=f"<b>Total</b><br>{int(total):,}".replace(',', '.'),
            x=0.5, y=0.5,
            font=dict(size=14, color=self.colors['text_primary']),
            showarrow=False,
            xref='paper', yref='paper'
        )
        
        fig.update_layout(
            title={
                'text': title, 
                'font': {'size': 14, 'color': self.colors['text_primary'], 'family': 'Inter'}, 
                'x': 0.5, 
                'xanchor': 'center'
            },
            plot_bgcolor=self.colors['bg_secondary'],
            paper_bgcolor=self.colors['bg_secondary'],
            showlegend=True,
            legend=dict(
                orientation='v',
                yanchor='middle',
                y=0.5,
                xanchor='left',
                x=1.02,
                bgcolor='rgba(30, 41, 59, 0.8)',
                bordercolor='rgba(148, 163, 184, 0.2)',
                borderwidth=1,
                font=dict(color=self.colors['text_primary'], size=10),
                itemsizing='constant'
            ),
            margin=dict(l=20, r=150, t=50, b=30),
            height=380,
            hoverlabel=dict(bgcolor=self.colors['bg_primary'], font_size=13, font_color='#ffffff')
        )
        
        return fig
    
    def create_pyramid_chart(
        self, 
        df_fallecidos: pd.DataFrame,
        title: str = 'Pirámide de Edad y Sexo - Fallecidos'
    ) -> go.Figure:
        """
        Crea un gráfico de pirámide poblacional
        
        Args:
            df_fallecidos: DataFrame con datos de fallecidos
            title: Título del gráfico
        """
        if df_fallecidos.empty:
            fig = go.Figure()
            fig.update_layout(
                title={'text': title, 'font': {'size': 14, 'color': self.colors['text_primary']}},
                plot_bgcolor=self.colors['bg_secondary'],
                paper_bgcolor=self.colors['bg_secondary']
            )
            return fig
        
        bins = [0, 40, 50, 60, 70, 80, 90, 100]
        labels = ['<40', '40-49', '50-59', '60-69', '70-79', '80-89', '90+']
        
        df_copy = df_fallecidos.copy()
        df_copy['Rango_Edad'] = pd.cut(df_copy['Edad'], bins=bins, labels=labels, right=False)
        edad_sexo = df_copy.groupby(['Rango_Edad', 'Sexo']).size().unstack(fill_value=0)
        
        if 'Hombre' not in edad_sexo.columns:
            edad_sexo['Hombre'] = 0
        if 'Mujer' not in edad_sexo.columns:
            edad_sexo['Mujer'] = 0
        
        total = len(df_copy)
        hombres_vals = edad_sexo['Hombre'].values
        mujeres_vals = edad_sexo['Mujer'].values
        hombres_pct = [(v / total * 100) if total > 0 else 0 for v in hombres_vals]
        mujeres_pct = [(v / total * 100) if total > 0 else 0 for v in mujeres_vals]
        
        fig = go.Figure()
        
        # Hombres a la izquierda
        fig.add_trace(go.Bar(
            name='Hombres',
            y=edad_sexo.index.astype(str),
            x=[-v for v in hombres_vals],
            orientation='h',
            marker_color=self.colors['primary'],
            text=[f'{int(v)}' for v in hombres_vals],
            textposition='inside',
            textfont=dict(color='white', size=14, family='Inter', weight='bold'),
            hovertemplate='<b>%{y}</b><br>Hombres: %{customdata[0]} (%{customdata[1]:.1f}%)<extra></extra>',
            customdata=list(zip(hombres_vals, hombres_pct)),
            insidetextanchor='middle'
        ))
        
        # Mujeres a la derecha
        fig.add_trace(go.Bar(
            name='Mujeres',
            y=edad_sexo.index.astype(str),
            x=mujeres_vals,
            orientation='h',
            marker_color='#ef4444',
            text=[f'{int(v)}' for v in mujeres_vals],
            textposition='inside',
            textfont=dict(color='white', size=14, family='Inter', weight='bold'),
            hovertemplate='<b>%{y}</b><br>Mujeres: %{customdata[0]} (%{customdata[1]:.1f}%)<extra></extra>',
            customdata=list(zip(mujeres_vals, mujeres_pct)),
            insidetextanchor='middle'
        ))
        
        max_val = max(max(hombres_vals), max(mujeres_vals)) if len(hombres_vals) > 0 else 5
        
        annotations = []
        for i, (h_val, h_pct, m_val, m_pct) in enumerate(zip(hombres_vals, hombres_pct, mujeres_vals, mujeres_pct)):
            rango = str(edad_sexo.index[i])
            annotations.append(dict(
                x=-(h_val + 1.5), y=rango, text=f'{h_pct:.1f}%',
                showarrow=False, font=dict(size=11, color=self.colors['muted']), xanchor='right'
            ))
            annotations.append(dict(
                x=m_val + 1.5, y=rango, text=f'{m_pct:.1f}%',
                showarrow=False, font=dict(size=11, color=self.colors['muted']), xanchor='left'
            ))
        
        fig.update_layout(
            title={
                'text': title, 
                'font': {'size': 14, 'color': self.colors['text_primary']}, 
                'x': 0.5
            },
            barmode='relative',
            plot_bgcolor=self.colors['bg_secondary'],
            paper_bgcolor=self.colors['bg_secondary'],
            xaxis={
                'title': '', 
                'gridcolor': 'rgba(148, 163, 184, 0.1)', 
                'color': self.colors['muted'],
                'range': [-(max_val + 6), max_val + 6], 
                'showticklabels': False, 
                'zeroline': False
            },
            yaxis={
                'title': '', 
                'gridcolor': 'rgba(148, 163, 184, 0.1)', 
                'color': self.colors['text_primary'], 
                'tickfont': {'size': 12}
            },
            legend=dict(
                orientation="h", 
                y=-0.12, 
                x=0.5, 
                xanchor="center", 
                bgcolor='rgba(30, 41, 59, 0.8)',
                font={'color': self.colors['text_primary'], 'size': 12}, 
                bordercolor='rgba(148, 163, 184, 0.2)', 
                borderwidth=1
            ),
            hoverlabel=dict(bgcolor=self.colors['bg_primary'], font_color='#ffffff', font_size=14),
            margin=dict(l=70, r=70, t=50, b=60),
            height=400,
            bargap=0.15,
            annotations=annotations
        )
        
        return fig
    
    def create_kpi_card(
        self, 
        title: str, 
        value: str, 
        subtitle: str = "",
        color: str = None,
        trend: str = None,
        trend_value: str = None,
        inverse: bool = False
    ) -> dbc.Card:
        """
        Crea una tarjeta KPI
        
        Args:
            title: Título del KPI
            value: Valor principal
            subtitle: Subtítulo
            color: Color del acento
            trend: Tipo de tendencia ('up', 'down', 'stable')
            trend_value: Valor de la tendencia
            inverse: Si True, invierte el color de la tendencia
        """
        accent_color = color or self.colors['primary']
        if accent_color in ['#34495e', '#2c3e50', '#1e293b', '#2c3e50']:
            accent_color = self.colors['primary']
        
        # Elemento de tendencia
        trend_element = None
        if trend and trend_value:
            if trend == 'up':
                trend_color = '#dc2626' if inverse else '#16a34a'
                trend_icon = '▲'
            elif trend == 'down':
                trend_color = '#16a34a' if inverse else '#dc2626'
                trend_icon = '▼'
            else:
                trend_color = '#64748b'
                trend_icon = '●'
            
            trend_element = html.Div([
                html.Span(f"{trend_icon} {trend_value}", style={
                    'fontSize': '12px', 
                    'fontWeight': '600',
                    'color': trend_color
                }),
                html.Span(" vs anterior", style={
                    'fontSize': '11px',
                    'color': '#64748b',
                    'marginLeft': '4px'
                })
            ], style={
                'marginTop': '8px',
                'display': 'flex',
                'alignItems': 'center'
            })
        
        card_content = [
            html.Div([
                html.Div(style={
                    'width': '4px',
                    'height': '100%',
                    'background': accent_color,
                    'borderRadius': '2px',
                    'position': 'absolute',
                    'left': '0',
                    'top': '0'
                })
            ], style={'position': 'relative'}),
            html.H6(self._capitalizar(title), style={
                'color': self.colors['muted'], 
                'fontSize': '12px', 
                'marginBottom': '8px', 
                'fontWeight': '500', 
                'textTransform': 'uppercase', 
                'letterSpacing': '0.3px'
            }),
            html.H2(value, style={
                'color': self.colors['text_primary'], 
                'fontWeight': '700', 
                'marginBottom': '4px', 
                'fontSize': '28px', 
                'lineHeight': '1'
            }),
            html.P(self._capitalizar(subtitle), style={
                'color': '#64748b', 
                'fontSize': '12px', 
                'margin': '0', 
                'fontWeight': '400'
            })
        ]
        
        if trend_element:
            card_content.append(trend_element)
        
        return dbc.Card(
            dbc.CardBody([
                html.Div(card_content, style={'textAlign': 'left', 'paddingLeft': '12px'})
            ], style={'padding': '16px'}),
            className="kpi-card h-100",
            style={
                'borderLeft': f'4px solid {accent_color}',
                'borderRadius': '8px', 
                'boxShadow': '0 1px 3px rgba(0,0,0,0.2)',
                'background': 'rgba(30, 41, 59, 0.8)',
                'height': '100%',
                'position': 'relative'
            }
        )
    
    def _capitalizar(self, texto: str) -> str:
        """Capitaliza la primera letra"""
        if not texto or not isinstance(texto, str):
            return texto
        return texto[0].upper() + texto[1:].lower() if len(texto) > 1 else texto.upper()
    
    def get_graph_config(self) -> Dict[str, Any]:
        """Retorna la configuración de gráficos estáticos"""
        return self.graph_config


# Instancia global del factory
chart_factory = ChartFactory()
