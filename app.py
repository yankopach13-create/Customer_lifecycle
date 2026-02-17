"""
Цикл жизни клиента в продукте.
Streamlit-приложение для загрузки отчётов из Qlik по шаблону.
"""

import re
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# Структура документа: столбец 0 — категория/продукт, 1—2 — период, 3 — количество, 4 — код клиента
COL_CATEGORY = "category"
COL_PERIOD_MAIN = "period_main"
COL_PERIOD_SUB = "period_sub"
COL_QUANTITY = "quantity"
COL_CLIENT = "client_id"

# 8 поведенческих кластеров: порядок по убыванию активности, название → описание
CLUSTER_8_ORDER = [
    "Активные (VIP)",
    "Регулярные с высоким объёмом",
    "Крупные нерегулярные",
    "Средняя активность",
    "Периодические (малый объём)",
    "Низкая активность",
    "Разовая крупная покупка",
    "Разовая покупка",
]
CLUSTER_8_DESCRIPTIONS = {
    "Активные (VIP)": "Высокий объём покупок и высокая регулярность: покупают часто и много.",
    "Регулярные с высоким объёмом": "Высокий объём и стабильная регулярность: постоянные крупные покупатели.",
    "Крупные нерегулярные": "Высокий объём, но покупают не в каждый период: крупные, но редкие покупки.",
    "Средняя активность": "Средний объём и средняя регулярность: умеренная вовлечённость.",
    "Периодические (малый объём)": "Низкий объём, но покупают часто: стабильные малые покупки.",
    "Низкая активность": "Низкий объём и невысокая регулярность: эпизодические малые покупки.",
    "Разовая крупная покупка": "Высокий объём за один или два периода: разовая крупная сделка.",
    "Разовая покупка": "Низкий объём и одна-две покупки: попробовали продукт.",
    "Не покупали": "В выбранном окне не покупали анализируемый продукт.",
}


def _norm_client_id(ser: pd.Series) -> pd.Series:
    """Приводит коды клиентов к одному строковому виду (348385 и 348385.0 → одинаково)."""
    s = ser.astype(str).str.strip()
    # убираем хвост .0 у целых чисел, чтобы совпадали коды из разных файлов
    return s.str.replace(r"\.0$", "", regex=True)


def load_and_normalize(uploaded_file):
    """Читает Excel и приводит к стандартным столбцам: category, period_main, period_sub, quantity, client_id."""
    if uploaded_file is None:
        return None
    raw = pd.read_excel(uploaded_file, header=0)
    # Берём первые 5 столбцов по позиции
    cols = raw.iloc[:, :5].copy()
    cols.columns = [COL_CATEGORY, COL_PERIOD_MAIN, COL_PERIOD_SUB, COL_QUANTITY, COL_CLIENT]
    # Убираем строки с пустыми ключевыми полями
    cols = cols.dropna(subset=[COL_PERIOD_MAIN, COL_CLIENT])
    cols[COL_QUANTITY] = pd.to_numeric(cols[COL_QUANTITY], errors="coerce").fillna(0).astype(int)
    cols[COL_CATEGORY] = cols[COL_CATEGORY].astype(str).str.strip()
    cols[COL_CLIENT] = _norm_client_id(cols[COL_CLIENT])
    return cols


def merge_and_prepare(df1, df2):
    """Объединяет два документа и готовит период, порядок периодов и когорту (первый период клиента)."""
    df = pd.concat([df1, df2], ignore_index=True)
    df[COL_PERIOD_MAIN] = df[COL_PERIOD_MAIN].astype(str).str.strip()
    df[COL_PERIOD_SUB] = df[COL_PERIOD_SUB].astype(str).str.strip()
    # Порядок периодов для определения когорты
    period_order = (
        df[[COL_PERIOD_MAIN, COL_PERIOD_SUB]]
        .drop_duplicates()
        .sort_values([COL_PERIOD_MAIN, COL_PERIOD_SUB])
        .reset_index(drop=True)
    )
    period_order["period_rank"] = period_order.index
    df = df.merge(
        period_order,
        on=[COL_PERIOD_MAIN, COL_PERIOD_SUB],
        how="left",
    )
    first_rank = df.groupby(COL_CLIENT)["period_rank"].min().rename("first_period_rank")
    df = df.merge(first_rank, left_on=COL_CLIENT, right_index=True, how="left")
    rank_to_period = period_order.set_index("period_rank")[[COL_PERIOD_MAIN, COL_PERIOD_SUB]]
    return df, period_order, rank_to_period, first_rank


def format_period_short(period_main, period_sub):
    """Форматирует период как 25/1, 25/2 (год/неделя)."""
    pm, ps = str(period_main).strip(), str(period_sub).strip()
    year_match = re.search(r"20\d{2}|\d{4}", pm)
    year_short = year_match.group(0)[-2:] if year_match else (pm[-2:] if len(pm) >= 2 else "")
    week_match = re.search(r"\d+", ps)
    week = week_match.group(0) if week_match else ps
    return f"{year_short}/{week}" if year_short and week else f"{pm} {ps}".strip()


# Сокращения месяцев для строки «По данным за янв-фев 2025»
_MONTH_ABBR = {
    "январь": "янв", "февраль": "фев", "март": "мар", "апрель": "апр",
    "май": "май", "июнь": "июн", "июль": "июл", "август": "авг",
    "сентябрь": "сен", "октябрь": "окт", "ноябрь": "ноя", "декабрь": "дек",
}


def format_period_range_for_caption(cohorts_to_use, cohort_ranks, rank_to_period, k_periods, is_months):
    """
    Формирует подпись диапазона периодов: «По данным за 1-6 недель 2025» или «По данным за янв-фев 2025».
    """
    if not cohorts_to_use or not rank_to_period.index.size:
        return ""
    r_min = min(cohort_ranks[lb] for lb in cohorts_to_use)
    r_max = max(cohort_ranks[lb] for lb in cohorts_to_use) + int(k_periods) - 1
    r_max = min(r_max, rank_to_period.index.max())
    r_min = max(r_min, rank_to_period.index.min())
    first = rank_to_period.loc[r_min]
    last = rank_to_period.loc[r_max]
    pm_f, ps_f = str(first[COL_PERIOD_MAIN]).strip(), str(first[COL_PERIOD_SUB]).strip()
    pm_l, ps_l = str(last[COL_PERIOD_MAIN]).strip(), str(last[COL_PERIOD_SUB]).strip()
    year_match = re.search(r"20\d{2}|\d{4}", pm_f)
    year = year_match.group(0) if year_match else (pm_l if re.search(r"\d{4}", pm_l) else "")
    if is_months:
        abbr = lambda s: _MONTH_ABBR.get(s.lower(), s[:3].lower() if len(s) >= 3 else s)
        part = f"{abbr(ps_f)}-{ps_l}" if ps_f != ps_l else abbr(ps_f)
        return f"По данным за {part} {year}"
    w_f = re.search(r"\d+", ps_f)
    w_l = re.search(r"\d+", ps_l)
    week_f = w_f.group(0) if w_f else ps_f
    week_l = w_l.group(0) if w_l else ps_l
    return f"По данным за {week_f}-{week_l} недель {year}"


def build_stacked_area(
    df_plot, x_col, value_col, stack_col, title, value_label,
    x_order=None, show_title=True, xaxis_title=None, xaxis_side="bottom",
    margin_override=None,
):
    """Строит стековую диаграмму с областями (stacked area)."""
    if df_plot.empty:
        fig = go.Figure()
        fig.add_annotation(text="Нет данных", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        fig.update_layout(title=dict(text=title or "", x=0.5, xanchor="center") if show_title and title else {})
        return fig
    x_vals = x_order if x_order is not None else df_plot[x_col].unique().tolist()
    stacks = df_plot[stack_col].unique().tolist()
    fig = go.Figure()
    for s in stacks:
        sub = df_plot[df_plot[stack_col] == s]
        sub = sub.set_index(x_col)[value_col].reindex(x_vals).fillna(0)
        fig.add_trace(
            go.Scatter(
                x=x_vals,
                y=sub.tolist(),
                name=str(s),
                mode="lines",
                fill="tonexty",
                stackgroup="one",
                line=dict(width=0.5),
            )
        )
    margin = margin_override if margin_override is not None else dict(t=60, b=50)
    layout_kw = dict(
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=margin,
        template="plotly_white",
        yaxis_title=value_label,
    )
    if show_title and title:
        layout_kw["title"] = dict(text=title, x=0.5, xanchor="center")
    if xaxis_title is not None:
        layout_kw["xaxis_title"] = xaxis_title
        layout_kw["xaxis"] = dict(side=xaxis_side)
    else:
        layout_kw["xaxis_title"] = x_col
    fig.update_layout(**layout_kw)
    return fig


# Высота каждого подграфика в объединённой фигуре (пиксели)
COMBINED_CHART_ROW_HEIGHT = 260


def build_combined_two_charts(
    clients_by_period,
    qty_by_period,
    x_col,
    period_labels_short,
    stack_col,
    add_total=False,
    clients_total_values=None,
    qty_total_values=None,
):
    """
    Строит одну фигуру с двумя подграфиками (общая ось X).
    Одинаковые категории — одинаковые цвета в обоих графиках.
    При add_total и 2+ категориях добавляется серия «Итого» (один цвет в легенде).
    """
    x_vals = period_labels_short
    if clients_by_period.empty and qty_by_period.empty:
        fig = go.Figure()
        fig.add_annotation(text="Нет данных", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        return fig

    stacks_cl = clients_by_period[stack_col].unique().tolist() if not clients_by_period.empty else []
    stacks_q = qty_by_period[stack_col].unique().tolist() if not qty_by_period.empty else []
    all_stacks = list(dict.fromkeys(stacks_cl + stacks_q))
    if add_total:
        all_stacks = ["Итого"] + all_stacks
    palette = px.colors.qualitative.Plotly
    color_map = {s: palette[i % len(palette)] for i, s in enumerate(all_stacks)}

    fig = make_subplots(
        rows=2,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.04,
        row_heights=[1, 1],
        subplot_titles=("", ""),
    )

    # Верхний график: клиенты (сначала стек по категориям, затем линия Итого поверх)
    if not clients_by_period.empty:
        for s in stacks_cl:
            sub = clients_by_period[clients_by_period[stack_col] == s]
            sub = sub.set_index(x_col)["clients_count"].reindex(x_vals).fillna(0)
            fig.add_trace(
                go.Scatter(
                    x=x_vals,
                    y=sub.tolist(),
                    name=str(s),
                    mode="lines",
                    fill="tonexty",
                    stackgroup="one",
                    line=dict(width=0.5, color=color_map.get(s, None)),
                ),
                row=1,
                col=1,
            )
        if add_total and clients_total_values is not None:
            fig.add_trace(
                go.Scatter(
                    x=x_vals,
                    y=list(clients_total_values),
                    name="Итого",
                    mode="lines",
                    line=dict(width=1.5, color=color_map.get("Итого", "#636EFA"), dash="dash"),
                ),
                row=1,
                col=1,
            )

    # Нижний график: товар (те же цвета по категориям), легенду не дублируем
    if not qty_by_period.empty:
        for s in stacks_q:
            sub = qty_by_period[qty_by_period[stack_col] == s]
            sub = sub.set_index(x_col)[COL_QUANTITY].reindex(x_vals).fillna(0)
            fig.add_trace(
                go.Scatter(
                    x=x_vals,
                    y=sub.tolist(),
                    name=str(s),
                    mode="lines",
                    fill="tonexty",
                    stackgroup="two",
                    line=dict(width=0.5, color=color_map.get(s, None)),
                    showlegend=False,
                ),
                row=2,
                col=1,
            )
        if add_total and qty_total_values is not None:
            fig.add_trace(
                go.Scatter(
                    x=x_vals,
                    y=list(qty_total_values),
                    name="Итого",
                    mode="lines",
                    line=dict(width=1.5, color=color_map.get("Итого", "#636EFA"), dash="dash"),
                    showlegend=False,
                ),
                row=2,
                col=1,
            )

    total_height = COMBINED_CHART_ROW_HEIGHT * 2
    fig.update_layout(
        height=total_height,
        hovermode="x unified",
        template="plotly_white",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=40, b=40, l=80, r=40),
        hoverlabel=dict(
            namelength=-1,
            font=dict(size=12, color="black"),
            bgcolor="white",
            bordercolor="gray",
        ),
    )
    fig.update_xaxes(title_text="", side="top", row=1, col=1)
    fig.update_xaxes(title_text="", row=2, col=1)
    # Подписи осей Y слева от графика
    fig.update_yaxes(title_text="Количество клиентов", row=1, col=1, side="left")
    fig.update_yaxes(title_text="Количество товара", row=2, col=1, side="left")
    fig.update_xaxes(showspikes=True, spikemode="across+marker", spikecolor="gray", spikethickness=1)
    return fig


# --- Конфигурация страницы (для Streamlit Cloud) ---
st.set_page_config(
    page_title="Цикл жизни клиента в продукте",
    page_icon="🔄",
    layout="wide",
)

# --- Заголовок (отдельно от контента) ---
st.title("🔄 Цикл жизни клиента в продукте")
st.divider()

# --- Две колонки: инструкция слева, шаблон и загрузчик справа ---
col_instruction, col_template = st.columns([1, 1])

with col_instruction:
    st.subheader("Инструкция к загрузке 1 документа")
    st.markdown("""
    1. Зайдите в Qlik, раздел «Анализ чеков», лист «Конструктор».
    2. Отберите анализируемые продукты/категорию в одном из разрезов Группа1 / Группа2 / Группа3 / Группа4.
    3. Отберите анализируемый период и разрез (год–месяц или год–неделя).
    4. Выведите отчёт по шаблону справа.
    5. Скачайте документ в Qlik и загрузите в ячейку справа.
    """)
    st.markdown("---")
    st.subheader("Инструкция к загрузке 2 документа")
    st.markdown("""
    1. Перейдите на лист «Продажи» и отберите клиентов анализируемого периода и продукта/категории.
    2. Выберите альтернативные категории и выведите отчёт по шаблону справа на листе «Конструктор».
    3. Скачайте документ в Qlik и загрузите в ячейку справа.
    """)

with col_template:
    st.subheader("📋 Шаблон загрузки данных из Qlik")
    try:
        st.image("qlik_template_categories.png", use_container_width=True)
    except FileNotFoundError:
        st.warning("Изображение шаблона не найдено. Убедитесь, что файл `qlik_template_categories.png` в корне проекта.")
    st.markdown("---")
    st.caption("Документ 1 — выгрузка из листа «Конструктор»:")
    uploaded_file_1 = st.file_uploader(
        "Документ 1",
        type=["xlsx", "xls"],
        key="qlik_upload_1",
        label_visibility="collapsed",
    )
    st.caption("Документ 2 — выгрузка по альтернативным категориям:")
    uploaded_file_2 = st.file_uploader(
        "Документ 2",
        type=["xlsx", "xls"],
        key="qlik_upload_2",
        label_visibility="collapsed",
    )

# --- Блок отображения расчётов (графики) — показывается после загрузки обоих документов ---
if uploaded_file_1 and uploaded_file_2:
    st.divider()
    try:
        df1 = load_and_normalize(uploaded_file_1)
        df2 = load_and_normalize(uploaded_file_2)
    except Exception as e:
        st.error(f"Ошибка чтения файлов: {e}")
        df1 = df2 = None

    if df1 is not None and df2 is not None and not df1.empty and not df2.empty:
        categories_from_doc1 = sorted(df1[COL_CATEGORY].dropna().unique().tolist())
        category_label = ", ".join(categories_from_doc1) if categories_from_doc1 else "—"
        st.markdown(f"### Якорный продукт когорт: :violet[{category_label}]")

        df, period_order, rank_to_period, _ = merge_and_prepare(df1, df2)
        period_labels_short = [
            format_period_short(row[COL_PERIOD_MAIN], row[COL_PERIOD_SUB])
            for _, row in period_order[[COL_PERIOD_MAIN, COL_PERIOD_SUB]].iterrows()
        ]
        period_rank_to_short = dict(zip(period_order["period_rank"], period_labels_short))
        categories_from_doc2 = sorted(df2[COL_CATEGORY].dropna().unique().tolist())
        categories_from_doc1_set = set(categories_from_doc1)
        # В списке категорий: сначала из документа 1 (анализируемая), потом из документа 2 (другие)
        all_categories = categories_from_doc1 + [c for c in categories_from_doc2 if c not in categories_from_doc1_set]
        # Когорты с подписью вида "2025/01 (N клиентов)"
        cohort_options = []
        for r in sorted(rank_to_period.index):
            row = rank_to_period.loc[r]
            pm, ps = str(row[COL_PERIOD_MAIN]).strip(), str(row[COL_PERIOD_SUB]).strip()
            short = period_labels_short[r] if r < len(period_labels_short) else f"{pm} {ps}"
            n_clients = df1[(df1[COL_PERIOD_MAIN].astype(str).str.strip() == pm) & (df1[COL_PERIOD_SUB].astype(str).str.strip() == ps)][COL_CLIENT].nunique()
            label = f"{short} ({n_clients} клиентов)"
            cohort_options.append((r, label))
        cohort_labels = [lb for _, lb in cohort_options]
        cohort_ranks = {lb: r for r, lb in cohort_options}

        # Верхняя строка: слева — выбор когорты и категорий, справа — таблица данных
        col_filters, col_table = st.columns([1, 3])
        with col_filters:
            st.caption("Выберите когорту клиентов и анализируемый продукт")
            selected_cohort_label = st.selectbox(
                "Когорта",
                options=cohort_labels,
                key="cohort_select",
                label_visibility="collapsed",
            )
            # Шире только чипы с выбранными категориями, не сам выпадающий список
            st.markdown(
                """<style>
                span[data-baseweb="tag"] { min-width: 180px; max-width: 420px; }
                </style>""",
                unsafe_allow_html=True,
            )
            selected_categories = st.multiselect(
                "Категории",
                options=all_categories,
                default=categories_from_doc1,
                key="category_select",
                label_visibility="collapsed",
            )

        # Когорта = клиенты из документа 1, купившие на выбранной неделе (в документе 1)
        cohort_rank = cohort_ranks[selected_cohort_label]
        pm, ps = rank_to_period.loc[cohort_rank, COL_PERIOD_MAIN], rank_to_period.loc[cohort_rank, COL_PERIOD_SUB]
        pm, ps = str(pm).strip(), str(ps).strip()
        cohort_clients = set(
            df1[(df1[COL_PERIOD_MAIN].astype(str).str.strip() == pm) & (df1[COL_PERIOD_SUB].astype(str).str.strip() == ps)][COL_CLIENT].tolist()
        )
        # Нормализуем типы периода к строке (как в period_order), чтобы merge не падал по dtype
        df1_norm = df1.copy()
        df1_norm[COL_PERIOD_MAIN] = df1_norm[COL_PERIOD_MAIN].astype(str).str.strip()
        df1_norm[COL_PERIOD_SUB] = df1_norm[COL_PERIOD_SUB].astype(str).str.strip()
        df2_norm = df2.copy()
        df2_norm[COL_PERIOD_MAIN] = df2_norm[COL_PERIOD_MAIN].astype(str).str.strip()
        df2_norm[COL_PERIOD_SUB] = df2_norm[COL_PERIOD_SUB].astype(str).str.strip()
        # Документ 1: период для сопоставления с period_order (добавляем period_rank и period_label_short)
        df1_with_period = df1_norm.merge(
            period_order[[COL_PERIOD_MAIN, COL_PERIOD_SUB, "period_rank"]],
            on=[COL_PERIOD_MAIN, COL_PERIOD_SUB],
            how="left",
        )
        df1_with_period["period_label_short"] = df1_with_period["period_rank"].map(period_rank_to_short)
        df2_with_period = df2_norm.merge(
            period_order[[COL_PERIOD_MAIN, COL_PERIOD_SUB, "period_rank"]],
            on=[COL_PERIOD_MAIN, COL_PERIOD_SUB],
            how="left",
        )
        df2_with_period["period_label_short"] = df2_with_period["period_rank"].map(period_rank_to_short)
        # Данные по анализируемой категории — из документа 1; по другим категориям — из документа 2
        selected_in_doc1 = [c for c in selected_categories if c in categories_from_doc1_set]
        selected_in_doc2 = [c for c in selected_categories if c in set(categories_from_doc2)]
        # Коды клиентов в документе 2 приводим к тому же виду, что в когорте (для надёжного сопоставления)
        df2_with_period["_client_norm"] = _norm_client_id(df2_with_period[COL_CLIENT])
        df1_with_period["_client_norm"] = _norm_client_id(df1_with_period[COL_CLIENT])
        parts = []
        if selected_in_doc1:
            parts.append(
                df1_with_period[
                    df1_with_period["_client_norm"].isin(cohort_clients)
                    & df1_with_period[COL_CATEGORY].isin(selected_in_doc1)
                ].copy()
            )
        if selected_in_doc2:
            parts.append(
                df2_with_period[
                    df2_with_period["_client_norm"].isin(cohort_clients)
                    & df2_with_period[COL_CATEGORY].isin(selected_in_doc2)
                ].copy()
            )
        if parts:
            df_plot = pd.concat(parts, ignore_index=True)
            stack_col = COL_CATEGORY
        else:
            # Ничего не выбрано — показываем активность когорты только по документу 1
            df_plot = df1_with_period[df1_with_period["_client_norm"].isin(cohort_clients)].copy()
            df_plot["_total"] = "Активные клиенты"
            stack_col = "_total"
        df_plot = df_plot.drop(columns=["_client_norm"], errors="ignore")

        x_col_short = "period_label_short"

        # Верхний график: количество клиентов по периодам (стек по категориям или всего)
        if stack_col == COL_CATEGORY:
            clients_by_period = (
                df_plot.groupby([x_col_short, stack_col])[COL_CLIENT]
                .nunique()
                .reset_index()
                .rename(columns={COL_CLIENT: "clients_count"})
            )
        else:
            clients_by_period = (
                df_plot.groupby(x_col_short)[COL_CLIENT]
                .nunique()
                .reset_index()
                .rename(columns={COL_CLIENT: "clients_count"})
            )
            clients_by_period[stack_col] = "Активные клиенты"

        # Нижний график: данные по количеству товара (те же фильтры)
        if stack_col == COL_CATEGORY:
            qty_by_period = (
                df_plot.groupby([x_col_short, stack_col])[COL_QUANTITY]
                .sum()
                .reset_index()
            )
        else:
            qty_by_period = (
                df_plot.groupby(x_col_short)[COL_QUANTITY]
                .sum()
                .reset_index()
            )
            qty_by_period[stack_col] = "Товар"

        # Две таблицы: левая — клиенты, правая — товар (1-я строка итого, далее разрез по категориям)
        clients_total = (
            df_plot.groupby(x_col_short)[COL_CLIENT]
            .nunique()
            .reindex(period_labels_short)
            .fillna(0)
            .astype(int)
        )
        qty_total = (
            df_plot.groupby(x_col_short)[COL_QUANTITY]
            .sum()
            .reindex(period_labels_short)
            .fillna(0)
            .astype(int)
        )
        # Строки = категории, столбцы = недели (периоды)
        clients_by_cat = (
            df_plot.groupby([stack_col, x_col_short])[COL_CLIENT]
            .nunique()
            .unstack(fill_value=0)
            .reindex(columns=period_labels_short)
            .fillna(0)
            .astype(int)
        )
        qty_by_cat = (
            df_plot.groupby([stack_col, x_col_short])[COL_QUANTITY]
            .sum()
            .unstack(fill_value=0)
            .reindex(columns=period_labels_short)
            .fillna(0)
            .astype(int)
        )
        rows_clients = ["Итого клиентов когорты"] + clients_by_cat.index.tolist()
        table_clients = pd.DataFrame(
            [clients_total.values] + [clients_by_cat.loc[c].values for c in clients_by_cat.index],
            index=rows_clients,
            columns=period_labels_short,
        )
        rows_qty = ["Итого товаров"] + qty_by_cat.index.tolist()
        table_qty = pd.DataFrame(
            [qty_total.values] + [qty_by_cat.loc[c].values for c in qty_by_cat.index],
            index=rows_qty,
            columns=period_labels_short,
        )
        with col_table:
            col_tbl_left, col_tbl_right = st.columns(2)
            with col_tbl_left:
                st.caption("Количество клиентов")
                st.dataframe(table_clients, use_container_width=True, height="content")
            with col_tbl_right:
                st.caption("Количество товара")
                st.dataframe(table_qty, use_container_width=True, height="content")
            # Синхронный горизонтальный скролл двух таблиц (без объединения)
            st.markdown(
                """
                <script>
                (function() {
                    function findScrollable(el) {
                        if (!el) return null;
                        var s = getComputedStyle(el);
                        if ((s.overflowX === 'auto' || s.overflowX === 'scroll' || s.overflow === 'auto') && el.scrollWidth > el.clientWidth) return el;
                        for (var c = el.firstElementChild; c; c = c.nextElementSibling) {
                            var r = findScrollable(c);
                            if (r) return r;
                        }
                        return null;
                    }
                    function run() {
                        var cols = document.querySelectorAll('[data-testid="column"]');
                        var pair = [];
                        cols.forEach(function(col) {
                            var frame = col.querySelector('[data-testid="stDataFrame"]');
                            if (frame) pair.push(col);
                        });
                        if (pair.length >= 2) {
                            var lastTwo = [pair[pair.length-2], pair[pair.length-1]];
                            var left = findScrollable(lastTwo[0]);
                            var right = findScrollable(lastTwo[1]);
                            if (left && right && !left._synced) {
                                left._synced = true;
                                left.addEventListener('scroll', function() { right.scrollLeft = left.scrollLeft; });
                            }
                        }
                    }
                    setTimeout(run, 1000);
                })();
                </script>
                """,
                unsafe_allow_html=True,
            )

        # График под блоком выбора и таблицы — на всю ширину, выше
        add_total = stack_col == COL_CATEGORY and len(clients_by_period[stack_col].unique()) >= 2
        clients_total_arr = clients_total.values if add_total else None
        qty_total_arr = qty_total.values if add_total else None
        fig_combined = build_combined_two_charts(
            clients_by_period,
            qty_by_period,
            x_col_short,
            period_labels_short,
            stack_col,
            add_total=add_total,
            clients_total_values=clients_total_arr,
            qty_total_values=qty_total_arr,
        )
        st.subheader("Стековая диаграмма с областями")
        st.plotly_chart(fig_combined, use_container_width=True)

        # --- Блок «Продажи анализируемого продукта на объём якорного» ---
        st.divider()
        st.subheader("Продажи анализируемого продукта на объём якорного")

        # Определение типа периода по данным (недели или месяцы)
        period_sub_str = period_order[COL_PERIOD_SUB].astype(str).str.lower()
        is_months = period_sub_str.str.contains(r"янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек", regex=True).any()
        period_word = "месяцев" if is_months else "недель"

        st.markdown('<div id="sales-block-wrap">', unsafe_allow_html=True)
        col_cohorts_block, col_analyzed_block, col_params = st.columns([1, 1, 1])
        with col_cohorts_block:
            cohort_start_block = st.selectbox(
                "С когорты",
                options=cohort_labels,
                index=0,
                key="block_cohort_start",
            )
            cohort_end_block = st.selectbox(
                "По когорту",
                options=cohort_labels,
                index=0,
                key="block_cohort_end",
            )
        with col_analyzed_block:
            selected_categories_block = st.multiselect(
                "Анализируемый продукт",
                options=all_categories,
                default=categories_from_doc1,
                key="block_categories",
                help="Категории для расчёта ожидаемых продаж (расчёт использует только этот выбор).",
            )
        with col_params:
            n_anchor = st.number_input("Кол-во якорного товара", min_value=1, value=10, step=1, key="block_n_anchor")
            k_periods = st.number_input(
                "Недель/месяцев с покупки якорного (включая неделю/месяц когорты)",
                min_value=1,
                value=5,
                step=1,
                key="block_k_weeks",
            )
        st.markdown('</div>', unsafe_allow_html=True)

        idx_start = cohort_labels.index(cohort_start_block)
        idx_end = cohort_labels.index(cohort_end_block)
        if idx_start <= idx_end:
            cohorts_to_use = cohort_labels[idx_start : idx_end + 1]
        else:
            cohorts_to_use = cohort_labels[idx_end : idx_start + 1]

        if not cohorts_to_use:
            st.caption("Выберите хотя бы одну когорту для расчёта.")
        else:
                # Клиенты выбранных когорт (нормализованный id)
                cohort_clients_block = set()
                for lb in cohorts_to_use:
                    r = cohort_ranks[lb]
                    pm, ps = rank_to_period.loc[r, COL_PERIOD_MAIN], rank_to_period.loc[r, COL_PERIOD_SUB]
                    pm, ps = str(pm).strip(), str(ps).strip()
                    clients_r = df1[(df1[COL_PERIOD_MAIN].astype(str).str.strip() == pm) & (df1[COL_PERIOD_SUB].astype(str).str.strip() == ps)][COL_CLIENT]
                    cohort_clients_block.update(_norm_client_id(clients_r).tolist())
                # Для каждого клиента — его неделя когорты (min period_rank по док 1)
                df1_cr = df1_with_period.copy()
                df1_cr["_client_norm"] = _norm_client_id(df1_cr[COL_CLIENT])
                df1_cr = df1_cr[df1_cr["_client_norm"].isin(cohort_clients_block)]
                client_cohort_rank = df1_cr.groupby("_client_norm")["period_rank"].min().to_dict()

                # Окно для каждого клиента: [cohort_rank, cohort_rank + k_periods - 1]
                def in_window(row):
                    c = row.get("_client_norm")
                    r0 = client_cohort_rank.get(c)
                    if r0 is None:
                        return False
                    pr = row.get("period_rank")
                    if pd.isna(pr):
                        return False
                    return r0 <= pr < r0 + k_periods

                # Якорный: док 1 (вся категория якоря), только клиенты блока и окно
                df1_block = df1_with_period.copy()
                df1_block["_client_norm"] = _norm_client_id(df1_block[COL_CLIENT])
                df1_block = df1_block[df1_block["_client_norm"].isin(cohort_clients_block)]
                df1_block["_in_window"] = df1_block.apply(in_window, axis=1)
                q_anchor = df1_block.loc[df1_block["_in_window"], COL_QUANTITY].sum()

                # Анализируемый: по категориям (док 1 и док 2) для разбивки при нескольких категориях
                selected_in_doc1_block = [c for c in selected_categories_block if c in categories_from_doc1_set]
                selected_in_doc2_block = [c for c in selected_categories_block if c in set(categories_from_doc2)]
                parts_an = []
                if selected_in_doc1_block:
                    d1 = df1_with_period[df1_with_period[COL_CATEGORY].isin(selected_in_doc1_block)].copy()
                    d1["_client_norm"] = _norm_client_id(d1[COL_CLIENT])
                    d1 = d1[d1["_client_norm"].isin(cohort_clients_block)]
                    d1["_in_window"] = d1.apply(in_window, axis=1)
                    parts_an.append(d1.loc[d1["_in_window"], [COL_CATEGORY, COL_QUANTITY]])
                if selected_in_doc2_block:
                    d2 = df2_with_period[df2_with_period[COL_CATEGORY].isin(selected_in_doc2_block)].copy()
                    d2["_client_norm"] = _norm_client_id(d2[COL_CLIENT])
                    d2 = d2[d2["_client_norm"].isin(cohort_clients_block)]
                    d2["_in_window"] = d2.apply(in_window, axis=1)
                    parts_an.append(d2.loc[d2["_in_window"], [COL_CATEGORY, COL_QUANTITY]])
                if parts_an:
                    df_an = pd.concat(parts_an, ignore_index=True)
                    q_by_cat = df_an.groupby(COL_CATEGORY)[COL_QUANTITY].sum().reindex(selected_categories_block).fillna(0).astype(int)
                else:
                    q_by_cat = pd.Series(dtype=int)
                q_analyzed = int(q_by_cat.sum()) if len(q_by_cat) else 0

                if q_anchor and q_anchor > 0:
                    r_ratio = q_analyzed / q_anchor
                    expected = n_anchor * r_ratio
                    expected_int = int(round(expected))
                    anchor_name = category_label
                    period_range_caption = format_period_range_for_caption(
                        cohorts_to_use, cohort_ranks, rank_to_period, k_periods, is_months
                    )
                    # Одна категория — как раньше; несколько — разбивка «из них X ед. категория1 и Y ед. категория2»
                    if len(selected_categories_block) > 1 and len(q_by_cat) > 0:
                        expected_by_cat = (q_by_cat / q_anchor * n_anchor).round(1)
                        _fmt = lambda x: f"{x:.1f}".replace(".", ",")
                        parts_main = [f'<span class="block-num">{_fmt(expected_by_cat[c])}</span> ед. <span class="block-product">{c}</span>' for c in selected_categories_block if c in expected_by_cat.index]
                        main_tail = " и ".join(parts_main)
                        main_html = (
                            f'При продаже <span class="block-num">{int(n_anchor)}</span> ед. <span class="block-product">{anchor_name}</span> в течении '
                            f'<span class="block-num">{int(k_periods)}</span> {period_word} будет продано '
                            f'<span class="block-num">{expected_int}</span> ед., из них {main_tail}.'
                        )
                        ratio_parts = [f'<span class="block-num">{_fmt(q_by_cat[c] / q_anchor)}</span> ед. <span class="block-product">{c}</span>' for c in selected_categories_block if c in q_by_cat.index]
                        ref_html = f'Ед. анализируемого товара на ед. якорного товара: <span class="block-num">{r_ratio:.2f}</span> ед., из них {" и ".join(ratio_parts)}.'
                    else:
                        analyzed_names = selected_categories_block[0] if selected_categories_block else "анализируемого продукта"
                        main_html = (
                            f'При продаже <span class="block-num">{int(n_anchor)}</span> ед. <span class="block-product">{anchor_name}</span> в течении '
                            f'<span class="block-num">{int(k_periods)}</span> {period_word} будет продано '
                            f'<span class="block-num">{expected_int}</span> ед. <span class="block-product">{analyzed_names}</span>.'
                        )
                        ref_html = f'Ед. анализируемого товара на ед. якорного товара: <span class="block-num">{r_ratio:.2f}</span>'
                    st.markdown(
                        f"""
                        <style>
                        .block-result-box {{ background: #343a40; border: 1px solid #dee2e6; border-radius: 8px; padding: 1rem 1.25rem; margin: 0.5rem 0; color: white; }}
                        .block-result-box .block-period-caption {{ font-weight: 600; letter-spacing: 0.02em; border-bottom: 1px solid rgba(255,255,255,0.35); padding-bottom: 0.4rem; margin-bottom: 0.5rem; display: block; }}
                        .block-result-box .block-num {{ color: #e85d04; font-size: 1.25rem; font-weight: bold; }}
                        .block-result-box .block-product {{ font-style: italic; background: rgba(255, 255, 255, 0.1); color: rgba(255, 255, 255, 0.95); padding: 0.1em 0.35em; border-radius: 4px; }}
                        </style>
                        <div class="block-result-box">
                        <span class="block-period-caption">{period_range_caption}</span>
                        <p style="margin: 0 0 0.5rem 0; font-size: 1rem;">{main_html}</p>
                        <p style="margin: 0; font-size: 0.95rem;">{ref_html}</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                else:
                    st.warning("В выбранных когортах и периоде нет покупок якорного товара — коэффициент не рассчитан.")

        # --- Блок «Кластерный анализ» ---
        st.divider()
        st.subheader("Кластерный анализ")
        st.caption("Сегментация клиентов по объёму покупок и регулярности покупок выбранного продукта в первые K периодов после когорты.")

        col_cohorts_cl, col_analyzed_cl, col_params_cl = st.columns([1, 1, 1])
        with col_cohorts_cl:
            cohort_start_cluster = st.selectbox(
                "С когорты",
                options=cohort_labels,
                index=0,
                key="cluster_cohort_start",
            )
            cohort_end_cluster = st.selectbox(
                "По когорту",
                options=cohort_labels,
                index=0,
                key="cluster_cohort_end",
            )
        with col_analyzed_cl:
            selected_categories_cluster = st.multiselect(
                "Анализируемый продукт",
                options=all_categories,
                default=categories_from_doc1,
                key="cluster_categories",
                help="Категории, по которым считаются объём и регулярность покупок для кластеризации.",
            )
        with col_params_cl:
            k_periods_cluster = st.number_input(
                "Недель/месяцев с покупки якорного (включая неделю/месяц когорты)",
                min_value=1,
                value=5,
                step=1,
                key="cluster_k_periods",
            )

        idx_start_c = cohort_labels.index(cohort_start_cluster)
        idx_end_c = cohort_labels.index(cohort_end_cluster)
        if idx_start_c <= idx_end_c:
            cohorts_to_use_c = cohort_labels[idx_start_c : idx_end_c + 1]
        else:
            cohorts_to_use_c = cohort_labels[idx_end_c : idx_start_c + 1]

        if not cohorts_to_use_c:
            st.caption("Выберите хотя бы одну когорту для кластеризации.")
        elif not selected_categories_cluster:
            st.warning("Выберите анализируемый продукт для кластеризации.")
        else:
            # Клиенты выбранных когорт (нормализованный id): определяются по документу 1 (якорный продукт когорт)
            cohort_clients_c = set()
            for lb in cohorts_to_use_c:
                r = cohort_ranks[lb]
                pm, ps = rank_to_period.loc[r, COL_PERIOD_MAIN], rank_to_period.loc[r, COL_PERIOD_SUB]
                pm, ps = str(pm).strip(), str(ps).strip()
                clients_r = df1[
                    (df1[COL_PERIOD_MAIN].astype(str).str.strip() == pm)
                    & (df1[COL_PERIOD_SUB].astype(str).str.strip() == ps)
                ][COL_CLIENT]
                cohort_clients_c.update(_norm_client_id(clients_r).tolist())

            if not cohort_clients_c:
                st.info("В выбранных когортах нет клиентов (по документу 1).")
            else:
                # Для каждого клиента — его период когорты (min period_rank по документу 1)
                df1_cr = df1_with_period.copy()
                df1_cr["_client_norm"] = _norm_client_id(df1_cr[COL_CLIENT])
                df1_cr = df1_cr[df1_cr["_client_norm"].isin(cohort_clients_c)]
                client_cohort_rank = df1_cr.groupby("_client_norm")["period_rank"].min()

                # Сколько периодов доступно для наблюдения (для поздних когорт окно может упираться в конец данных)
                max_rank = int(period_order["period_rank"].max())
                k_int = int(k_periods_cluster)
                available_periods = (max_rank - client_cohort_rank + 1).clip(lower=0, upper=k_int).astype(int)

                # Собираем покупки анализируемого продукта из документа 1 и/или документа 2
                selected_in_doc1_c = [c for c in selected_categories_cluster if c in categories_from_doc1_set]
                selected_in_doc2_c = [c for c in selected_categories_cluster if c in set(categories_from_doc2)]

                def _filter_to_dynamic_window(df_src: pd.DataFrame) -> pd.DataFrame:
                    """Фильтрует строки когорты в окне [cohort_rank, cohort_rank + K)."""
                    tmp = df_src.copy()
                    tmp["_client_norm"] = _norm_client_id(tmp[COL_CLIENT])
                    tmp = tmp[tmp["_client_norm"].isin(cohort_clients_c)]
                    r0 = tmp["_client_norm"].map(client_cohort_rank)
                    delta = tmp["period_rank"] - r0
                    mask = delta.notna() & tmp["period_rank"].notna() & (delta >= 0) & (delta < k_int)
                    return tmp.loc[mask, ["_client_norm", "period_rank", COL_QUANTITY]]

                parts_p = []
                if selected_in_doc1_c:
                    parts_p.append(
                        _filter_to_dynamic_window(
                            df1_with_period[df1_with_period[COL_CATEGORY].isin(selected_in_doc1_c)]
                        )
                    )
                if selected_in_doc2_c:
                    parts_p.append(
                        _filter_to_dynamic_window(
                            df2_with_period[df2_with_period[COL_CATEGORY].isin(selected_in_doc2_c)]
                        )
                    )

                if parts_p:
                    df_p = pd.concat(parts_p, ignore_index=True)
                else:
                    df_p = pd.DataFrame(columns=["_client_norm", "period_rank", COL_QUANTITY])

                # Метрики по клиенту: объём и регулярность (доля периодов с покупкой)
                per_client = pd.DataFrame({"client_id": sorted(cohort_clients_c)})
                per_client["cohort_rank"] = per_client["client_id"].map(client_cohort_rank).astype("float")
                per_client["available_periods"] = per_client["client_id"].map(available_periods).fillna(k_int).astype(int)

                if not df_p.empty:
                    agg = (
                        df_p.groupby("_client_norm")
                        .agg(
                            volume=(COL_QUANTITY, "sum"),
                            active_periods=("period_rank", "nunique"),
                        )
                        .reset_index()
                        .rename(columns={"_client_norm": "client_id"})
                    )
                    per_client = per_client.merge(agg, on="client_id", how="left")
                per_client["volume"] = per_client["volume"].fillna(0).astype(int)
                per_client["active_periods"] = per_client["active_periods"].fillna(0).astype(int)
                denom = per_client["available_periods"].replace(0, 1)
                per_client["regularity"] = (per_client["active_periods"] / denom).clip(0, 1).astype(float)

                # Назначение одному из 8 поведенческих кластеров по объёму и регулярности (правила по перцентилям)
                per_client["cluster"] = "Не покупали"
                df_fit = per_client[per_client["volume"] > 0].copy()
                v33_val = v67_val = 0.0
                if not df_fit.empty:
                    v33_val = float(df_fit["volume"].quantile(1 / 3))
                    v67_val = float(df_fit["volume"].quantile(2 / 3))
                    v33 = v33_val
                    v67 = v67_val
                    r33 = 1 / 3
                    r67 = 2 / 3

                    def _assign_cluster(row):
                        v, r = row["volume"], row["regularity"]
                        if v >= v67:
                            if r >= r67:
                                return "Активные (VIP)"
                            if r >= r33:
                                return "Регулярные с высоким объёмом"
                            return "Разовая крупная покупка"
                        if v >= v33:
                            if r >= r67:
                                return "Средняя активность"
                            if r >= r33:
                                return "Средняя активность"
                            return "Крупные нерегулярные"
                        if r >= r67:
                            return "Периодические (малый объём)"
                        if r >= r33:
                            return "Низкая активность"
                        return "Разовая покупка"

                    df_fit["cluster"] = df_fit.apply(_assign_cluster, axis=1)
                    per_client = per_client.merge(df_fit[["client_id", "cluster"]], on="client_id", how="left", suffixes=("", "_fit"))
                    per_client["cluster"] = per_client["cluster_fit"].fillna(per_client["cluster"])
                    per_client = per_client.drop(columns=["cluster_fit"], errors="ignore")

                total_clients = len(per_client)
                k_int_cluster = int(k_periods_cluster)
                period_unit = "месяц" if is_months else "неделю"
                summary = (
                    per_client.groupby("cluster", dropna=False)
                    .agg(
                        clients=("client_id", "count"),
                        pct=("client_id", lambda s: 100.0 * len(s) / total_clients if total_clients else 0.0),
                        total_volume=("volume", "sum"),
                        avg_regularity=("regularity", "mean"),
                    )
                    .reset_index()
                )
                summary["avg_client_per_period"] = (
                    (summary["total_volume"] / summary["clients"].replace(0, 1) / k_int_cluster)
                    .round(2)
                )
                # Всегда 8 поведенческих кластеров + Не покупали; недостающие — 0
                for c in CLUSTER_8_ORDER:
                    if c not in summary["cluster"].values:
                        summary = pd.concat(
                            [
                                summary,
                                pd.DataFrame(
                                    [{
                                        "cluster": c,
                                        "clients": 0,
                                        "pct": 0.0,
                                        "total_volume": 0,
                                        "avg_regularity": 0.0,
                                        "avg_client_per_period": 0.0,
                                    }]
                                ),
                            ],
                            ignore_index=True,
                        )
                if "Не покупали" not in summary["cluster"].values:
                    summary = pd.concat(
                        [
                            summary,
                            pd.DataFrame(
                                [{
                                    "cluster": "Не покупали",
                                    "clients": 0,
                                    "pct": 0.0,
                                    "total_volume": 0,
                                    "avg_regularity": 0.0,
                                    "avg_client_per_period": 0.0,
                                }]
                            ),
                        ],
                        ignore_index=True,
                    )
                # Порядок: по убыванию активности (8 кластеров), затем «Не покупали»
                order_map = {name: i for i, name in enumerate(CLUSTER_8_ORDER)}
                order_map["Не покупали"] = 999
                summary["__order"] = summary["cluster"].map(lambda x: order_map.get(x, 500))
                summary = summary.sort_values("__order").drop(columns=["__order"])

                total_volume_all = per_client["volume"].sum()
                avg_client_per_period_all = total_volume_all / total_clients / k_int_cluster if (total_clients and k_int_cluster) else 0
                avg_regularity_all = per_client["regularity"].mean()
                row_итого = pd.DataFrame(
                    [{
                        "cluster": "Итого",
                        "clients": total_clients,
                        "pct": 100.0,
                        "total_volume": int(total_volume_all),
                        "avg_client_per_period": round(avg_client_per_period_all, 2),
                        "avg_regularity": round(avg_regularity_all, 3),
                    }]
                )
                summary = pd.concat([row_итого, summary], ignore_index=True)
                summary["total_volume"] = summary["total_volume"].astype(int)

                col_cluster = "Кластер"
                col_volume = "Объём продукта за период"
                col_avg_client = f"Средний объём продукта на клиента в {period_unit}"
                col_regularity = "Средняя регулярность покупки"
                period_word_plural = "месяцев" if is_months else "недель"
                days_per_period = 30 if is_months else 7
                summary["pct_fmt"] = summary["pct"].round(1).astype(str) + "%"

                def _criteria_text(name: str, v33: float, v67: float, k: int, is_m: bool) -> str:
                    v33s = f"{v33:.0f}" if v33 == int(v33) else f"{v33:.1f}"
                    v67s = f"{v67:.0f}" if v67 == int(v67) else f"{v67:.1f}"
                    pw = "месяцев" if is_m else "недель"
                    dp = 30 if is_m else 7
                    # Интервал в днях из доли: каждые (dp/доля) дней
                    def _days(ratio: float) -> str:
                        if ratio <= 0:
                            return "—"
                        d = round(dp / ratio)
                        return f"{max(1, int(d))} дн."
                    if name == "Активные (VIP)":
                        return f"Объём ≥ {v67s} ед. (верхняя треть). Присутствуют в ≥{max(1, round(2/3*k))} {pw} из {k} (≥67%). Приходят не реже чем каждые {_days(2/3)}."
                    if name == "Регулярные с высоким объёмом":
                        return f"Объём ≥ {v67s} ед. Присутствуют в 33–67% {pw} из {k}. Приходят в среднем каждые {_days(0.5)}–{_days(1/3)}."
                    if name == "Разовая крупная покупка":
                        return f"Объём ≥ {v67s} ед. Присутствуют реже 33% {pw} из {k}. Приходят реже чем каждые {_days(0.33)}."
                    if name == "Средняя активность":
                        return f"Объём {v33s}–{v67s} ед. (средняя треть). Присутствуют не реже 33% {pw} из {k}. Приходят не реже чем каждые {_days(1/3)}."
                    if name == "Крупные нерегулярные":
                        return f"Объём {v33s}–{v67s} ед. Присутствуют реже 33% {pw} из {k}. Приходят реже чем каждые {_days(0.33)}."
                    if name == "Периодические (малый объём)":
                        return f"Объём < {v33s} ед. (нижняя треть). Присутствуют в ≥{max(1, round(2/3*k))} {pw} из {k} (≥67%). Приходят не реже чем каждые {_days(2/3)}."
                    if name == "Низкая активность":
                        return f"Объём < {v33s} ед. Присутствуют в 33–67% {pw} из {k}. Приходят в среднем каждые {_days(0.5)}–{_days(1/3)}."
                    if name == "Разовая покупка":
                        return f"Объём < {v33s} ед. Присутствуют реже 33% {pw} из {k}. Приходят реже чем каждые {_days(0.33)} или одна покупка."
                    if name == "Не покупали":
                        return "Нет покупок анализируемого продукта в выбранном окне."
                    return ""

                st.markdown("<div style='margin-top: 1.5rem;'></div>", unsafe_allow_html=True)
                cluster_names_for_download = summary["cluster"].tolist()
                col_left_actions, col_table = st.columns([1, 4])
                with col_left_actions:
                    st.caption("**Описание и критерии** — наведите на **?** слева от строки в таблице.")
                    st.caption("**Коды клиентов** — выберите кластер и нажмите «Скачать».")
                    selected_cluster_download = st.selectbox(
                        "Кластер для скачивания",
                        options=cluster_names_for_download,
                        key="cluster_download_select",
                        label_visibility="collapsed",
                    )
                    ids_for_download = per_client[per_client["cluster"] == selected_cluster_download]["client_id"].tolist()
                    download_data = "\n".join(str(c) for c in ids_for_download)
                    st.download_button(
                        "Скачать коды (.txt)",
                        data=download_data,
                        file_name="client_codes.txt",
                        mime="text/plain",
                        key="cluster_download_btn",
                    )

                desc = CLUSTER_8_DESCRIPTIONS
                rows_html = []
                for _, r in summary.iterrows():
                    cluster_name = r["cluster"]
                    crit = _criteria_text(cluster_name, v33_val, v67_val, k_int_cluster, is_months)
                    desc_t = (desc.get(cluster_name, "") or "").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
                    crit_esc = crit.replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
                    tip_content = f"<strong>Описание:</strong><br>{desc_t}<br><br><strong>Критерии отбора:</strong><br>{crit_esc}"
                    if cluster_name == "Итого":
                        cell_icons = ""
                        cell_cluster = "<strong>Итого</strong>"
                    else:
                        cell_icons = (
                            f'<span class="cluster-tt-wrap">'
                            f'<span class="cluster-tt-icon">?</span>'
                            f'<span class="cluster-tt-box">{tip_content}</span></span>'
                        )
                        cell_cluster = cluster_name
                    pct_val = r["pct_fmt"]
                    avg_r = r["avg_regularity"] if pd.notna(r["avg_regularity"]) else 0
                    x_per = round(avg_r * k_int_cluster, 1)
                    y_pct = round(avg_r * 100, 1)
                    line1 = f"Присутствуют в {x_per} {period_word_plural} из {k_int_cluster} ({y_pct}%)"
                    if avg_r > 0.001:
                        z_days = max(1, round(days_per_period / avg_r))
                        line2 = f"Приходят в среднем каждые {z_days} дн."
                    else:
                        line2 = "Приходят редко или одна покупка"
                    reg_val = f"{line1}<br>{line2}"
                    rows_html.append(
                        f"<tr><td class=\"col-icons\">{cell_icons}</td><td>{cell_cluster}</td>"
                        f"<td>{int(r['clients'])}</td><td>{pct_val}</td>"
                        f"<td>{int(r['total_volume'])}</td><td>{r['avg_client_per_period']:.2f}</td><td>{reg_val}</td></tr>"
                    )
                thead = (
                    f"<thead><tr>"
                    f"<th class=\"col-icons\"></th><th>{col_cluster}</th>"
                    f"<th>Клиентов</th><th>% клиентов</th><th>{col_volume}</th><th>{col_avg_client}</th><th>{col_regularity}</th>"
                    f"</tr></thead>"
                )
                tbody = "<tbody>" + "".join(rows_html) + "</tbody>"
                with col_table:
                    st.markdown(
                        f'<div class="cluster-table-wrap"><table class="cluster-table">{thead}{tbody}</table></div>'
                        '<style>'
                        '.cluster-table-wrap {{ margin: 0.5rem 0; overflow-x: auto; }} '
                        '.cluster-table {{ width: 100%; border-collapse: separate; border-spacing: 0; font-size: 0.8rem; '
                        'border: 1px solid #dee2e6; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }} '
                        '.cluster-table thead th {{ position: sticky; top: 0; z-index: 100; '
                        'background: #343a40; color: #fff; font-weight: 600; padding: 6px 8px; text-align: left; '
                        'font-size: 0.8rem; box-shadow: 0 2px 2px rgba(0,0,0,0.2); white-space: nowrap; }} '
                        '.cluster-table th.col-icons, .cluster-table td.col-icons {{ width: 28px; max-width: 28px; padding: 4px 6px; text-align: center; }} '
                        '.cluster-table td {{ padding: 5px 8px; border-bottom: 1px solid #eee; background: #fff; vertical-align: top; }} '
                        '.cluster-table td:nth-child(2) {{ font-weight: 500; }} '
                        '.cluster-tt-wrap {{ position: relative; display: inline-flex; justify-content: center; }} '
                        '.cluster-tt-icon {{ display: inline-flex; align-items: center; justify-content: center; width: 18px; height: 18px; '
                        'border-radius: 50%; background: #6c757d; color: #fff; font-size: 0.7rem; font-weight: bold; cursor: help; }} '
                        '.cluster-tt-box {{ display: none; position: absolute; left: 50%; transform: translateX(-50%); bottom: 100%; margin-bottom: 4px; '
                        'background: #2d3748; color: #e2e8f0; padding: 8px 12px; border-radius: 8px; font-size: 0.75rem; line-height: 1.3; '
                        'max-width: 320px; width: max-content; box-shadow: 0 4px 12px rgba(0,0,0,0.25); z-index: 9999; pointer-events: none; }} '
                        '.cluster-tt-wrap:hover .cluster-tt-box {{ display: block; }} '
                        '.cluster-table tbody tr:hover td {{ background-color: #f8f9fa; }} '
                        '.cluster-table tbody tr:first-child td {{ background: #e85d04 !important; color: #fff !important; font-weight: bold; }} '
                        '.cluster-table tbody tr:first-child:hover td {{ background: #e85d04 !important; }} '
                        '.cluster-table tbody tr:first-child .cluster-tt-icon {{ background: rgba(255,255,255,0.5); }} '
                        '</style>',
                        unsafe_allow_html=True,
                    )
    else:
        st.warning("Загрузите оба документа в формате по шаблону (5 столбцов: категория, период, период, количество, код клиента).")
