"""
Цикл жизни клиента в продукте.
Streamlit-приложение для загрузки отчётов из Qlik по шаблону.
"""

import re
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# Структура документа: столбец 0 — категория/продукт, 1—2 — период, 3 — количество, 4 — код клиента
COL_CATEGORY = "category"
COL_PERIOD_MAIN = "period_main"
COL_PERIOD_SUB = "period_sub"
COL_QUANTITY = "quantity"
COL_CLIENT = "client_id"


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
):
    """
    Строит одну фигуру с двумя подграфиками (общая ось X).
    Одинаковые категории — одинаковые цвета в обоих графиках.
    """
    x_vals = period_labels_short
    if clients_by_period.empty and qty_by_period.empty:
        fig = go.Figure()
        fig.add_annotation(text="Нет данных", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        return fig

    stacks_cl = clients_by_period[stack_col].unique().tolist() if not clients_by_period.empty else []
    stacks_q = qty_by_period[stack_col].unique().tolist() if not qty_by_period.empty else []
    all_stacks = list(dict.fromkeys(stacks_cl + stacks_q))  # порядок: сначала из клиентов, потом товар
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

    # Верхний график: клиенты
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

    total_height = COMBINED_CHART_ROW_HEIGHT * 2
    fig.update_layout(
        height=total_height,
        hovermode="x unified",
        template="plotly_white",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=40, b=40, l=50, r=120),
        hoverlabel=dict(
            namelength=-1,
            font=dict(size=12),
            bgcolor="white",
            bordercolor="gray",
        ),
    )
    fig.update_xaxes(title_text="", side="top", row=1, col=1)
    fig.update_xaxes(title_text="", row=2, col=1)
    # Подписи осей Y справа от графика
    fig.update_yaxes(title_text="Количество клиентов", row=1, col=1, side="right")
    fig.update_yaxes(title_text="Количество товара", row=2, col=1, side="right")
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
        st.markdown(f"### Якорный продукт: :violet[{category_label}]")

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
        cohort_options = []
        for r in sorted(rank_to_period.index):
            row = rank_to_period.loc[r]
            label = f"{row[COL_PERIOD_MAIN]} {row[COL_PERIOD_SUB]}".strip()
            cohort_options.append((r, label))
        cohort_labels = [lb for _, lb in cohort_options]
        cohort_ranks = {lb: r for r, lb in cohort_options}

        # Верхняя строка: слева — выбор когорты и категорий, справа — таблица данных
        col_filters, col_table = st.columns([1, 3])
        with col_filters:
            st.caption("Выберите когорту и категорию для анализа.")
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

        # Таблица справа: строки — клиенты / товар, столбцы — недели
        clients_per_period = (
            df_plot.groupby(x_col_short)[COL_CLIENT]
            .nunique()
            .reindex(period_labels_short)
            .fillna(0)
            .astype(int)
        )
        qty_per_period = (
            df_plot.groupby(x_col_short)[COL_QUANTITY]
            .sum()
            .reindex(period_labels_short)
            .fillna(0)
            .astype(int)
        )
        table_data = pd.DataFrame(
            [clients_per_period.values, qty_per_period.values],
            index=["Количество клиентов", "Количество товара"],
            columns=period_labels_short,
        )
        with col_table:
            st.dataframe(table_data, use_container_width=True, height=120)

        # График под блоком выбора и таблицы — на всю ширину, выше
        fig_combined = build_combined_two_charts(
            clients_by_period,
            qty_by_period,
            x_col_short,
            period_labels_short,
            stack_col,
        )
        st.plotly_chart(fig_combined, use_container_width=True)
    else:
        st.warning("Загрузите оба документа в формате по шаблону (5 столбцов: категория, период, период, количество, код клиента).")
