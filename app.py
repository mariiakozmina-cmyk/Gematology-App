import streamlit as st
import pandas as pd
import sqlite3
import re

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")


# --- ФУНКЦИИ ФОРМАТИРОВАНИЯ ---
def format_okpd(raw_text):
    """Добавляет точки через каждые 2 цифры"""
    clean = re.sub(r'\D', '', raw_text)
    return ".".join(clean[i:i + 2] for i in range(0, len(clean), 2))


# --- БАЗА ДАННЫХ (Версия 10) ---
def get_connection():
    return sqlite3.connect('procurement_v10.db')


def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS purchases
                 (
                     id
                     INTEGER
                     PRIMARY
                     KEY
                     AUTOINCREMENT,
                     subdivision
                     TEXT,
                     name
                     TEXT,
                     year_placement
                     INTEGER,
                     ifo
                     TEXT,
                     okpd2
                     TEXT,
                     kosgu
                     TEXT,
                     basis
                     TEXT,
                     request_num
                     TEXT,
                     plan_graph_num
                     TEXT,
                     plan_2026
                     REAL,
                     plan_2027
                     REAL,
                     plan_2028
                     REAL
                 )''')

    c.execute('''CREATE TABLE IF NOT EXISTS budget_breakdown
                 (
                     purchase_id
                     INTEGER,
                     year
                     INTEGER,
                     ifo_name
                     TEXT,
                     amount
                     REAL
                 )''')
    conn.commit()
    conn.close()


init_db()

IFO_SOURCES = ["ВБ", "ГЗ", "ОМС", "Прочее"]

st.title("📋 Реестр закупок")

# --- ФОРМА ВВОДА (Все поля вернулись!) ---
with st.expander("➕ Добавить новую позицию", expanded=True):
    with st.form("new_entry", clear_on_submit=True):
        # Строка 1: Основное
        row1 = st.columns([2, 3, 1, 2])
        sub = row1[0].selectbox("Подразделение", ["Автохозяйство", "Админ. отдел", "Лаборатория", "АХО"])
        name = row1[1].text_input("Наименование")
        y_place = row1[2].selectbox("Год размещения", list(range(2026, 2031)))
        ifo_main = row1[3].multiselect("ИФО (источники)", IFO_SOURCES, default=["ГЗ"])

        # Строка 2: Кодировки
        row2 = st.columns([2, 1, 1])
        # Для ОКПД2 используем подсказку, форматирование применится при сохранении
        okpd_raw = row2[0].text_input("ОКПД2 (вводите цифры)", placeholder="Напр: 123456")
        kosgu = row2[1].selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = row2[2].selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])

        # Строка 3: Номера
        row3 = st.columns(2)
        req_n = row3[0].text_input("Номер предложения на закупку")
        graph_n = row3[1].text_input("Номер план-графика")

        if st.form_submit_button("Сохранить позицию"):
            okpd_fixed = format_okpd(okpd_raw)
            conn = get_connection()
            conn.execute('''INSERT INTO purchases
                            (subdivision, name, year_placement, ifo, okpd2, kosgu, basis, request_num, plan_graph_num,
                             plan_2026, plan_2027, plan_2028)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0)''',
                         (sub, name, y_place, ", ".join(ifo_main), okpd_fixed, kosgu, basis, req_n, graph_n))
            conn.commit()
            conn.close()
            st.success(f"Добавлено! ОКПД2 отформатирован: {okpd_fixed}")
            st.rerun()

# --- ОСНОВНАЯ ТАБЛИЦА ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    st.subheader("Главный реестр")

    # Русские названия колонок
    display_df = df.copy()
    display_df.columns = [
        "ID", "Подразделение", "Наименование", "Год размещения", "ИФО",
        "ОКПД2", "КОСГУ", "Основание", "Номер предложения на закупку", "Номер план-графика",
        "Планируемая сумма, руб.; 2026 год", "Планируемая сумма, руб.; 2027 год", "Планируемая сумма, руб.; 2028 год"
    ]

    # Исправленный data_editor (теперь сохранение работает без ошибок)
    edited_df = st.data_editor(display_df, use_container_width=True, hide_index=True, key="main_table")

    if st.button("💾 Применить правки из таблицы"):
        conn = get_connection()
        for _, row in edited_df.iterrows():
            conn.execute('''UPDATE purchases
                            SET subdivision=?,
                                name=?,
                                year_placement=?,
                                ifo=?,
                                okpd2=?,
                                kosgu=?,
                                basis=?,
                                request_num=?,
                                plan_graph_num=?,
                                plan_2026=?,
                                plan_2027=?,
                                plan_2028=?
                            WHERE id = ?''',
                         (row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11],
                          row[12], row[0]))
        conn.commit()
        conn.close()
        st.success("Все изменения в базе сохранены!")

    # --- БЛОК РАСПРЕДЕЛЕНИЯ (ИФО) ---
    st.divider()
    st.subheader("💰 Распределение бюджета по источникам")

    selected_name = st.selectbox("Выберите закупку для ввода сумм", df['name'].unique())
    sel_id = int(df[df['name'] == selected_name]['id'].values[0])

    years = [2026, 2027, 2028]
    cols = st.columns(3)

    for i, year in enumerate(years):
        with cols[i]:
            st.markdown(f"**Бюджет на {year} год**")
            with st.container(border=True):
                total_for_year = 0.0
                for source in IFO_SOURCES:
                    conn = get_connection()
                    old_val = conn.execute(
                        "SELECT amount FROM budget_breakdown WHERE purchase_id=? AND year=? AND ifo_name=?",
                        (sel_id, year, source)).fetchone()
                    conn.close()

                    val = st.number_input(f"{source} ({year})", value=old_val[0] if old_val else 0.0, format="%.2f",
                                          key=f"break_{sel_id}_{year}_{source}")
                    total_for_year += val

                    conn = get_connection()
                    conn.execute("DELETE FROM budget_breakdown WHERE purchase_id=? AND year=? AND ifo_name=?",
                                 (sel_id, year, source))
                    conn.execute("INSERT INTO budget_breakdown VALUES (?,?,?,?)", (sel_id, year, source, val))
                    conn.commit()
                    conn.close()

                st.write(f"**ИТОГО {year}: {total_for_year:,.2f} руб.**")
                # Обновляем итоговую сумму в главной таблице
                conn = get_connection()
                conn.execute(f"UPDATE purchases SET plan_{year} = ? WHERE id = ?", (total_for_year, sel_id))
                conn.commit()
                conn.close()

    if st.button("🔄 Обновить итоговые суммы в таблице"):
        st.rerun()