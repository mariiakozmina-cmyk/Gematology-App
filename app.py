import streamlit as st
import pandas as pd
import sqlite3
import re

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")


# --- ФУНКЦИИ ФОРМАТИРОВАНИЯ ---
def format_okpd(raw_text):
    clean = re.sub(r'\D', '', raw_text)
    return ".".join(clean[i:i + 2] for i in range(0, len(clean), 2))


# --- БАЗА ДАННЫХ (Версия 11) ---
def get_connection():
    return sqlite3.connect('procurement_v11.db')


def init_db():
    conn = get_connection()
    c = conn.cursor()
    # Главная таблица (закупки)
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

    # Отдельная таблица для контрактов
    c.execute('''CREATE TABLE IF NOT EXISTS contracts
                 (
                     id
                     INTEGER
                     PRIMARY
                     KEY
                     AUTOINCREMENT,
                     purchase_id
                     INTEGER,
                     contract_num
                     TEXT,
                     one_s_num
                     TEXT,
                     contract_sum
                     REAL
                 )''')

    # Таблица для распределения бюджета по ИФО
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

# --- ФОРМА ВВОДА ---
with st.expander("➕ Добавить новую позицию", expanded=True):
    with st.form("new_entry", clear_on_submit=True):
        row1 = st.columns([2, 3, 1, 2])
        sub = row1[0].selectbox("Подразделение", ["Автохозяйство", "Админ. отдел", "Лаборатория", "АХО"])
        name = row1[1].text_input("Наименование")
        y_place = row1[2].selectbox("Год размещения", list(range(2026, 2031)))
        ifo_main = row1[3].multiselect("ИФО (источники)", IFO_SOURCES, default=["ГЗ"])

        row2 = st.columns([2, 1, 1])
        okpd_raw = row2[0].text_input("ОКПД2 (вводите цифры)", placeholder="Напр: 123456")
        kosgu = row2[1].selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = row2[2].selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])

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
            st.rerun()

# --- ОСНОВНАЯ ТАБЛИЦА ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    st.subheader("Главный реестр")
    edited_df = st.data_editor(df, use_container_width=True, hide_index=True, key="main_table")

    if st.button("💾 Сохранить правки в таблице"):
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
        st.success("Данные обновлены!")

    # --- БЛОК КОНТРАКТОВ И ИФО ---
    st.divider()
    selected_name = st.selectbox("Выберите закупку для работы", df['name'].unique())
    sel_id = int(df[df['name'] == selected_name]['id'].values[0])

    c_left, c_right = st.columns(2)

    with c_left:
        st.subheader("💰 Распределение ИФО")
        years = [2026, 2027, 2028]
        for year in years:
            with st.container(border=True):
                st.write(f"**Бюджет {year}**")
                for source in IFO_SOURCES:
                    conn = get_connection()
                    old_val = conn.execute(
                        "SELECT amount FROM budget_breakdown WHERE purchase_id=? AND year=? AND ifo_name=?",
                        (sel_id, year, source)).fetchone()
                    conn.close()
                    val = st.number_input(f"{source} ({year})", value=old_val[0] if old_val else 0.0, format="%.2f",
                                          key=f"b_{sel_id}_{year}_{source}")

                    conn = get_connection()
                    conn.execute("DELETE FROM budget_breakdown WHERE purchase_id=? AND year=? AND ifo_name=?",
                                 (sel_id, year, source))
                    conn.execute("INSERT INTO budget_breakdown VALUES (?,?,?,?)", (sel_id, year, source, val))
                    conn.commit()
                    conn.close()

    with c_right:
        st.subheader("🔗 Контракты")
        with st.form("add_contract"):
            c_num = st.text_input("№ контракта")
            c_1s = st.text_input("№ 1С")
            c_sum = st.number_input("Сумма", format="%.2f")
            if st.form_submit_button("Добавить контракт"):
                conn = get_connection()
                conn.execute(
                    "INSERT INTO contracts (purchase_id, contract_num, one_s_num, contract_sum) VALUES (?,?,?,?)",
                    (sel_id, c_num, c_1s, c_sum))
                conn.commit()
                conn.close()
                st.rerun()

        conn = get_connection()
        conts = pd.read_sql_query(
            f"SELECT contract_num, one_s_num, contract_sum FROM contracts WHERE purchase_id={sel_id}", conn)
        conn.close()
        st.dataframe(conts, use_container_width=True)
