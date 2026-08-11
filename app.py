import streamlit as st
import pandas as pd
import sqlite3

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")

# --- БАЗА ДАННЫХ ---
def get_connection():
    return sqlite3.connect('procurement_v5.db')

def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS purchases
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  subdivision TEXT, name TEXT, year INTEGER,
                  plan_2026 REAL, plan_2027 REAL, plan_2028 REAL,
                  okpd2 TEXT, kosgu TEXT, basis TEXT, ifo TEXT, 
                  request_num TEXT, plan_graph_num TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS contracts
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  purchase_id INTEGER, contract_num TEXT, contract_date TEXT,
                  one_s_num TEXT, ifo TEXT, contract_sum REAL, extra_sum REAL)''')
    conn.commit()
    conn.close()

init_db()

# --- ИНТЕРФЕЙС ---
st.title("📋 Реестр закупок и контрактов")

# --- ФОРМА ВВОДА (ВСЕ ПОЛЯ ВЕРНУЛИСЬ) ---
with st.expander("➕ Добавить новую позицию закупки", expanded=True):
    with st.form("new_proc", clear_on_submit=True):
        col1, col2, col3 = st.columns([2, 3, 1])
        sub = col1.selectbox("Подразделение", ["Автохозяйство", "Админ. отдел", "Лаборатория", "АХО"])
        name = col2.text_input("Наименование объекта")
        year = col3.number_input("Год разм.", value=2026, step=1)

        c1, c2, c3 = st.columns(3)
        p26 = c1.number_input("План 2026 (руб.)", format="%.2f", step=0.01)
        p27 = c2.number_input("План 2027 (руб.)", format="%.2f", step=0.01)
        p28 = c3.number_input("План 2028 (руб.)", format="%.2f", step=0.01)

        d1, d2, d3, d4 = st.columns(4)
        okpd = d1.selectbox("ОКПД2", ["32.50.13.110", "21.20.23.110", "Другое..."])
        kosgu = d2.selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = d3.selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])

        # Логика для ИФО: список + свой вариант
        ifo_options = ["00000000000000000130", "00000000000000000180", "Ввести свой вариант..."]
        ifo_select = d4.selectbox("ИФО", ifo_options)

        # Если выбрано "Ввести свой вариант", показываем доп. поле
        final_ifo = ifo_select
        if ifo_select == "Ввести свой вариант...":
            final_ifo = st.text_input("Введите ИФО вручную")

        e1, e2 = st.columns(2)
        req_n = e1.text_input("№ предложения на закупку")
        graph_n = e2.text_input("№ план-графика")

        if st.form_submit_button("Сохранить позицию"):
            conn = get_connection()
            conn.cursor().execute('''INSERT INTO purchases 
                (subdivision, name, year, plan_2026, plan_2027, plan_2028, okpd2, kosgu, basis, ifo, request_num, plan_graph_num)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''',
                (sub, name, year, p26, p27, p28, okpd, kosgu, basis, final_ifo, req_n, graph_n))
            conn.commit()
            conn.close()
            st.rerun()

# --- ТАБЛИЦА С РУССКИМИ ЗАГОЛОВКАМИ ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    # Переименовываем колонки для отображения
    display_df = df.copy()
    display_df.columns = [
        "ID", "Подразделение", "Наименование", "Год",
        "План 2026", "План 2027", "План 2028",
        "ОКПД2", "КОСГУ", "Основание", "ИФО", "№ предл.", "№ графика"
    ]

    st.subheader("Главный реестр")
    # Редактируемая таблица с копейками
    edited_df = st.data_editor(display_df, use_container_width=True, hide_index=True)

    # Кнопка сохранения правок (если мама что-то изменила прямо в ячейках)
    if st.button("💾 Сохранить правки в таблице"):
        conn = get_connection()
        c = conn.cursor()
        for _, row in edited_df.iterrows():
            c.execute('''UPDATE purchases SET 
                         subdivision=?, name=?, year=?, plan_2026=?, plan_2027=?, plan_2028=?,
                         okpd2=?, kosgu=?, basis=?, ifo=?, request_num=?, plan_graph_num=?
                         WHERE id=?''', (row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[0]))
        conn.commit()
        conn.close()
        st.success("Изменения сохранены!")

# --- КОНТРАКТЫ (НИЖНЯЯ ПАНЕЛЬ) ---
st.divider()
col_c1, col_c2 = st.columns([1, 2])

with col_c1:
    st.subheader("📝 Добавить контракт")
    if not df.empty:
        target_id = st.selectbox("К какой закупке?", df['id'],
                                format_func=lambda x: f"ID {x}: {df[df['id']==x]['name'].values[0][:30]}")
        with st.form("add_contract"):
            c_num = st.text_input("№ контракта")
            c_1s = st.text_input("№ 1С")
            c_sum = st.number_input("Сумма", format="%.2f", step=0.01)
            if st.form_submit_button("Привязать"):
                conn = get_connection()
                conn.cursor().execute("INSERT INTO contracts (purchase_id, contract_num, one_s_num, contract_sum) VALUES (?,?,?,?)",
                                     (target_id, c_num, c_1s, c_sum))
                conn.commit()
                conn.close()
                st.rerun()

with col_c2:
    st.subheader("📄 Связанные контракты")
    if not df.empty:
        conn = get_connection()
        contracts_df = pd.read_sql_query(f"SELECT contract_num, one_s_num, contract_sum FROM contracts WHERE purchase_id={target_id}", conn)
        conn.close()
        st.dataframe(contracts_df, use_container_width=True, hide_index=True)