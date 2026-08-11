import streamlit as st
import pandas as pd
import sqlite3

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")


# --- БАЗА ДАННЫХ (Версия 4 - с расширенными контрактами) ---
def get_connection():
    return sqlite3.connect('procurement_v4.db')


def init_db():
    conn = get_connection()
    c = conn.cursor()
    # Основная таблица
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
                     year
                     INTEGER,
                     plan_2026
                     REAL,
                     plan_2027
                     REAL,
                     plan_2028
                     REAL,
                     okpd2
                     TEXT,
                     kosgu
                     TEXT,
                     basis
                     TEXT,
                     ifo
                     TEXT,
                     request_num
                     TEXT,
                     plan_graph_num
                     TEXT
                 )''')
    # Таблица контрактов (расширенная)
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
                     contract_date
                     TEXT,
                     one_s_num
                     TEXT,
                     ifo
                     TEXT,
                     contract_sum
                     REAL,
                     extra_sum
                     REAL
                 )''')
    conn.commit()
    conn.close()


init_db()


# --- ФУНКЦИИ ОБНОВЛЕНИЯ ---
def update_purchase(row):
    conn = get_connection()
    c = conn.cursor()
    c.execute('''UPDATE purchases
                 SET subdivision=?,
                     name=?,
                     year=?,
                     plan_2026=?,
                     plan_2027=?,
                     plan_2028=?,
                     okpd2=?,
                     kosgu=?,
                     basis=?,
                     ifo=?,
                     request_num=?,
                     plan_graph_num=?
                 WHERE id = ?''', (*row[1:], row[0]))
    conn.commit()
    conn.close()


# --- ИНТЕРФЕЙС ---
st.title("📋 Учет закупок и контрактов")

# Поток ввода новых данных
with st.expander("➕ Добавить новую позицию закупки"):
    with st.form("new_proc", clear_on_submit=True):
        col1, col2, col3 = st.columns([2, 3, 1])
        sub = col1.selectbox("Подразделение", ["Автохозяйство", "Админ. отдел", "Лаборатория", "АХО"])
        name = col2.text_input("Наименование объекта")
        year = col3.number_input("Год разм.", value=2026, step=1)  # "Колесико" для года

        c1, c2, c3 = st.columns(3)
        p26 = c1.number_input("План 2026 (руб.)", format="%.2f", step=0.01)
        p27 = c2.number_input("План 2027 (руб.)", format="%.2f", step=0.01)
        p28 = c3.number_input("План 2028 (руб.)", format="%.2f", step=0.01)

        if st.form_submit_button("Сохранить позицию"):
            conn = get_connection()
            conn.cursor().execute('''INSERT INTO purchases
                                         (subdivision, name, year, plan_2026, plan_2027, plan_2028)
                                     VALUES (?, ?, ?, ?, ?, ?)''', (sub, name, year, p26, p27, p28))
            conn.commit()
            conn.close()
            st.rerun()

# --- ОСНОВНАЯ ЧАСТЬ (Таблица + Детали) ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

col_main, col_details = st.columns([2, 1])  # Делим экран: таблица и мини-таблица контрактов

with col_main:
    st.subheader("Главный реестр")
    if not df.empty:
        # Редактируемая таблица
        edited_df = st.data_editor(df, use_container_width=True, hide_index=True, key="main_editor")

        # Кнопка сохранения изменений в таблице
        if st.button("💾 Сохранить изменения в таблице"):
            for index, row in edited_df.iterrows():
                update_purchase(tuple(row))
            st.success("Данные обновлены!")
    else:
        st.info("Реестр пуст")

with col_details:
    st.subheader("🔗 Детали контрактов")
    if not df.empty:
        # Выбираем строку для просмотра контрактов
        selected_id = st.selectbox("Выберите закупку для контрактов", df['id'],
                                   format_func=lambda x: f"ID {x}: {df[df['id'] == x]['name'].values[0][:30]}...")

        # Показываем существующие контракты
        conn = get_connection()
        c_df = pd.read_sql_query(f"SELECT * FROM contracts WHERE purchase_id={selected_id}", conn)
        conn.close()

        if not c_df.empty:
            st.write("Список контрактов:")
            st.dataframe(c_df.drop(columns=['id', 'purchase_id']), hide_index=True)

        # Форма добавления контракта (то, что просила мама)
        with st.form("add_contract"):
            st.markdown("**Новый контракт**")
            c_num = st.text_input("№ контракта")
            c_date = st.text_input("Дата (дд.мм.гггг)")
            c_1s = st.text_input("№ 1С")
            c_ifo = st.text_input("ИФО")
            c_sum = st.number_input("Сумма", format="%.2f", step=0.01)
            c_extra = st.number_input("Доп. сумма", format="%.2f", step=0.01)

            if st.form_submit_button("Привязать контракт"):
                conn = get_connection()
                conn.cursor().execute('''INSERT INTO contracts
                                         (purchase_id, contract_num, contract_date, one_s_num, ifo, contract_sum,
                                          extra_sum)
                                         VALUES (?, ?, ?, ?, ?, ?, ?)''',
                                      (selected_id, c_num, c_date, c_1s, c_ifo, c_sum, c_extra))
                conn.commit()
                conn.close()
                st.rerun()