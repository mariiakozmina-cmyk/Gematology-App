import streamlit as st
import pandas as pd
import sqlite3

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")


# --- 1. БАЗА ДАННЫХ (НОВАЯ ВЕРСИЯ v3) ---
def get_connection():
    return sqlite3.connect('procurement_v3.db')


def init_db():
    conn = get_connection()
    c = conn.cursor()
    # Создаем таблицу по твоему ТЗ из аудио
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
    # Таблица для контрактов (оставляем для расчета "Сыграно")
    c.execute('''CREATE TABLE IF NOT EXISTS contracts
                 (
                     id
                     INTEGER
                     PRIMARY
                     KEY
                     AUTOINCREMENT,
                     purchase_id
                     INTEGER,
                     contract_info
                     TEXT,
                     contract_sum
                     REAL
                 )''')
    conn.commit()
    conn.close()


init_db()

# --- 2. БОКОВАЯ ПАНЕЛЬ (ВВОД) ---
with st.sidebar:
    st.header("➕ Новая позиция")
    with st.form("input_form", clear_on_submit=True):
        sub = st.selectbox("Подразделение", ["Автохозяйство", "Административно-хозяйственный отдел"])
        name = st.text_area("Наименование")

        col_sum1, col_sum2, col_sum3 = st.columns(3)
        p26 = col_sum1.number_input("Планируемая сумма, руб., 2026 ГОД", min_value=0.0)
        p27 = col_sum2.number_input("Планируемая сумма, руб., 2027 ГОД", min_value=0.0)
        p28 = col_sum3.number_input("Планируемая сумма, руб., 2028 ГОД", min_value=0.0)

        okpd = st.text_input("ОКПД2")
        kosgu = st.text_input("КОСГУ")
        base = st.selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])
        ifo = st.text_input("ИФО")
        req_n = st.text_input("№ предложения на закупку")
        graph_n = st.text_input("№ план-графика")

        if st.form_submit_button("Добавить в реестр"):
            conn = get_connection()
            conn.cursor().execute('''INSERT INTO purchases
                                     (subdivision, name, plan_2026, plan_2027, plan_2028, okpd2, kosgu, basis, ifo,
                                      request_num, plan_graph_num)
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                  (sub, name, p26, p27, p28, okpd, kosgu, base, ifo, req_n, graph_n))
            conn.commit()
            conn.close()
            st.rerun()

# --- 3. ОСНОВНАЯ ТАБЛИЦА ---
st.title("📋 План закупок НМИЦ Гематологии")

conn = get_connection()
# Запрос, который считает "Сыграно" и "Остаток" только для 2026 года (как в твоей таблице)
query = '''
        SELECT p.subdivision, \
               p.name, \
               p.plan_2026, \
               p.plan_2027, \
               p.plan_2028,
               p.okpd2, \
               p.kosgu, \
               p.basis, \
               p.ifo, \
               p.request_num, \
               p.plan_graph_num,
               IFNULL(SUM(c.contract_sum), 0)                 as played_sum,
               (p.plan_2026 - IFNULL(SUM(c.contract_sum), 0)) as remainder,
               p.id
        FROM purchases p
                 LEFT JOIN contracts c ON p.id = c.purchase_id
        GROUP BY p.id \
        '''
df = pd.read_sql_query(query, conn)
conn.close()

if not df.empty:
    # Прячем ID от пользователя, но оставляем его в данных для работы
    display_df = df.drop(columns=['id'])
    display_df.columns = [
        "Подразделение", "Наименование", "Планируемая сумма, руб., 2026 ГОД", "Планируемая сумма, руб., 2027 ГОД", "Планируемая сумма, руб., 2028 ГОД",
        "ОКПД2", "КОСГУ", "Основание", "ИФО", "№ предл.", "№ графика", "Сыграно", "Остаток"
    ]

    st.data_editor(display_df, use_container_width=True, hide_index=True)

    # --- 4. ДОБАВЛЕНИЕ КОНТРАКТОВ (ВНИЗУ) ---
    st.divider()
    with st.expander("📝 Добавить контракт к существующей закупке"):
        # Создаем список для выбора: "ID: Наименование"
        choice_list = {f"ID {row['id']}: {row['name'][:50]}...": row['id'] for idx, row in df.iterrows()}
        selected_choice = st.selectbox("Выберите закупку", choice_list.keys())

        c_col1, c_col2 = st.columns(2)
        c_info = c_col1.text_input("Данные контракта (№, дата)")
        c_sum = c_col2.number_input("Сумма контракта", min_value=0.0)

        if st.button("Привязать контракт"):
            p_id = choice_list[selected_choice]
            conn = get_connection()
            conn.cursor().execute("INSERT INTO contracts (purchase_id, contract_info, contract_sum) VALUES (?,?,?)",
                                  (int(p_id), c_info, c_sum))
            conn.commit()
            conn.close()
            st.success("Контракт добавлен!")
            st.rerun()
else:
    st.info("Пока данных нет. Используйте панель слева 👈")