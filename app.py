import streamlit as st
import pandas as pd
import sqlite3
import re

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")


# --- ФУНКЦИИ ФОРМАТИРОВАНИЯ ---
def format_okpd(raw_text):
    clean = re.sub(r'\D', '', raw_text)
    return ".".join(clean[i:i + 2] for i in range(0, len(clean), 2))


def to_money(val):
    try:
        return "{:.2f}".format(float(val.replace(',', '.').replace(' ', '')))
    except:
        return "0.00"


# --- БАЗА ДАННЫХ ---
def get_connection():
    return sqlite3.connect('procurement_v8.db')


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
    conn.commit()
    conn.close()


init_db()

st.title("📋 Реестр закупок")

# --- ФОРМА ВВОДА ---
with st.expander("➕ Добавить новую позицию", expanded=True):
    with st.form("new_entry", clear_on_submit=True):
        row1 = st.columns([2, 3, 1])
        sub = row1[0].selectbox("Подразделение", ["Автохозяйство", "Админ. отдел", "Лаборатория", "АХО"])
        name = row1[1].text_input("Наименование")
        year_pl = row1[2].selectbox("Год размещения", list(range(2026, 2031)))

        row2 = st.columns(3)
        # Текстовый ввод без кнопок +/-, авто-формат копеек при сохранении
        p26_raw = row2[0].text_input("Планируемая сумма, руб.; 2026 год", value="0.00")
        p27_raw = row2[1].text_input("Планируемая сумма, руб.; 2027 год", value="0.00")
        p28_raw = row2[2].text_input("Планируемая сумма, руб.; 2028 год", value="0.00")

        row3 = st.columns(4)
        okpd_raw = row3[0].text_input("ОКПД2", placeholder="123456")
        kosgu = row3[1].selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = row3[2].selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])

        ifo_list = ["00000000000000000130", "00000000000000000180", "Свой вариант..."]
        ifo_sel = row3[3].selectbox("ИФО", ifo_list)
        final_ifo = ifo_sel
        if ifo_sel == "Свой вариант...":
            final_ifo = st.text_input("Введите ИФО вручную")

        row4 = st.columns(2)
        req_n = row4[0].text_input("Номер предложения на закупку")
        graph_n = row4[1].text_input("Номер план-графика")

        if st.form_submit_button("Сохранить"):
            p26 = float(to_money(p26_raw))
            p27 = float(to_money(p27_raw))
            p28 = float(to_money(p28_raw))
            okpd_fixed = format_okpd(okpd_raw)

            conn = get_connection()
            conn.cursor().execute('''INSERT INTO purchases
                                     (subdivision, name, year_placement, plan_2026, plan_2027, plan_2028, okpd2, kosgu,
                                      basis, ifo, request_num, plan_graph_num)
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                  (sub, name, year_pl, p26, p27, p28, okpd_fixed, kosgu, basis, final_ifo, req_n,
                                   graph_n))
            conn.commit()
            conn.close()
            st.rerun()

# --- ТАБЛИЦА С ДАННЫМИ ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    st.subheader("Главный реестр (редактируемый)")

    # Переименовываем для красоты
    display_df = df.copy()
    display_df.columns = [
        "ID", "Подразделение", "Наименование", "Год размещения",
        "Планируемая сумма, руб.; 2026 год", "Планируемая сумма, руб.; 2027 год", "Планируемая сумма, руб.; 2028 год",
        "ОКПД2", "КОСГУ", "Основание", "ИФО", "№ предл.", "№ графика"
    ]

    # Полностью редактируемая таблица с выбором строк
    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        key="editor",
        selection_mode="single_row"
    )

    # Кнопка сохранения правок из таблицы
    if st.button("💾 Применить правки из таблицы"):
        conn = get_connection()
        cur = conn.cursor()
        for _, row in edited_df.iterrows():
            cur.execute('''UPDATE purchases
                           SET subdivision=?,
                               name=?,
                               year_placement=?,
                               plan_2026=?,
                               plan_2027=?,
                               plan_2028=?,
                               okpd2=?,
                               kosgu=?,
                               basis=?,
                               ifo=?,
                               request_num=?,
                               plan_graph_num=?
                           WHERE id = ?''',
                        (row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11],
                         row[12], row[0]))
        conn.commit()
        conn.close()
        st.success("Изменения в базе обновлены!")

    # --- БЛОК КОНТРАКТОВ (появляется при выборе строки) ---
    selected_indices = st.session_state.get("editor", {}).get("selection", {}).get("rows", [])
    if selected_indices:
        sel_idx = selected_indices[0]
        sel_row = df.iloc[sel_idx]

        st.divider()
        st.subheader(f"🔗 Контракты по позиции: {sel_row['name']}")

        c_left, c_right = st.columns(2)

        with c_left:
            with st.form("new_contract"):
                c_num = st.text_input("№ контракта")
                c_1s = st.text_input("№ 1С")
                c_sum = st.text_input("Сумма контракта", value="0.00")
                if st.form_submit_button("Привязать контракт"):
                    conn = get_connection()
                    conn.cursor().execute(
                        "INSERT INTO contracts (purchase_id, contract_num, one_s_num, contract_sum) VALUES (?,?,?,?)",
                        (int(sel_row['id']), c_num, c_1s, float(to_money(c_sum))))
                    conn.commit()
                    conn.close()
                    st.rerun()

        with c_right:
            conn = get_connection()
            contracts = pd.read_sql_query(
                f"SELECT contract_num, one_s_num, contract_sum FROM contracts WHERE purchase_id={sel_row['id']}", conn)
            conn.close()
            st.write("Список привязанных контрактов:")
            st.dataframe(contracts, use_container_width=True, hide_index=True)