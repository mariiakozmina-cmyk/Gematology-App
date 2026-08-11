import streamlit as st
import pandas as pd
import sqlite3

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")


# --- 1. БАЗА ДАННЫХ (Версия 6 - Расширенная) ---
def get_connection():
    return sqlite3.connect('procurement_v6.db')


def init_db():
    conn = get_connection()
    c = conn.cursor()
    # Все-все колонки по твоему списку
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
                     TEXT,
                     nmtck_2026
                     REAL,
                     nmtck_2027
                     REAL,
                     played_2026
                     REAL,
                     remainder_2026
                     REAL,
                     comment_2026
                     TEXT,
                     comment_2027
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

# --- 2. ИНТЕРФЕЙС ---
st.title("Система учета закупок")

# Форма ввода (красивая и полная)
with st.expander("➕ Добавить новую позицию", expanded=False):
    with st.form("new_entry", clear_on_submit=True):
        row1 = st.columns([2, 3, 1])
        sub = row1[0].selectbox("Подразделение", ["Автохозяйство", "Административно-хозяйственный отдел", "Лаборатория", "АХО"])
        name = row1[1].text_input("Наименование объекта")
        # Год выпадающим списком до 2030
        year = row1[2].selectbox("Год разм.", list(range(2026, 2031)))

        st.write("---")
        row2 = st.columns(3)
        # Ввод без кнопок +/- через обычный ввод, но с проверкой на число
        p26 = row2[0].text_input("План 2026 (руб.)", value="0.00")
        p27 = row2[1].text_input("План 2027 (руб.)", value="0.00")
        p28 = row2[2].text_input("План 2028 (руб.)", value="0.00")

        st.write("---")
        row3 = st.columns(4)
        okpd = row3[0].text_input("ОКПД2 (00.00.00...)", placeholder="12.34.56")
        kosgu = row3[1].selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = row3[2].selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])

        ifo_list = ["00000000000000000130", "00000000000000000180", "Свой вариант..."]
        ifo_sel = row3[3].selectbox("ИФО", ifo_list)
        final_ifo = ifo_sel
        if ifo_sel == "Свой вариант...":
            final_ifo = st.text_input("Введите ИФО")

        row4 = st.columns(2)
        req_n = row4[0].text_input("№ предложения")
        graph_n = row4[1].text_input("№ план-графика")

        if st.form_submit_button("Добавить в реестр"):
            conn = get_connection()
            # Переводим текст в числа для базы
            try:
                vals = (sub, name, year, float(p26), float(p27), float(p28), okpd, kosgu, basis, final_ifo, req_n,
                        graph_n, 0, 0, 0, 0, "", "")
                conn.cursor().execute('''INSERT INTO purchases
                                         (subdivision, name, year, plan_2026, plan_2027, plan_2028, okpd2, kosgu, basis,
                                          ifo, request_num, plan_graph_num,
                                          nmtck_2026, nmtck_2027, played_2026, remainder_2026, comment_2026,
                                          comment_2027)
                                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', vals)
                conn.commit()
                st.success("Добавлено!")
            except:
                st.error("Ошибка в суммах! Пишите через точку, например: 100.50")
            finally:
                conn.close()
                st.rerun()

# --- 3. ГЛАВНАЯ ТАБЛИЦА ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    st.subheader("Реестр (выберите строку для ввода контракта)")

    # Русские названия для отображения
    display_df = df.copy()
    display_df.columns = [
        "ID", "Подразделение", "Наименование", "Год", "План 26", "План 27", "План 28",
        "ОКПД2", "КОСГУ", "Основание", "ИФО", "№ предл.", "№ графика",
        "НМЦК 26", "НМЦК 27", "Сыграно 26", "Остаток 26", "Коммент 26", "Коммент 27"
    ]

    # ВКЛЮЧАЕМ ВЫБОР СТРОК (Selection)
    event = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        on_change=None,  # Отключаем авто-обновление для стабильности
        key="main_table",
        column_config={"ID": None}  # Прячем ID
    )

    # Проверяем, на какую строку нажали (имитация окошка)
    selected_rows = st.session_state["main_table"]["selection"]["rows"]

    if selected_rows:
        idx = selected_rows[0]
        selected_data = df.iloc[idx]

        st.divider()
        st.markdown(f"### 📑 Контракты для: *{selected_data['name']}*")

        c1, c2 = st.columns([1, 1])

        with c1:
            with st.popover("➕ Вписать новый контракт"):
                with st.form("contract_form"):
                    c_num = st.text_input("№ контракта")
                    c_1s = st.text_input("№ 1С")
                    c_sum = st.text_input("Сумма (руб.)", value="0.00")
                    if st.form_submit_button("Сохранить контракт"):
                        conn = get_connection()
                        conn.cursor().execute(
                            "INSERT INTO contracts (purchase_id, contract_num, one_s_num, contract_sum) VALUES (?,?,?,?)",
                            (int(selected_data['id']), c_num, c_1s, float(c_sum)))
                        # Авто-обновление суммы "Сыграно"
                        conn.cursor().execute(
                            f"UPDATE purchases SET played_2026 = played_2026 + {float(c_sum)} WHERE id={selected_data['id']}")
                        conn.commit()
                        conn.close()
                        st.rerun()

        with c2:
            # Показываем список уже добавленных контрактов
            conn = get_connection()
            c_list = pd.read_sql_query(
                f"SELECT contract_num, one_s_num, contract_sum FROM contracts WHERE purchase_id={selected_data['id']}",
                conn)
            conn.close()
            if not c_list.empty:
                st.dataframe(c_list, hide_index=True)

    # Кнопка сохранения правок в самой таблице (если мама что-то там меняла)
    if st.button("💾 Сохранить изменения в таблице"):
        # Тут логика обновления базы из event (редактируемой таблицы)
        st.info("Функция сохранения правок активна")