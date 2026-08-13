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


def capitalize_first_letter(text):
    """Делает первую букву строки заглавной, не меняя остальные слова"""
    if not text:
        return text
    text = str(text).strip()
    if not text:
        return text
    return text[0].upper() + text[1:]


# --- БАЗА ДАННЫХ (Версия 11) ---
def get_connection():
    return sqlite3.connect('procurement_v10.db')


def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS purchases
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     subdivision TEXT,
                     name TEXT,
                     year_placement INTEGER,
                     ifo TEXT,
                     okpd2 TEXT,
                     kosgu TEXT,
                     basis TEXT,
                     request_num TEXT,
                     plan_graph_num TEXT,
                     plan_2026 REAL,
                     plan_2027 REAL,
                     plan_2028 REAL,
                     nmck_2026 REAL DEFAULT 0,
                     nmck_2027 REAL DEFAULT 0
                 )''')

    c.execute('''CREATE TABLE IF NOT EXISTS budget_breakdown
                 (
                     purchase_id INTEGER,
                     year INTEGER,
                     ifo_name TEXT,
                     amount REAL
                 )''')

    c.execute('''CREATE TABLE IF NOT EXISTS nmck_breakdown
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     purchase_id INTEGER,
                     year INTEGER,
                     contract_name TEXT,
                     amount REAL,
                     onec_num TEXT
                 )''')

    try:
        c.execute("ALTER TABLE purchases ADD COLUMN nmck_2026 REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    try:
        c.execute("ALTER TABLE purchases ADD COLUMN nmck_2027 REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    try:
        c.execute("ALTER TABLE nmck_breakdown ADD COLUMN onec_num TEXT")
    except sqlite3.OperationalError:
        pass

    conn.commit()
    conn.close()


init_db()

IFO_SOURCES = ["ВБ", "ГЗ", "ОМС", "Прочее"]

st.title("📋 Реестр закупок")

# --- ФОРМА ВВОДА (Все поля сохранены!) ---
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
        okpd_raw = row2[0].text_input("ОКПД2 (вводите цифры)", placeholder="Напр: 123456")
        kosgu = row2[1].selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = row2[2].selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])

        # Строка 3: Номера
        row3 = st.columns(2)
        req_n = row3[0].text_input("Номер предложения на закупку")
        graph_n = row3[1].text_input("Номер план-графика")

        if st.form_submit_button("Сохранить позицию"):
            okpd_fixed = format_okpd(okpd_raw)
            name_fixed = capitalize_first_letter(name)
            conn = get_connection()
            conn.execute('''INSERT INTO purchases
                            (subdivision, name, year_placement, ifo, okpd2, kosgu, basis, request_num, plan_graph_num,
                             plan_2026, plan_2027, plan_2028, nmck_2026, nmck_2027)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0, 0, 0)''',
                         (sub, name_fixed, y_place, ", ".join(ifo_main), okpd_fixed, kosgu, basis, req_n, graph_n))
            conn.commit()
            conn.close()
            st.success(f"Добавлено! ОКПД2 отформатирован: {okpd_fixed}")
            st.rerun()

# --- ОСНОВНАЯ ТАБЛИЦА (Без 2028 года!) ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    st.subheader("Главный реестр")

    # Формируем таблицу БЕЗ 2028 года
    display_df = df[[
        "id", "subdivision", "name", "year_placement", "ifo",
        "okpd2", "kosgu", "basis", "request_num", "plan_graph_num",
        "plan_2026", "plan_2027", "nmck_2026", "nmck_2027"
    ]].copy()

    display_df.columns = [
        "ID", "Подразделение", "Наименование", "Год размещения", "ИФО",
        "ОКПД2", "КОСГУ", "Основание", "Номер предложения на закупку", "Номер план-графика",
        "Планируемая сумма, руб.; 2026 год", "Планируемая сумма, руб.; 2027 год",
        "Сумма по заявкам НМЦК 2026 год", "Сумма по заявкам НМЦК 2027 год"
    ]

    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        key="main_table_editor"
    )

    if st.button("💾 Применить правки из таблицы"):
        conn = get_connection()
        for _, row in edited_df.iterrows():
            updated_name = capitalize_first_letter(row["Наименование"])
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
                                nmck_2026=?,
                                nmck_2027=?
                            WHERE id = ?''',
                         (row["Подразделение"], updated_name, row["Год размещения"],
                          row["ИФО"], row["ОКПД2"], row["КОСГУ"], row["Основание"],
                          row["Номер предложения на закупку"], row["Номер план-графика"],
                          row["Планируемая сумма, руб.; 2026 год"], row["Планируемая сумма, руб.; 2027 год"],
                          row["Сумма по заявкам НМЦК 2026 год"], row["Сумма по заявкам НМЦК 2027 год"],
                          row["ID"]))
        conn.commit()
        conn.close()
        st.success("Все изменения в базе сохранены!")

    # --- УМНЫЙ СВЯЗАННЫЙ ПОИСК ЗАКУПКИ (НАЗВАНИЕ ➡️ ФИЛЬТР ID) ---
    st.divider()
    st.markdown("### 🎯 Выбор позиции для распределения ИФО и заявок НМЦК")

    col_s1, col_s2 = st.columns(2)

    with col_s1:
        all_names = list(df['name'].unique())
        chosen_name = st.selectbox("🔍 Поиск закупки по названию:", options=["-- Выберите --"] + all_names, key="search_name_box")

    # Фильтруем список доступных ID на основе выбранного названия
    if chosen_name != "-- Выберите --":
        matching_ids = list(df[df['name'] == chosen_name]['id'].unique())
    else:
        matching_ids = list(df['id'].unique())

    with col_s2:
        chosen_id = st.selectbox("🔢 Поиск по ID (индексу):", options=["-- Выберите ID --"] + matching_ids, key="search_id_box")

    # Определение итогового выбранного ID
    selected_id = None
    if chosen_id != "-- Выберите ID --":
        selected_id = int(chosen_id)
    elif chosen_name != "-- Выберите --":
        selected_id = int(matching_ids[0])

    # --- БЛОКИ ВВОДА ДАННЫХ (ИФО И НМЦК) ---
    if selected_id:
        row_data = df[df['id'] == selected_id].iloc[0]
        selected_name = row_data['name']
        sel_id = int(row_data['id'])

        # Определяем выбранные при создании ИФО (источники)
        raw_ifo_str = str(row_data['ifo']) if pd.notna(row_data['ifo']) else ""
        selected_sources = [s.strip() for s in raw_ifo_str.split(",") if s.strip()]
        # Если источники не указаны, показываем все базовые
        if not selected_sources:
            selected_sources = IFO_SOURCES

        st.success(f"📌 Выбрана позиция: **{selected_name}** (ID: {sel_id}, ИФО в позиции: {', '.join(selected_sources)})")

        # --- БЛОК 1: РАСПРЕДЕЛЕНИЕ БЮДЖЕТА ПО ИСТОЧНИКАМ (ТОЛЬКО ВЫБРАННЫЕ ИФО, 2026 и 2027 ГОДЫ) ---
        st.subheader("💰 Распределение бюджета по источникам (ИФО)")

        years = [2026, 2027]  # 2028 год убран!
        cols = st.columns(2)

        for i, year in enumerate(years):
            with cols[i]:
                st.markdown(f"**Бюджет на {year} год**")
                with st.container(border=True):
                    total_for_year = 0.0
                    for source in selected_sources:  # Покажем строго выбравшиеся источники!
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
                    conn = get_connection()
                    conn.execute(f"UPDATE purchases SET plan_{year} = ? WHERE id = ?", (total_for_year, sel_id))
                    conn.commit()
                    conn.close()

        # --- БЛОК 2: СУММА ПО ЗАЯВКАМ НМЦК ---
        st.divider()
        st.subheader("📝 Сумма по заявкам НМЦК (Контракты)")

        nmck_years = [2026, 2027]
        nmck_cols = st.columns(2)

        for i, year in enumerate(nmck_years):
            with nmck_cols[i]:
                st.markdown(f"**Заявки / Контракты на {year} год**")
                with st.container(border=True):
                    conn = get_connection()
                    contracts = conn.execute(
                        "SELECT id, contract_name, amount, onec_num FROM nmck_breakdown WHERE purchase_id=? AND year=?",
                        (sel_id, year)).fetchall()
                    conn.close()

                    total_nmck_year = sum(c[2] for c in contracts)
                    next_auto_num = f"№{len(contracts) + 1}"

                    if contracts:
                        for idx, (cid, cname, camount, onec_val) in enumerate(contracts, start=1):
                            c_col1, c_col2, c_col3, c_col4 = st.columns([1, 2, 2, 1])
                            c_col1.write(f"**№{idx}**")
                            c_col2.write(f"1С: {onec_val if onec_val else '—'}")
                            c_col3.write(f"{camount:,.2f} руб.")
                            if c_col4.button("❌", key=f"del_nmck_{cid}"):
                                conn = get_connection()
                                conn.execute("DELETE FROM nmck_breakdown WHERE id=?", (cid,))
                                conn.commit()
                                conn.close()
                                st.rerun()

                    st.caption(f"Добавить заявку ({next_auto_num}):")
                    add_col1, add_col2 = st.columns([2, 2])
                    new_onec = add_col1.text_input("Номер 1С", key=f"new_onec_{sel_id}_{year}")
                    new_camount = add_col2.number_input("Сумма, руб.", value=0.0, format="%.2f", key=f"new_camount_{sel_id}_{year}")

                    if st.button(f"➕ Добавить заявку {next_auto_num} ({year})", key=f"btn_add_nmck_{sel_id}_{year}"):
                        if new_camount > 0:
                            conn = get_connection()
                            conn.execute("INSERT INTO nmck_breakdown (purchase_id, year, contract_name, amount, onec_num) VALUES (?,?,?,?,?)",
                                         (sel_id, year, next_auto_num, new_camount, new_onec))
                            conn.commit()
                            conn.close()
                            st.rerun()

                    st.markdown(f"---")
                    st.write(f"**ИТОГО НМЦК {year}: {total_nmck_year:,.2f} руб.**")

                    conn = get_connection()
                    conn.execute(f"UPDATE purchases SET nmck_{year} = ? WHERE id = ?", (total_nmck_year, sel_id))
                    conn.commit()
                    conn.close()

        if st.button("🔄 Обновить итоговые суммы в таблице"):
            st.rerun()