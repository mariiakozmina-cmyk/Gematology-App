import streamlit as st
import pandas as pd
import sqlite3
import re
from datetime import datetime

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")


# --- ФУНКЦИИ ФОРМАТИРОВАНИЯ ---
def format_okpd(raw_text):
    """Добавляет точки через каждые 2 цифры"""
    clean = re.sub(r'\D', '', raw_text)
    return ".".join(clean[i:i + 2] for i in range(0, len(clean), 2))


def capitalize_first_letter(text):
    """Делает первой заглавной ТОЛЬКО первую букву строки"""
    if not text:
        return text
    text = str(text).strip()
    if not text:
        return text
    return text[0].upper() + text[1:]


# --- БАЗА ДАННЫХ (Версия 17: Учет последнего ДС + Прочее + Удаление) ---
def get_connection():
    return sqlite3.connect('procurement_v10.db')


def init_db():
    conn = get_connection()
    c = conn.cursor()

    # Таблица закупок
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
                     plan_2027 REAL DEFAULT 0,
                     plan_2028 REAL DEFAULT 0,
                     nmck_2027 REAL DEFAULT 0,
                     nmck_2028 REAL DEFAULT 0,
                     played_2027 REAL DEFAULT 0,
                     played_2028 REAL DEFAULT 0,
                     rem_2027 REAL DEFAULT 0,
                     rem_2028 REAL DEFAULT 0
                 )''')

    # Бюджет по ИФО
    c.execute('''CREATE TABLE IF NOT EXISTS budget_breakdown
                 (
                     purchase_id INTEGER,
                     year INTEGER,
                     ifo_name TEXT,
                     amount REAL
                 )''')

    # Позиционные заявки НМЦК
    c.execute('''CREATE TABLE IF NOT EXISTS nmck_applications
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     purchase_id INTEGER,
                     year INTEGER,
                     onec_num TEXT
                 )''')

    c.execute('''CREATE TABLE IF NOT EXISTS nmck_app_ifo_amounts
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     app_id INTEGER,
                     ifo_source TEXT,
                     amount REAL
                 )''')

    # Контракты
    c.execute('''CREATE TABLE IF NOT EXISTS contracts
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     purchase_id INTEGER,
                     year INTEGER,
                     contract_num TEXT,
                     onec_num TEXT,
                     contract_date TEXT,
                     comment
                     TEXT
                 )''')

    c.execute('''CREATE TABLE IF NOT EXISTS contract_ifo_amounts
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     contract_id INTEGER,
                     ifo_source TEXT,
                     amount REAL
                 )''')

    # Дополнительные соглашения (ДС)
    c.execute('''CREATE TABLE IF NOT EXISTS ds_agreements
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     contract_id INTEGER,
                     ds_num TEXT,
                     ds_date TEXT,
                     comment TEXT
                 )''')

    c.execute('''CREATE TABLE IF NOT EXISTS ds_ifo_amounts
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     ds_id INTEGER,
                     ifo_source TEXT,
                     amount REAL
                 )''')

    # Таблица пользовательских источников ИФО
    c.execute('''CREATE TABLE IF NOT EXISTS funding_sources
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     source_name TEXT UNIQUE
                 )''')

    conn.commit()
    conn.close()


init_db()


def get_all_ifo_sources():
    """Возвращает список базовых + пользовательских источников + 'Прочее'"""
    base_sources = ["ВБ", "ГЗ", "ОМС"]
    conn = get_connection()
    custom_sources = [row[0] for row in conn.execute("SELECT source_name FROM funding_sources").fetchall()]
    conn.close()

    # Объединяем без дубликатов, сохраняя порядок
    all_src = list(base_sources)
    for cs in custom_sources:
        if cs not in all_src:
            all_src.append(cs)
    all_src.append("Прочее")
    return all_src


st.title("📋 Реестр закупок")

# --- ФОРМА ВВОДА (С поддержкой 'Прочее') ---
with st.expander("➕ Добавить новую позицию", expanded=True):
    with st.form("new_entry", clear_on_submit=True):
        row1 = st.columns([2, 3, 1, 2])
        sub = row1[0].selectbox("Подразделение", ["Автохозяйство", "Админ. отдел", "Лаборатория", "АХО"])
        name = row1[1].text_input("Наименование")
        y_place = row1[2].selectbox("Год размещения", list(range(2027, 2032)))

        available_ifo = get_all_ifo_sources()
        ifo_main_selected = row1[3].multiselect("ИФО (источники)", available_ifo, default=[])

        # Поле для ввода нового источника, если выбрано "Прочее"
        custom_ifo_input = st.text_input("Если выбрали 'Прочее', укажите название нового источника ИФО:")

        row2 = st.columns([2, 1, 1])
        okpd_raw = row2[0].text_input("ОКПД2 (вводите цифры)", placeholder="Напр: 123456")
        kosgu = row2[1].selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = row2[2].selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])

        row3 = st.columns(2)
        req_n = row3[0].text_input("Номер предложения на закупку")
        graph_n = row3[1].text_input("Номер план-графика")

        submit_btn = st.form_submit_button("Сохранить позицию")

        if submit_btn:
            if not name or not name.strip():
                st.warning("⚠️ Пожалуйста, введите Наименование закупки перед сохранением!")
            else:
                # Обработка "Прочее"
                final_ifo_list = [s for s in ifo_main_selected if s != "Прочее"]
                if "Прочее" in ifo_main_selected and custom_ifo_input.strip():
                    new_src = custom_ifo_input.strip().upper()
                    conn = get_connection()
                    try:
                        conn.execute("INSERT INTO funding_sources (source_name) VALUES (?)", (new_src,))
                        conn.commit()
                    except sqlite3.IntegrityError:
                        pass
                    conn.close()
                    if new_src not in final_ifo_list:
                        final_ifo_list.append(new_src)

                okpd_fixed = format_okpd(okpd_raw)
                name_fixed = capitalize_first_letter(name)
                conn = get_connection()
                conn.execute('''INSERT INTO purchases
                                (subdivision, name, year_placement, ifo, okpd2, kosgu, basis, request_num, plan_graph_num,
                                 plan_2027, plan_2028, nmck_2027, nmck_2028, played_2027, played_2028, rem_2027, rem_2028)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0, 0, 0, 0, 0, 0)''',
                             (sub, name_fixed, y_place, ", ".join(final_ifo_list), okpd_fixed, kosgu, basis, req_n, graph_n))
                conn.commit()
                conn.close()
                st.success(f"Добавлено! Сохранено: {name_fixed}")
                st.rerun()

# --- ОСНОВНАЯ ТАБЛИЦА С МНОГОФУНКЦИОНАЛЬНЫМ ФИЛЬТРОМ (EXCEL-STYLE) ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    st.subheader("Главный реестр")

    # Панель быстрого фильтра Excel-Style
    st.caption("🔎 Быстрый фильтр реестра (как в Excel):")
    f_col1, f_col2 = st.columns([1, 2])
    filter_sub = f_col1.selectbox("Фильтр по Подразделению:", ["Все подразделения"] + list(df['subdivision'].unique()))
    filter_search = f_col2.text_input("Поиск по тексту в наименовании:")

    filtered_df = df.copy()
    if filter_sub != "Все подразделения":
        filtered_df = filtered_df[filtered_df['subdivision'] == filter_sub]
    if filter_search.strip():
        filtered_df = filtered_df[filtered_df['name'].str.contains(filter_search.strip(), case=False, na=False)]

    # Авто-расчет остатков
    filtered_df['rem_2027'] = filtered_df['plan_2027'] - filtered_df['nmck_2027'] - filtered_df['played_2027']
    filtered_df['rem_2028'] = filtered_df['plan_2028'] - filtered_df['nmck_2028'] - filtered_df['played_2028']

    display_df = filtered_df[[
        "id", "subdivision", "name", "year_placement", "ifo",
        "okpd2", "kosgu", "basis", "request_num", "plan_graph_num",
        "plan_2027", "plan_2028", "nmck_2027", "nmck_2028",
        "played_2027", "played_2028", "rem_2027", "rem_2028"
    ]].copy()

    display_df.columns = [
        "ID", "Подразделение", "Наименование", "Год размещения", "ИФО",
        "ОКПД2", "КОСГУ", "Основание", "Номер предложения на закупку", "Номер план-графика",
        "Планируемая сумма, руб.; 2027 год", "Планируемая сумма, руб.; 2028 год",
        "Сумма по заявкам НМЦК 2027 год", "Сумма по заявкам НМЦК 2028 год",
        "Сумма сыгранная 2027 год", "Сумма сыгранная 2028 год",
        "Остаток 2027 год", "Остаток 2028 год"
    ]

    column_configuration = {
        "ID": st.column_config.NumberColumn("ID", width="small"),
        "КОСГУ": st.column_config.TextColumn("КОСГУ", width="small"),
        "Основание": st.column_config.TextColumn("Основание", width="small"),
        "Год размещения": st.column_config.NumberColumn("Год размещения", width="small"),
        "Планируемая сумма, руб.; 2027 год": st.column_config.NumberColumn(format="%.2f руб."),
        "Планируемая сумма, руб.; 2028 год": st.column_config.NumberColumn(format="%.2f руб."),
        "Сумма по заявкам НМЦК 2027 год": st.column_config.NumberColumn(format="%.2f руб."),
        "Сумма по заявкам НМЦК 2028 год": st.column_config.NumberColumn(format="%.2f руб."),
        "Сумма сыгранная 2027 год": st.column_config.NumberColumn(format="%.2f руб."),
        "Сумма сыгранная 2028 год": st.column_config.NumberColumn(format="%.2f руб."),
        "Остаток 2027 год": st.column_config.NumberColumn(format="%.2f руб."),
        "Остаток 2028 год": st.column_config.NumberColumn(format="%.2f руб."),
    }

    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config=column_configuration,
        key="main_table_editor"
    )

    if st.button("💾 Применить правки из таблицы"):
        conn = get_connection()
        for _, row in edited_df.iterrows():
            updated_name = capitalize_first_letter(row["Наименование"])
            p27 = float(row["Планируемая сумма, руб.; 2027 год"])
            p28 = float(row["Планируемая сумма, руб.; 2028 год"])
            n27 = float(row["Сумма по заявкам НМЦК 2027 год"])
            n28 = float(row["Сумма по заявкам НМЦК 2028 год"])
            s27 = float(row["Сумма сыгранная 2027 год"])
            s28 = float(row["Сумма сыгранная 2028 год"])

            r27 = p27 - n27 - s27
            r28 = p28 - n28 - s28

            conn.execute('''UPDATE purchases
                            SET subdivision=?, name=?, year_placement=?, ifo=?, okpd2=?, kosgu=?, basis=?,
                                request_num=?, plan_graph_num=?, plan_2027=?, plan_2028=?, nmck_2027=?, nmck_2028=?,
                                played_2027=?, played_2028=?, rem_2027=?, rem_2028=?
                            WHERE id = ?''',
                         (row["Подразделение"], updated_name, row["Год размещения"],
                          row["ИФО"], row["ОКПД2"], row["КОСГУ"], row["Основание"],
                          row["Номер предложения на закупку"], row["Номер план-графика"],
                          p27, p28, n27, n28, s27, s28, r27, r28, row["ID"]))
        conn.commit()
        conn.close()
        st.success("Все изменения в базе сохранены!")

    # --- УМНЫЙ СВЯЗАННЫЙ ПОИСК ЗАКУПКИ ---
    st.divider()
    st.markdown("### 🎯 Выбор позиции для детализации")

    col_s1, col_s2 = st.columns(2)

    with col_s1:
        all_names = list(df['name'].unique())
        chosen_name = st.selectbox("🔍 Поиск закупки по названию:", options=["-- Выберите --"] + all_names, key="search_name_box")

    if chosen_name != "-- Выберите --":
        matching_ids = list(df[df['name'] == chosen_name]['id'].unique())
    else:
        matching_ids = list(df['id'].unique())

    with col_s2:
        chosen_id = st.selectbox("🔢 Поиск по ID (индексу):", options=["-- Выберите ID --"] + matching_ids, key="search_id_box")

    selected_id = None
    if chosen_id != "-- Выберите ID --":
        selected_id = int(chosen_id)
    elif chosen_name != "-- Выберите --":
        selected_id = int(matching_ids[0])

    # --- БЛОКИ ВВОДА ДАННЫХ ---
    if selected_id:
        row_data = df[df['id'] == selected_id].iloc[0]
        selected_name = row_data['name']
        sel_id = int(row_data['id'])

        raw_ifo_str = str(row_data['ifo']) if pd.notna(row_data['ifo']) else ""
        selected_sources_for_purchase = [s.strip() for s in raw_ifo_str.split(",") if s.strip()]
        if not selected_sources_for_purchase:
            selected_sources_for_purchase = ["ВБ", "ГЗ", "ОМС"]

        st.success(f"📌 Выбрана позиция: **{selected_name}** (ID: {sel_id}, ИФО: {', '.join(selected_sources_for_purchase)})")

        # --- БЛОК 1: РАСПРЕДЕЛЕНИЕ БЮДЖЕТА ПО ИСТОЧНИКАМ (ИФО) ---
        with st.expander("💰 Распределение бюджета по источникам (ИФО)", expanded=True):
            years = [2027, 2028]
            cols = st.columns(2)

            for i, year in enumerate(years):
                with cols[i]:
                    st.markdown(f"**Бюджет на {year} год**")
                    with st.container(border=True):
                        total_for_year = 0.0
                        for source in selected_sources_for_purchase:
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
        with st.expander("📝 Сумма по заявкам НМЦК", expanded=False):
            nmck_years = [2027, 2028]
            nmck_cols = st.columns(2)

            for i, year in enumerate(nmck_years):
                with nmck_cols[i]:
                    st.markdown(f"**Заявки на {year} год**")
                    with st.container(border=True):
                        conn = get_connection()
                        applications = conn.execute(
                            "SELECT id, onec_num FROM nmck_applications WHERE purchase_id=? AND year=?",
                            (sel_id, year)).fetchall()

                        total_nmck_year = 0.0

                        if applications:
                            for app_idx, (app_id, onec_val) in enumerate(applications):
                                st.markdown(f"**Номер 1С: {onec_val if onec_val else '—'}**")
                                app_ifo_amounts = conn.execute(
                                    "SELECT id, ifo_source, amount FROM nmck_app_ifo_amounts WHERE app_id=?",
                                    (app_id,)).fetchall()

                                current_app_total = sum(item[2] for item in app_ifo_amounts)
                                total_nmck_year += current_app_total

                                for item_id, ifo_src, amount in app_ifo_amounts:
                                    st.write(f"- {ifo_src}: {amount:,.2f} руб.")

                                st.markdown(f"**Остаток после этой заявки:**")
                                for source in selected_sources_for_purchase:
                                    budget_ifo = conn.execute(
                                        "SELECT amount FROM budget_breakdown WHERE purchase_id=? AND year=? AND ifo_name=?",
                                        (sel_id, year, source)).fetchone()
                                    budget_ifo = budget_ifo[0] if budget_ifo else 0.0

                                    current_nmck_ifo_sum = conn.execute(
                                        "SELECT SUM(t2.amount) FROM nmck_applications t1 JOIN nmck_app_ifo_amounts t2 ON t1.id = t2.app_id WHERE t1.purchase_id=? AND t1.year=? AND t2.ifo_source=?",
                                        (sel_id, year, source)).fetchone()
                                    current_nmck_ifo_sum = current_nmck_ifo_sum[0] if current_nmck_ifo_sum[0] is not None else 0.0

                                    st.write(f"- {source}: {budget_ifo - current_nmck_ifo_sum:,.2f} руб.")

                                c_col1, c_col2 = st.columns([1, 1])
                                if c_col1.button("✏️", key=f"edit_nmck_btn_{app_id}"):
                                    st.session_state[f"editing_nmck_{app_id}"] = not st.session_state.get(f"editing_nmck_{app_id}", False)
                                if c_col2.button("❌", key=f"del_nmck_app_{app_id}"):
                                    conn.execute("DELETE FROM nmck_applications WHERE id=?", (app_id,))
                                    conn.execute("DELETE FROM nmck_app_ifo_amounts WHERE app_id=?", (app_id,))
                                    conn.commit()
                                    st.rerun()

                                if st.session_state.get(f"editing_nmck_{app_id}", False):
                                    with st.form(key=f"form_edit_nmck_{app_id}"):
                                        e_onec = st.text_input("Изменить Номер 1С", value=onec_val if onec_val else "")
                                        st.markdown("**Изменить суммы по ИФО:**")
                                        edited_ifo_amounts = {}
                                        for item_id, ifo_src, amount in app_ifo_amounts:
                                            edited_ifo_amounts[ifo_src] = st.number_input(f"{ifo_src}", value=float(amount), format="%.2f", key=f"edit_app_ifo_{item_id}")

                                        if st.form_submit_button("Сохранить изменения заявки"):
                                            conn.execute("UPDATE nmck_applications SET onec_num=? WHERE id=?", (e_onec, app_id))
                                            for ifo_src, amount in edited_ifo_amounts.items():
                                                conn.execute("UPDATE nmck_app_ifo_amounts SET amount=? WHERE app_id=? AND ifo_source=?", (amount, app_id, ifo_src))
                                            conn.commit()
                                            st.session_state[f"editing_nmck_{app_id}"] = False
                                            st.rerun()
                                st.markdown("---")

                        st.caption("Добавить новую заявку:")
                        new_onec = st.text_input("Номер 1С", key=f"new_onec_app_{sel_id}_{year}")
                        st.markdown("**Суммы по источникам:**")
                        new_app_ifo_amounts = {}
                        for source in selected_sources_for_purchase:
                            new_app_ifo_amounts[source] = st.number_input(f"{source}", value=0.0, format="%.2f", key=f"new_app_ifo_{sel_id}_{year}_{source}")

                        if st.button(f"➕ Добавить заявку ({year})", key=f"btn_add_nmck_app_{sel_id}_{year}"):
                            total_new_app_amount = sum(new_app_ifo_amounts.values())
                            if total_new_app_amount > 0:
                                cursor = conn.execute("INSERT INTO nmck_applications (purchase_id, year, onec_num) VALUES (?,?,?)",
                                                     (sel_id, year, new_onec))
                                new_app_id = cursor.lastrowid
                                for ifo_src, amount in new_app_ifo_amounts.items():
                                    if amount > 0:
                                        conn.execute("INSERT INTO nmck_app_ifo_amounts (app_id, ifo_source, amount) VALUES (?,?,?)",
                                                     (new_app_id, ifo_src, amount))
                                conn.commit()
                                st.rerun()

                        st.markdown(f"---")
                        st.write(f"**ИТОГО НМЦК {year}: {total_nmck_year:,.2f} руб.**")

                        conn.execute(f"UPDATE purchases SET nmck_{year} = ? WHERE id = ?", (total_nmck_year, sel_id))
                        conn.commit()
                        conn.close()

        # --- БЛОК 3: СУММА СЫГРАННАЯ (УЧЕТ ТОЛЬКО ПОСЛЕДНЕГО ДС) ---
        with st.expander("🤝 Контракты (Сумма сыгранная и ДС)", expanded=False):
            cnt_years = [2027, 2028]
            cnt_cols = st.columns(2)

            for i, year in enumerate(cnt_years):
                with cnt_cols[i]:
                    st.markdown(f"**Контракты на {year} год**")
                    with st.container(border=True):
                        conn = get_connection()
                        req_1c_tuples = conn.execute(
                            "SELECT DISTINCT onec_num FROM nmck_applications WHERE purchase_id=? AND year=? AND onec_num IS NOT NULL AND onec_num != ''",
                            (sel_id, year)).fetchall()
                        avail_1c_from_apps = [t[0] for t in req_1c_tuples]

                        contracts = conn.execute(
                            "SELECT id, contract_num, onec_num, contract_date, comment FROM contracts WHERE purchase_id=? AND year=?",
                            (sel_id, year)).fetchall()

                        next_cnt_num = f"№{len(contracts) + 1}"
                        total_played_year = 0.0

                        if contracts:
                            for idx, (contract_id, cnum, c1c, cdate, ccomm) in enumerate(contracts, start=1):
                                # Базовые суммы основного контракта
                                contract_ifo_amounts = conn.execute(
                                    "SELECT id, ifo_source, amount FROM contract_ifo_amounts WHERE contract_id=?",
                                    (contract_id,)).fetchall()
                                base_contract_ifo_map = {item[1]: item[2] for item in contract_ifo_amounts}

                                # Проверяем, есть ли ДС у контракта
                                last_ds = conn.execute(
                                    "SELECT id, ds_num, ds_date, comment FROM ds_agreements WHERE contract_id=? ORDER BY id DESC LIMIT 1",
                                    (contract_id,)).fetchone()

                                effective_contract_ifo_map = {}

                                if last_ds:
                                    # ЛОГИКА: Если есть ДС — берем суммы ТОЛЬКО из самого последнего ДС!
                                    last_ds_id = last_ds[0]
                                    ds_ifo_items = conn.execute(
                                        "SELECT ifo_source, amount FROM ds_ifo_amounts WHERE ds_id=?",
                                        (last_ds_id,)).fetchall()
                                    last_ds_map = {item[0]: item[1] for item in ds_ifo_items}

                                    for source in selected_sources_for_purchase:
                                        # Если в последнем ДС сумма указана, берем её, иначе — базовую контракта
                                        effective_contract_ifo_map[source] = last_ds_map.get(source, base_contract_ifo_map.get(source, 0.0))
                                else:
                                    # Если ДС нет — берем суммы основного контракта
                                    for source in selected_sources_for_purchase:
                                        effective_contract_ifo_map[source] = base_contract_ifo_map.get(source, 0.0)

                                full_contract_total = sum(effective_contract_ifo_map.values())
                                total_played_year += full_contract_total

                                st.markdown(f"**Контракт №{idx}** (1С: **{c1c}**, Дата: {cdate})")
                                for source, effective_amount in effective_contract_ifo_map.items():
                                    st.write(f"- {source}: {effective_amount:,.2f} руб.")
                                if last_ds:
                                    st.caption(f"*(Учтены актуальные данные из последнего ДС №{last_ds[1]})*")
                                st.write(f"**ИТОГО КОНТРАКТА: {full_contract_total:,.2f} руб.**")
                                st.caption(f"💬 {ccomm}" if ccomm else "")

                                # Расчет остатка по ИФО после этого контракта
                                st.markdown(f"**Остаток после этого контракта:**")
                                for source in selected_sources_for_purchase:
                                    budget_ifo = conn.execute(
                                        "SELECT amount FROM budget_breakdown WHERE purchase_id=? AND year=? AND ifo_name=?",
                                        (sel_id, year, source)).fetchone()
                                    budget_ifo = budget_ifo[0] if budget_ifo else 0.0

                                    nmck_ifo_sum = conn.execute(
                                        "SELECT SUM(t2.amount) FROM nmck_applications t1 JOIN nmck_app_ifo_amounts t2 ON t1.id = t2.app_id WHERE t1.purchase_id=? AND t1.year=? AND t2.ifo_source=?",
                                        (sel_id, year, source)).fetchone()
                                    nmck_ifo_sum = nmck_ifo_sum[0] if nmck_ifo_sum[0] is not None else 0.0

                                    st.write(f"- {source}: {budget_ifo - nmck_ifo_sum - effective_contract_ifo_map.get(source, 0.0):,.2f} руб.")

                                k_col1, k_col2 = st.columns([1, 1])
                                if k_col1.button("✏️", key=f"edit_cnt_btn_{contract_id}"):
                                    st.session_state[f"editing_cnt_{contract_id}"] = not st.session_state.get(f"editing_cnt_{contract_id}", False)
                                if k_col2.button("❌", key=f"del_cnt_{contract_id}"):
                                    conn.execute("DELETE FROM contracts WHERE id=?", (contract_id,))
                                    conn.execute("DELETE FROM contract_ifo_amounts WHERE contract_id=?", (contract_id,))
                                    conn.execute("DELETE FROM ds_agreements WHERE contract_id=?", (contract_id,))
                                    conn.execute("DELETE FROM ds_ifo_amounts WHERE ds_id IN (SELECT id FROM ds_agreements WHERE contract_id=?);", (contract_id,))
                                    conn.commit()
                                    st.rerun()

                                if st.session_state.get(f"editing_cnt_{contract_id}", False):
                                    with st.form(key=f"form_edit_cnt_{contract_id}"):
                                        ec_1c = st.selectbox("Изменить 1С", options=avail_1c_from_apps, index=avail_1c_from_apps.index(c1c) if c1c in avail_1c_from_apps else 0)
                                        ec_date = st.date_input("Изменить Дату", value=datetime.strptime(cdate, "%d-%m-%Y").date(), format="DD-MM-YYYY")
                                        st.markdown("**Изменить суммы по ИФО (основа контракта):**")
                                        edited_contract_ifo_amounts = {}
                                        for item_id, ifo_src, amount in contract_ifo_amounts:
                                            edited_contract_ifo_amounts[ifo_src] = st.number_input(f"{ifo_src}", value=float(amount), format="%.2f", key=f"edit_contract_ifo_{item_id}")
                                        ec_comm = st.text_input("Изменить Комментарий", value=ccomm if ccomm else "")
                                        if st.form_submit_button("Сохранить изменения контракта"):
                                            conn.execute("UPDATE contracts SET onec_num=?, contract_date=?, comment=? WHERE id=?",
                                                         (ec_1c, ec_date.strftime("%d-%m-%Y"), ec_comm, contract_id))
                                            for ifo_src, amount in edited_contract_ifo_amounts.items():
                                                conn.execute("UPDATE contract_ifo_amounts SET amount=? WHERE contract_id=? AND ifo_source=?", (amount, contract_id, ifo_src))
                                            conn.commit()
                                            st.session_state[f"editing_cnt_{contract_id}"] = False
                                            st.rerun()
                                st.markdown("---")

                                # --- БЛОК ДОПОЛНИТЕЛЬНЫХ СОГЛАШЕНИЙ (ДС) ---
                                with st.expander(f"📑 Доп. соглашения (ДС) к контракту №{idx}", expanded=False):
                                    ds_items = conn.execute("SELECT id, ds_num, ds_date, comment FROM ds_agreements WHERE contract_id=?", (contract_id,)).fetchall()

                                    if ds_items:
                                        for ds_rec_id, ds_num_val, ds_date_val, ds_comm_val in ds_items:
                                            st.markdown(f"**ДС №{ds_num_val}** (Дата: {ds_date_val})")
                                            ds_ifo_items = conn.execute(
                                                "SELECT id, ifo_source, amount FROM ds_ifo_amounts WHERE ds_id=?",
                                                (ds_rec_id,)).fetchall()
                                            for ds_ifo_id, ds_ifo_src, ds_ifo_amt in ds_ifo_items:
                                                st.write(f"- {ds_ifo_src}: {ds_ifo_amt:,.2f} руб.")
                                            st.caption(f"💬 {ds_comm_val}" if ds_comm_val else "")

                                            ds_b1, ds_b2 = st.columns([1,1])
                                            if ds_b1.button("✏️", key=f"edit_ds_btn_{ds_rec_id}"):
                                                st.session_state[f"editing_ds_{ds_rec_id}"] = not st.session_state.get(f"editing_ds_{ds_rec_id}", False)
                                            if ds_b2.button("❌", key=f"del_ds_{ds_rec_id}"):
                                                conn.execute("DELETE FROM ds_agreements WHERE id=?", (ds_rec_id,))
                                                conn.execute("DELETE FROM ds_ifo_amounts WHERE ds_id=?", (ds_rec_id,))
                                                conn.commit()
                                                st.rerun()

                                            if st.session_state.get(f"editing_ds_{ds_rec_id}", False):
                                                with st.form(key=f"form_edit_ds_{ds_rec_id}"):
                                                    e_ds_num = st.text_input("Изменить № ДС", value=ds_num_val)
                                                    e_ds_date = st.date_input("Изменить Дату ДС", value=datetime.strptime(ds_date_val, "%d-%m-%Y").date(), format="DD-MM-YYYY")
                                                    st.markdown("**Изменить суммы по ИФО:**")
                                                    edited_ds_ifo_amounts = {}
                                                    for ds_ifo_id, ds_ifo_src, ds_ifo_amt in ds_ifo_items:
                                                        edited_ds_ifo_amounts[ds_ifo_src] = st.number_input(f"{ds_ifo_src}", value=float(ds_ifo_amt), format="%.2f", key=f"edit_ds_ifo_{ds_ifo_id}")
                                                    e_ds_comm = st.text_input("Изменить Комментарий", value=ds_comm_val if ds_comm_val else "")
                                                    if st.form_submit_button("Сохранить изменения ДС"):
                                                        conn.execute("UPDATE ds_agreements SET ds_num=?, ds_date=?, comment=? WHERE id=?",
                                                                     (e_ds_num, e_ds_date.strftime("%d-%m-%Y"), e_ds_comm, ds_rec_id))
                                                        for ifo_src, amount in edited_ds_ifo_amounts.items():
                                                            conn.execute("UPDATE ds_ifo_amounts SET amount=? WHERE ds_id=? AND ifo_source=?", (amount, ds_rec_id, ifo_src))
                                                        conn.commit()
                                                        st.session_state[f"editing_ds_{ds_rec_id}"] = False
                                                        st.rerun()
                                            st.markdown("---")

                                    st.caption("Добавить новое Дополнительное Соглашение (ДС):")
                                    new_ds_num = st.text_input("№ ДС", key=f"new_ds_num_{contract_id}")
                                    new_ds_date = st.date_input("Дата ДС", format="DD-MM-YYYY", key=f"new_ds_date_{contract_id}")
                                    st.markdown("**Суммы скорректированного контракта по источникам:**")
                                    new_ds_ifo_amounts = {}
                                    for source in selected_sources_for_purchase:
                                        # Подставляем последние значения для удобства
                                        new_ds_ifo_amounts[source] = st.number_input(f"{source}", value=float(effective_contract_ifo_map.get(source, 0.0)), format="%.2f", key=f"new_ds_ifo_{contract_id}_{source}")
                                    new_ds_comm = st.text_input("Комментарий к ДС", key=f"new_ds_comm_{contract_id}")

                                    if st.button(f"➕ Добавить ДС к контракту №{idx}", key=f"btn_add_ds_{contract_id}"):
                                        cursor = conn.execute("INSERT INTO ds_agreements (contract_id, ds_num, ds_date, comment) VALUES (?,?,?,?)",
                                                             (contract_id, new_ds_num if new_ds_num else "ДС", new_ds_date.strftime("%d-%m-%Y"), new_ds_comm))
                                        new_ds_id = cursor.lastrowid
                                        for ifo_src, amount in new_ds_ifo_amounts.items():
                                            conn.execute("INSERT INTO ds_ifo_amounts (ds_id, ifo_source, amount) VALUES (?,?,?)",
                                                         (new_ds_id, ifo_src, amount))
                                        conn.commit()
                                        st.rerun()

                                st.markdown("---")

                        st.caption(f"Добавить новый контракт ({next_cnt_num}):")

                        if not avail_1c_from_apps:
                            st.info("💡 *Сначала добавьте хотя бы одну заявку с Номером 1С выше.*")
                        else:
                            ca1, ca2, ca3 = st.columns([2, 2, 2])
                            c_1c_sel = ca1.selectbox("Номер 1С (из заявок)", options=avail_1c_from_apps, key=f"c_1c_{sel_id}_{year}")
                            c_date_val = ca2.date_input("Дата контракта", format="DD-MM-YYYY", key=f"c_date_{sel_id}_{year}")
                            c_comm_val = ca3.text_input("Комментарий", key=f"c_comm_{sel_id}_{year}")

                            st.markdown("**Суммы по источникам (основа контракта):**")
                            selected_app_id_for_1c = conn.execute("SELECT id FROM nmck_applications WHERE purchase_id=? AND year=? AND onec_num=?", (sel_id, year, c_1c_sel)).fetchone()
                            selected_app_ifo_amounts = {}
                            if selected_app_id_for_1c:
                                app_ifo_data = conn.execute("SELECT ifo_source, amount FROM nmck_app_ifo_amounts WHERE app_id=?", (selected_app_id_for_1c[0],)).fetchall()
                                selected_app_ifo_amounts = {item[0]: item[1] for item in app_ifo_data}

                            new_contract_ifo_amounts = {}
                            for source in selected_sources_for_purchase:
                                new_contract_ifo_amounts[source] = st.number_input(f"{source}", value=selected_app_ifo_amounts.get(source, 0.0), format="%.2f", key=f"new_cnt_ifo_{sel_id}_{year}_{source}")

                            if st.button(f"➕ Добавить контракт {next_cnt_num} ({year})", key=f"btn_add_cnt_{sel_id}_{year}"):
                                total_new_contract_amount = sum(new_contract_ifo_amounts.values())
                                if total_new_contract_amount > 0:
                                    cursor = conn.execute(
                                        "INSERT INTO contracts (purchase_id, year, contract_num, onec_num, contract_date, comment) VALUES (?,?,?,?,?,?)",
                                        (sel_id, year, next_cnt_num, c_1c_sel, c_date_val.strftime("%d-%m-%Y"), c_comm_val))
                                    new_contract_id = cursor.lastrowid
                                    for ifo_src, amount in new_contract_ifo_amounts.items():
                                        if amount > 0:
                                            conn.execute("INSERT INTO contract_ifo_amounts (contract_id, ifo_source, amount) VALUES (?,?,?)",
                                                         (new_contract_id, ifo_src, amount))
                                    conn.commit()
                                    st.rerun()

                        st.markdown("---")
                        st.write(f"**ИТОГО СЫГРАНО {year}: {total_played_year:,.2f} руб.**")

                        conn.execute(f"UPDATE purchases SET played_{year} = ? WHERE id = ?", (total_played_year, sel_id))
                        conn.commit()
                        conn.close()

        if st.button("🔄 Обновить итоговые суммы в таблице"):
            st.rerun()

# --- БЛОК БЕЗОПАСНОГО УДАЛЕНИЯ ПОЗИЦИИ ---
st.divider()
with st.expander("🗑️ Удаление позиции из реестра"):
    conn = get_connection()
    all_purchases_to_delete = conn.execute("SELECT id, name, subdivision FROM purchases").fetchall()
    conn.close()

    if not all_purchases_to_delete:
        st.info("Реестр пуст, нечего удалять.")
    else:
        delete_options = [f"ID: {p[0]} | {p[1]} ({p[2]})" for p in all_purchases_to_delete]
        selected_to_delete_str = st.selectbox("Выберите закупку для ПОЛНОГО удаления:", delete_options)
        del_id = int(selected_to_delete_str.split("ID: ")[1].split(" |")[0])

        st.warning(f"⚠️ Внимание! Вместе с закупкой будут удалены ВСЕ связанные заявки, контракты и ДС.")
        confirm_del = st.checkbox("Я понимаю, что действие необратимо", key="confirm_del_check")

        if st.button("🔴 Безопасно удалить закупку", disabled=not confirm_del):
            conn = get_connection()
            # 1. Удаляем ДС и их детализацию
            conn.execute("DELETE FROM ds_ifo_amounts WHERE ds_id IN (SELECT id FROM ds_agreements WHERE contract_id IN (SELECT id FROM contracts WHERE purchase_id=?));", (del_id,))
            conn.execute("DELETE FROM ds_agreements WHERE contract_id IN (SELECT id FROM contracts WHERE purchase_id=?);", (del_id,))

            # 2. Удаляем Контракты и их детализацию
            conn.execute("DELETE FROM contract_ifo_amounts WHERE contract_id IN (SELECT id FROM contracts WHERE purchase_id=?);", (del_id,))
            conn.execute("DELETE FROM contracts WHERE purchase_id=?;", (del_id,))

            # 3. Удаляем Заявки и их детализацию
            conn.execute("DELETE FROM nmck_app_ifo_amounts WHERE app_id IN (SELECT id FROM nmck_applications WHERE purchase_id=?);", (del_id,))
            conn.execute("DELETE FROM nmck_applications WHERE purchase_id=?;", (del_id,))

            # 4. Удаляем Бюджет
            conn.execute("DELETE FROM budget_breakdown WHERE purchase_id=?;", (del_id,))

            # 5. Удаляем саму закупку
            conn.execute("DELETE FROM purchases WHERE id=?;", (del_id,))

            conn.commit()
            conn.close()
            st.success("Позиция и все связанные данные успешно удалены!")
            st.rerun()