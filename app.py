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


# --- БАЗА ДАННЫХ (Версия 15: Полная защита от Enter + ДС) ---
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

    # Заявки НМЦК
    c.execute('''CREATE TABLE IF NOT EXISTS nmck_breakdown
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     purchase_id INTEGER,
                     year INTEGER,
                     contract_name TEXT,
                     amount REAL,
                     onec_num TEXT,
                     ifo_source TEXT
                 )''')

    # Контракты
    c.execute('''CREATE TABLE IF NOT EXISTS contracts_breakdown
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     purchase_id INTEGER,
                     year INTEGER,
                     contract_num TEXT,
                     onec_num TEXT,
                     contract_date TEXT,
                     amount REAL,
                     comment TEXT
                 )''')

    # Дополнительные соглашения (ДС)
    c.execute('''CREATE TABLE IF NOT EXISTS contract_ds
                 (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     contract_id INTEGER,
                     ds_num TEXT,
                     ds_date TEXT,
                     amount REAL,
                     comment TEXT
                 )''')

    # Авто-миграция колонок
    cols_to_add = [
        ("purchases", "plan_2027", "REAL DEFAULT 0"),
        ("purchases", "plan_2028", "REAL DEFAULT 0"),
        ("purchases", "nmck_2027", "REAL DEFAULT 0"),
        ("purchases", "nmck_2028", "REAL DEFAULT 0"),
        ("purchases", "played_2027", "REAL DEFAULT 0"),
        ("purchases", "played_2028", "REAL DEFAULT 0"),
        ("purchases", "rem_2027", "REAL DEFAULT 0"),
        ("purchases", "rem_2028", "REAL DEFAULT 0"),
        ("nmck_breakdown", "onec_num", "TEXT"),
        ("nmck_breakdown", "ifo_source", "TEXT"),
    ]

    for tbl, col, col_type in cols_to_add:
        try:
            c.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {col_type}")
        except sqlite3.OperationalError:
            pass

    conn.commit()
    conn.close()


init_db()

IFO_SOURCES = ["ВБ", "ГЗ", "ОМС", "Прочее"]

st.title("📋 Реестр закупок")

# --- ФОРМА ВВОДА (Защищена от полупустого создания по Enter!) ---
with st.expander("➕ Добавить новую позицию", expanded=True):
    with st.form("new_entry", clear_on_submit=True):
        row1 = st.columns([2, 3, 1, 2])
        sub = row1[0].selectbox("Подразделение", ["Автохозяйство", "Админ. отдел", "Лаборатория", "АХО"])
        name = row1[1].text_input("Наименование")
        y_place = row1[2].selectbox("Год размещения", list(range(2027, 2032)))
        ifo_main = row1[3].multiselect("ИФО (источники)", IFO_SOURCES, default=[])

        row2 = st.columns([2, 1, 1])
        okpd_raw = row2[0].text_input("ОКПД2 (вводите цифры)", placeholder="Напр: 123456")
        kosgu = row2[1].selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = row2[2].selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])

        row3 = st.columns(2)
        req_n = row3[0].text_input("Номер предложения на закупку")
        graph_n = row3[1].text_input("Номер план-графика")

        submit_btn = st.form_submit_button("Сохранить позицию")

        if submit_btn:
            # Защитная проверка: Не даем создать пустую строку от случайного нажатия Enter
            if not name or not name.strip():
                st.warning("⚠️ Пожалуйста, введите Наименование закупки перед сохранением!")
            else:
                okpd_fixed = format_okpd(okpd_raw)
                name_fixed = capitalize_first_letter(name)
                conn = get_connection()
                conn.execute('''INSERT INTO purchases
                                (subdivision, name, year_placement, ifo, okpd2, kosgu, basis, request_num, plan_graph_num,
                                 plan_2027, plan_2028, nmck_2027, nmck_2028, played_2027, played_2028, rem_2027, rem_2028)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0, 0, 0, 0, 0, 0)''',
                             (sub, name_fixed, y_place, ", ".join(ifo_main), okpd_fixed, kosgu, basis, req_n, graph_n))
                conn.commit()
                conn.close()
                st.success(f"Добавлено! Сохранено: {name_fixed}")
                st.rerun()

# --- ОСНОВНАЯ ТАБЛИЦА ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    st.subheader("Главный реестр")

    # Авто-расчет остатков
    df['rem_2027'] = df['plan_2027'] - df['nmck_2027'] - df['played_2027']
    df['rem_2028'] = df['plan_2028'] - df['nmck_2028'] - df['played_2028']

    display_df = df[[
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
        selected_sources = [s.strip() for s in raw_ifo_str.split(",") if s.strip()]
        if not selected_sources:
            selected_sources = IFO_SOURCES

        st.success(f"📌 Выбрана позиция: **{selected_name}** (ID: {sel_id}, ИФО: {', '.join(selected_sources)})")

        # --- БЛОК 1: РАСПРЕДЕЛЕНИЕ БЮДЖЕТА ПО ИСТОЧНИКАМ (ИФО) ---
        with st.expander("💰 Распределение бюджета по источникам (ИФО)", expanded=True):
            years = [2027, 2028]
            cols = st.columns(2)

            for i, year in enumerate(years):
                with cols[i]:
                    st.markdown(f"**Бюджет на {year} год**")
                    with st.container(border=True):
                        total_for_year = 0.0
                        for source in selected_sources:
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
                        contracts = conn.execute(
                            "SELECT id, contract_name, amount, onec_num, ifo_source FROM nmck_breakdown WHERE purchase_id=? AND year=?",
                            (sel_id, year)).fetchall()
                        conn.close()

                        total_nmck_year = sum(c[2] for c in contracts)
                        next_auto_num = f"№{len(contracts) + 1}"

                        if contracts:
                            for idx, (cid, cname, camount, onec_val, ifo_val) in enumerate(contracts, start=1):
                                c_col1, c_col2, c_col3, c_col4, c_col5, c_col6 = st.columns([1, 2, 2, 2, 1, 1])
                                c_col1.write(f"**№{idx}**")
                                c_col2.write(f"1С: {onec_val if onec_val else '—'}")
                                c_col3.write(f"ИФО: **{ifo_val if ifo_val else '—'}**")
                                c_col4.write(f"{camount:,.2f} руб.")

                                if c_col5.button("✏️", key=f"edit_nmck_btn_{cid}"):
                                    st.session_state[f"editing_nmck_{cid}"] = not st.session_state.get(f"editing_nmck_{cid}", False)

                                if c_col6.button("❌", key=f"del_nmck_{cid}"):
                                    conn = get_connection()
                                    conn.execute("DELETE FROM nmck_breakdown WHERE id=?", (cid,))
                                    conn.commit()
                                    conn.close()
                                    st.rerun()

                                if st.session_state.get(f"editing_nmck_{cid}", False):
                                    with st.form(key=f"form_edit_nmck_{cid}"):
                                        e_onec = st.text_input("Изменить Номер 1С", value=onec_val if onec_val else "")
                                        e_ifo = st.selectbox("Изменить ИФО", options=selected_sources, index=selected_sources.index(ifo_val) if ifo_val in selected_sources else 0)
                                        e_amt = st.number_input("Изменить Сумму", value=float(camount), format="%.2f")
                                        if st.form_submit_button("Сохранить изменения"):
                                            conn = get_connection()
                                            conn.execute("UPDATE nmck_breakdown SET onec_num=?, ifo_source=?, amount=? WHERE id=?",
                                                         (e_onec, e_ifo, e_amt, cid))
                                            conn.commit()
                                            conn.close()
                                            st.session_state[f"editing_nmck_{cid}"] = False
                                            st.rerun()

                        st.caption(f"Добавить заявку ({next_auto_num}):")
                        add_col1, add_col2, add_col3 = st.columns([2, 2, 2])
                        new_onec = add_col1.text_input("Номер 1С", key=f"new_onec_{sel_id}_{year}")
                        new_ifo_source = add_col2.selectbox("Источник ИФО", options=selected_sources, key=f"new_ifo_src_{sel_id}_{year}")
                        new_camount = add_col3.number_input("Сумма, руб.", value=0.0, format="%.2f", key=f"new_camount_{sel_id}_{year}")

                        if st.button(f"➕ Добавить заявку {next_auto_num} ({year})", key=f"btn_add_nmck_{sel_id}_{year}"):
                            if new_camount > 0:
                                conn = get_connection()
                                conn.execute("INSERT INTO nmck_breakdown (purchase_id, year, contract_name, amount, onec_num, ifo_source) VALUES (?,?,?,?,?,?)",
                                             (sel_id, year, next_auto_num, new_camount, new_onec, new_ifo_source))
                                conn.commit()
                                conn.close()
                                st.rerun()

                        st.markdown(f"---")
                        st.write(f"**ИТОГО НМЦК {year}: {total_nmck_year:,.2f} руб.**")

                        conn = get_connection()
                        conn.execute(f"UPDATE purchases SET nmck_{year} = ? WHERE id = ?", (total_nmck_year, sel_id))
                        conn.commit()
                        conn.close()

        # --- БЛОК 3: СУММА СЫГРАННАЯ (КОНТРАКТЫ + ДОПОЛНИТЕЛЬНЫЕ СОГЛАШЕНИЯ ДС) ---
        with st.expander("🤝 Контракты (Сумма сыгранная и ДС)", expanded=False):
            cnt_years = [2027, 2028]
            cnt_cols = st.columns(2)

            for i, year in enumerate(cnt_years):
                with cnt_cols[i]:
                    st.markdown(f"**Контракты на {year} год**")
                    with st.container(border=True):
                        conn = get_connection()
                        req_1c_tuples = conn.execute(
                            "SELECT DISTINCT onec_num FROM nmck_breakdown WHERE purchase_id=? AND year=? AND onec_num IS NOT NULL AND onec_num != ''",
                            (sel_id, year)).fetchall()
                        avail_1c = [t[0] for t in req_1c_tuples]

                        cnt_list = conn.execute(
                            "SELECT id, contract_num, onec_num, contract_date, amount, comment FROM contracts_breakdown WHERE purchase_id=? AND year=?",
                            (sel_id, year)).fetchall()
                        conn.close()

                        next_cnt_num = f"№{len(cnt_list) + 1}"
                        total_played_year = 0.0

                        if cnt_list:
                            for idx, (cid, cnum, c1c, cdate, camt, ccomm) in enumerate(cnt_list, start=1):
                                # Считаем все ДС для этого конкретного контракта
                                conn = get_connection()
                                ds_items = conn.execute("SELECT id, ds_num, ds_date, amount, comment FROM contract_ds WHERE contract_id=?", (cid,)).fetchall()
                                conn.close()
                                ds_sum = sum(d[3] for d in ds_items)
                                full_contract_total = camt + ds_sum
                                total_played_year += full_contract_total

                                k1, k2, k3, k4, k5, k6, k7 = st.columns([1, 2, 2, 2, 2, 1, 1])
                                k1.write(f"**№{idx}**")
                                k2.write(f"1С: **{c1c}**")
                                k3.write(f"Дата: {cdate}")
                                k4.write(f"Основа: {camt:,.2f} руб. (Всего с ДС: **{full_contract_total:,.2f}**)")
                                k5.caption(f"💬 {ccomm}" if ccomm else "")

                                if k6.button("✏️", key=f"edit_cnt_btn_{cid}"):
                                    st.session_state[f"editing_cnt_{cid}"] = not st.session_state.get(f"editing_cnt_{cid}", False)

                                if k7.button("❌", key=f"del_cnt_{cid}"):
                                    conn = get_connection()
                                    conn.execute("DELETE FROM contracts_breakdown WHERE id=?", (cid,))
                                    conn.execute("DELETE FROM contract_ds WHERE contract_id=?", (cid,))
                                    conn.commit()
                                    conn.close()
                                    st.rerun()

                                if st.session_state.get(f"editing_cnt_{cid}", False):
                                    with st.form(key=f"form_edit_cnt_{cid}"):
                                        ec_1c = st.selectbox("Изменить 1С", options=avail_1c, index=avail_1c.index(c1c) if c1c in avail_1c else 0)
                                        ec_amt = st.number_input("Изменить Сумму", value=float(camt), format="%.2f")
                                        ec_comm = st.text_input("Изменить Комментарий", value=ccomm if ccomm else "")
                                        if st.form_submit_button("Сохранить контракт"):
                                            conn = get_connection()
                                            conn.execute("UPDATE contracts_breakdown SET onec_num=?, amount=?, comment=? WHERE id=?",
                                                         (ec_1c, ec_amt, ec_comm, cid))
                                            conn.commit()
                                            conn.close()
                                            st.session_state[f"editing_cnt_{cid}"] = False
                                            st.rerun()

                                # --- ДОПОЛНИТЕЛЬНЫЕ СОГЛАШЕНИЯ (ДС) К КОНТРАКТУ ---
                                with st.expander(f"📑 Доп. соглашения (ДС) к контракту №{idx} (Сумма ДС: {ds_sum:,.2f} руб.)", expanded=False):
                                    if ds_items:
                                        for ds_id, ds_num, ds_date, ds_amt, ds_comm in ds_items:
                                            d1, d2, d3, d4, d5 = st.columns([2, 2, 2, 3, 1])
                                            d1.write(f"ДС: **{ds_num}**")
                                            d2.write(f"Дата: {ds_date}")
                                            d3.write(f"{ds_amt:,.2f} руб.")
                                            d4.caption(f"💬 {ds_comm}" if ds_comm else "")
                                            if d5.button("❌", key=f"del_ds_{ds_id}"):
                                                conn = get_connection()
                                                conn.execute("DELETE FROM contract_ds WHERE id=?", (ds_id,))
                                                conn.commit()
                                                conn.close()
                                                st.rerun()

                                    st.caption("Добавить Дополнительное Соглашение (ДС):")
                                    ds_c1, ds_c2, ds_c3, ds_c4 = st.columns([2, 2, 2, 3])
                                    new_ds_num = ds_c1.text_input("№ ДС", key=f"ds_num_{cid}")
                                    new_ds_date = ds_c2.date_input("Дата ДС", format="DD-MM-YYYY", key=f"ds_date_{cid}")
                                    new_ds_amt = ds_c3.number_input("Сумма ДС, руб.", value=0.0, format="%.2f", key=f"ds_amt_{cid}")
                                    new_ds_comm = ds_c4.text_input("Коммент к ДС", key=f"ds_comm_{cid}")

                                    if st.button(f"➕ Добавить ДС к контракту №{idx}", key=f"btn_add_ds_{cid}"):
                                        if new_ds_amt > 0:
                                            formatted_ds_date = new_ds_date.strftime("%d-%m-%Y")
                                            conn = get_connection()
                                            conn.execute("INSERT INTO contract_ds (contract_id, ds_num, ds_date, amount, comment) VALUES (?,?,?,?,?)",
                                                         (cid, new_ds_num if new_ds_num else "ДС", formatted_ds_date, new_ds_amt, new_ds_comm))
                                            conn.commit()
                                            conn.close()
                                            st.rerun()

                        st.caption(f"Добавить новый контракт ({next_cnt_num}):")

                        if not avail_1c:
                            st.info("💡 *Сначала добавьте хотя бы одну заявку с Номером 1С выше.*")
                        else:
                            ca1, ca2, ca3, ca4 = st.columns([2, 2, 2, 2])
                            c_1c_sel = ca1.selectbox("Номер 1С (из заявок)", options=avail_1c, key=f"c_1c_{sel_id}_{year}")
                            c_date_val = ca2.date_input("Дата контракта", format="DD-MM-YYYY", key=f"c_date_{sel_id}_{year}")
                            c_amt_val = ca3.number_input("Сумма, руб.", value=0.0, format="%.2f", key=f"c_amt_{sel_id}_{year}")
                            c_comm_val = ca4.text_input("Комментарий", key=f"c_comm_{sel_id}_{year}")

                            if st.button(f"➕ Добавить контракт {next_cnt_num} ({year})", key=f"btn_add_cnt_{sel_id}_{year}"):
                                if c_amt_val > 0:
                                    formatted_date_str = c_date_val.strftime("%d-%m-%Y")
                                    conn = get_connection()
                                    conn.execute(
                                        "INSERT INTO contracts_breakdown (purchase_id, year, contract_num, onec_num, contract_date, amount, comment) VALUES (?,?,?,?,?,?,?)",
                                        (sel_id, year, next_cnt_num, c_1c_sel, formatted_date_str, c_amt_val, c_comm_val))
                                    conn.commit()
                                    conn.close()
                                    st.rerun()

                        st.markdown("---")
                        st.write(f"**ИТОГО СЫГРАНО {year}: {total_played_year:,.2f} руб.**")

                        conn = get_connection()
                        conn.execute(f"UPDATE purchases SET played_{year} = ? WHERE id = ?", (total_played_year, sel_id))
                        conn.commit()
                        conn.close()

        if st.button("🔄 Обновить итоговые суммы в таблице"):
            st.rerun()