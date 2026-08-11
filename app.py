import streamlit as st
import pandas as pd
import sqlite3
import re

st.set_page_config(layout="wide", page_title="Закупки НМИЦ")

# --- ФУНКЦИЯ ДЛЯ ОКПД2 (авто-точки) ---
def format_okpd(raw_text):
    # Убираем всё, кроме цифр
    clean = re.sub(r'\D', '', raw_text)
    # Разбиваем по 2 цифры и соединяем точками
    formatted = ".".join(clean[i:i+2] for i in range(0, len(clean), 2))
    return formatted

# --- БАЗА ДАННЫХ ---
def get_connection():
    return sqlite3.connect('procurement_v7.db')

def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS purchases
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  subdivision TEXT, name TEXT, year INTEGER,
                  plan_2026 REAL, plan_2027 REAL, plan_2028 REAL,
                  okpd2 TEXT, kosgu TEXT, basis TEXT, ifo TEXT, 
                  request_num TEXT, plan_graph_num TEXT)''')
    conn.commit()
    conn.close()

init_db()

st.title("📋 Учет закупок (Авто-форматирование)")

# --- ФОРМА ВВОДА ---
with st.expander("➕ Добавить новую позицию", expanded=True):
    with st.form("new_entry"):
        col1, col2, col3 = st.columns([2, 3, 1])
        sub = col1.selectbox("Подразделение", ["Автохозяйство", "Админ. отдел", "Лаборатория", "АХО"])
        name = col2.text_input("Наименование объекта")
        year = col3.selectbox("Год разм.", list(range(2026, 2031)))

        st.write("### Финансы (авто-копейки)")
        c1, c2, c3 = st.columns(3)
        # format="%.2f" заставляет Python всегда показывать .00
        p26 = c1.number_input("План 2026", format="%.2f", step=0.01)
        p27 = c2.number_input("План 2027", format="%.2f", step=0.01)
        p28 = c3.number_input("План 2028", format="%.2f", step=0.01)

        st.write("### Кодировки")
        d1, d2, d3, d4 = st.columns(4)
        okpd_raw = d1.text_input("ОКПД2 (введите просто цифры)", placeholder="Напр: 123456")
        kosgu = d2.selectbox("КОСГУ", ["225", "226", "310", "340"])
        basis = d3.selectbox("Основание", ["44-ФЗ", "223-ФЗ", "ВБ", "ГЗ", "ОМС"])
        ifo = d4.text_input("ИФО")

        if st.form_submit_button("Сохранить"):
            # Применяем магию точек к ОКПД2 перед сохранением
            okpd_fixed = format_okpd(okpd_raw)

            conn = get_connection()
            conn.cursor().execute('''INSERT INTO purchases 
                (subdivision, name, year, plan_2026, plan_2027, plan_2028, okpd2, kosgu, basis, ifo)
                VALUES (?,?,?,?,?,?,?,?,?,?)''',
                (sub, name, year, p26, p27, p28, okpd_fixed, kosgu, basis, ifo))
            conn.commit()
            conn.close()
            st.success(f"Готово! ОКПД2 сохранен как: {okpd_fixed}")
            st.rerun()

# --- ТАБЛИЦА ---
conn = get_connection()
df = pd.read_sql_query("SELECT * FROM purchases", conn)
conn.close()

if not df.empty:
    # Настройка колонок в таблице, чтобы копейки были видны ВСЕГДА
    st.dataframe(df.drop(columns=['id']), use_container_width=True, hide_index=True,
                 column_config={
                     "plan_2026": st.column_config.NumberColumn("План 2026", format="%.2f руб."),
                     "plan_2027": st.column_config.NumberColumn("План 2027", format="%.2f руб."),
                     "plan_2028": st.column_config.NumberColumn("План 2028", format="%.2f руб."),
                 })