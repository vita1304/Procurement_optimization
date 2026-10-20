import io
import math
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Supply Chain & Inventory Calculator", layout="wide"
)

# Словник перекладів для двомовності
translations = {
    "UA": {
        "title": "📦 Калькулятор управління запасами та логістикою",
        "subtitle": "Автоматизований розрахунок **EOQ**, **Страхового запасу** та **Точки перезамовлення** з урахуванням українських бізнес-реалій.",
        "mode_label": "Оберіть режим роботи:",
        "mode_single": "Калькулятор одного товару",
        "mode_batch": "Пакетний розрахунок (Excel)",
        "lang_label": "Мова / Language:",
        # Режим 1
        "s1_title": "⚙️ Розрахунок параметрів для окремого SKU",
        "ann_dem": "Річна потреба (Demand, од.)",
        "ord_cost": "Вартість оформлення 1 замовлення (грн)",
        "hold_cost": "Вартість утримання 1 од. на складі за рік (грн)",
        "lead_time": "Середній час доставки (дні)",
        "max_lead_time": "Максимальна затримка доставки (дні)",
        "max_daily": "Максимальне споживання на день (од.)",
        "calc_btn": "Розрахувати показники",
        "success_msg": "✅ Розрахунок успішно виконано!",
        "res_eoq": "EOQ (Партія замовлення)",
        "res_ss": "Страховий запас",
        "res_rop": "Точка перезамовлення (ROP)",
        "res_orders": "Кількість замовлень на рік",
        # Режим 2
        "s2_title": "📂 Завантаження файлу зі списком товарів",
        "excel_instructions": """
        **Вимоги до структури Excel-файлу:**
        Для успішного розрахунку ваш файл повинен містити наступні колонки з точними назвами (або програма автоматично додасть їх результати):
        * `Annual_Demand` — річна потреба товару
        * `Ordering_Cost` — вартість оформлення замовлення
        * `Holding_Cost` — вартість зберігання одиниці
        * `Lead_Time` — стандартний час доставки (дні)
        * `Max_Lead_Time` — максимальний час затримки доставки (дні)
        * `Max_Daily_Demand` — максимальне щоденне споживання
        
        *Програма зчитає ці дані, розрахує нові показники та збереже їх у нових стовпчиках готового файлу.*
        """,
        "uploader_label": "Оберіть Excel-файл (.xlsx)",
        "preview_title": "### Попередній перегляд завантажених даних:",
        "batch_btn": "Запустити пакетний розрахунок",
        "batch_success": "🎉 Успішно розраховано для всіх товарів!",
        "download_btn": "📥 Завантажити результати в Excel",
        "error_msg": "Постала помилка при обробці файлу: ",
    },
    "EN": {
        "title": "📦 Supply Chain & Inventory Calculator",
        "subtitle": "Automated calculation of **EOQ**, **Safety Stock**, and **Reorder Point** adjusted for market realities.",
        "mode_label": "Select mode:",
        "mode_single": "Single Item Calculator",
        "mode_batch": "Batch Calculation (Excel)",
        "lang_label": "Language / Мова:",
        # Mode 1
        "s1_title": "⚙️ Parameters Calculation for a Single SKU",
        "ann_dem": "Annual Demand (units)",
        "ord_cost": "Ordering Cost per Order",
        "hold_cost": "Holding Cost per Unit/Year",
        "lead_time": "Average Lead Time (days)",
        "max_lead_time": "Max Lead Time Delay (days)",
        "max_daily": "Max Daily Demand (units)",
        "calc_btn": "Calculate Metrics",
        "success_msg": "✅ Calculation completed successfully!",
        "res_eoq": "EOQ (Order Quantity)",
        "res_ss": "Safety Stock",
        "res_rop": "Reorder Point (ROP)",
        "res_orders": "Orders per Year",
        # Mode 2
        "s2_title": "📂 Upload File with Product List",
        "excel_instructions": """
        **Excel File Structure Requirements:**
        For a successful calculation, your file must contain the following columns with exact names:
        * `Annual_Demand` — annual product demand
        * `Ordering_Cost` — cost of placing an order
        * `Holding_Cost` — holding cost per unit
        * `Lead_Time` — standard delivery time (days)
        * `Max_Lead_Time` — maximum delivery delay time (days)
        * `Max_Daily_Demand` — maximum daily consumption
        
        *The app will read these columns, compute the metrics, and append them into new columns.*
        """,
        "uploader_label": "Choose an Excel file (.xlsx)",
        "preview_title": "### Preview of Uploaded Data:",
        "batch_btn": "Run Batch Calculation",
        "batch_success": "🎉 Successfully calculated for all items!",
        "download_btn": "📥 Download Results as Excel",
        "error_msg": "An error occurred while processing the file: ",
    },
}

# Вибір мови у боковій панелі
selected_lang = st.sidebar.selectbox("🌐 Language / Мова", ["UA", "EN"])
t = translations[selected_lang]

st.title(t["title"])
st.markdown(t["subtitle"])

# Вибір режиму роботи через бокову панель
mode = st.sidebar.selectbox(
    t["mode_label"],
    [t["mode_single"], t["mode_batch"]],
)

# ==========================================
# РЕЖИМ 1: КАЛЬКУЛЯТОР ОДНОГО ТОВАРУ
# ==========================================
if mode == t["mode_single"]:
    st.subheader(t["s1_title"])

    col1, col2 = st.columns(2)

    with col1:
        annual_demand = st.number_input(
            t["ann_dem"],
            min_value=1.0,
            value=10000.0,
            step=100.0,
        )
        ordering_cost = st.number_input(
            t["ord_cost"],
            min_value=0.1,
            value=500.0,
            step=50.0,
        )
        holding_cost = st.number_input(
            t["hold_cost"],
            min_value=0.1,
            value=25.0,
            step=5.0,
        )

    with col2:
        lead_time_days = st.number_input(
            t["lead_time"], min_value=1.0, value=14.0, step=1.0
        )
        max_lead_time_days = st.number_input(
            t["max_lead_time"],
            min_value=1.0,
            value=25.0,
            step=1.0,
        )
        daily_demand_max = st.number_input(
            t["max_daily"],
            min_value=0.1,
            value=40.0,
            step=5.0,
        )

    if st.button(t["calc_btn"], type="primary"):
        # Математика EOQ
        eoq = (
            math.sqrt((2 * annual_demand * ordering_cost) / holding_cost)
            if holding_cost > 0
            else 0
        )

        # Середнє щоденне споживання
        avg_daily_demand = annual_demand / 365.0

        # Страховий запас
        safety_stock = (daily_demand_max * max_lead_time_days) - (
            avg_daily_demand * lead_time_days
        )
        safety_stock = max(0.0, safety_stock)

        # Точка перезамовлення (ROP)
        rop = (avg_daily_demand * lead_time_days) + safety_stock

        st.success(t["success_msg"])

        m1, m2, m3, m4 = st.columns(4)
        m1.metric(t["res_eoq"], f"{round(eoq, 2)}")
        m2.metric(t["res_ss"], f"{round(safety_stock, 2)}")
        m3.metric(t["res_rop"], f"{round(rop, 2)}")
        m4.metric(
            t["res_orders"],
            f"{round(annual_demand / eoq, 1) if eoq > 0 else 0}",
        )

# ==========================================
# РЕЖИМ 2: ПАКЕТНИЙ РОЗРАХУНОК (EXCEL)
# ==========================================
elif mode == t["mode_batch"]:
    st.subheader(t["s2_title"])
    st.markdown(t["excel_instructions"])

    uploaded_file = st.file_uploader(
        t["uploader_label"], type=["xlsx"]
    )

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.write(t["preview_title"])
            st.dataframe(df.head())

            if st.button(t["batch_btn"], type="primary"):
                results = []
                for _, row in df.iterrows():
                    d = row.get("Annual_Demand", 1000)
                    s = row.get("Ordering_Cost", 500)
                    h = row.get("Holding_Cost", 20)
                    lt = row.get("Lead_Time", 10)
                    plt_max = row.get("Max_Lead_Time", 15)
                    d_max = row.get("Max_Daily_Demand", 30)

                    # Формули
                    eoq = math.sqrt((2 * d * s) / h) if h > 0 else 0
                    avg_d = d / 365.0
                    ss = max(0.0, (d_max * plt_max) - (avg_d * lt))
                    rop = (avg_d * lt) + ss

                    results.append(
                        {
                            **row.to_dict(),
                            "EOQ": round(eoq, 2),
                            "Safety_Stock": round(ss, 2),
                            "Reorder_Point": round(rop, 2),
                        }
                    )

                result_df = pd.DataFrame(results)

                st.success(t["batch_success"])
                st.dataframe(result_df)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    result_df.to_excel(
                        writer, index=False, sheet_name="Supply_Calculations"
                    )
                processed_data = output.getvalue()

                st.download_button(
                    label=t["download_btn"],
                    data=processed_data,
                    file_name="supply_calculation_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error(t["error_msg"] + str(e))