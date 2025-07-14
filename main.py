import streamlit as st
import pandas as pd
import plotly.express as px

# Dictionary ngôn ngữ
LANG = {
    "vi": {
        "title": "🧠 Tạo biểu đồ từ Excel",
        "upload": "📥 Tải file Excel lên",
        "data": "📋 Dữ liệu từ file:",
        "subset": "📌 Dữ liệu đã chọn:",
        "x_axis": "🧭 Chọn trục X",
        "y_axis": "📊 Chọn trục Y",
        "chart_type": "📈 Loại biểu đồ",
        "draw_btn": "🎨 Vẽ biểu đồ",
        "start_row": "🔽 Dòng bắt đầu (từ 1)",
        "end_row": "🔼 Dòng kết thúc (bao gồm)",
        "lang_select": "🌐 Chọn ngôn ngữ",
        "chart_bar": "Biểu đồ cột",
        "chart_line": "Biểu đồ đường",
        "chart_pie": "Biểu đồ tròn",
    },
    "en": {
        "title": "🧠 Create Chart from Excel",
        "upload": "📥 Upload Excel file",
        "data": "📋 Data from file:",
        "subset": "📌 Selected data:",
        "x_axis": "🧭 Select X axis",
        "y_axis": "📊 Select Y axis",
        "chart_type": "📈 Chart type",
        "draw_btn": "🎨 Draw chart",
        "start_row": "🔽 Start row (from 1)",
        "end_row": "🔼 End row (inclusive)",
        "lang_select": "🌐 Select Language",
        "chart_bar": "Bar Chart",
        "chart_line": "Line Chart",
        "chart_pie": "Pie Chart",
    }
}

st.set_page_config(
    page_title="Excel to Chart",
    page_icon="📊",
    layout="wide",
)

# Chọn ngôn ngữ
lang = st.selectbox("🌐 Language / Ngôn ngữ", ["vi", "en"])
T = LANG[lang]

st.title(T["title"])

# Load file nếu chưa có trong session
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None

# Upload file mới (nếu có)
uploaded = st.file_uploader(T["upload"], type=["xlsx"])
if uploaded:
    st.session_state.uploaded_file = uploaded

# Dùng file trong session
file = st.session_state.uploaded_file
if file:
    df = pd.read_excel(file)
    st.write(T["data"])
    st.dataframe(df)

    total_rows = df.shape[0]
    start_row = st.number_input(T["start_row"], min_value=0, max_value=total_rows, value=0)
    end_row = st.number_input(T["end_row"], min_value=start_row, max_value=total_rows, value=min(start_row+3, total_rows))

    df_subset = df.iloc[start_row:end_row]
    st.write(T["subset"])
    st.dataframe(df_subset)

    numeric_cols = df_subset.select_dtypes(include=['number']).columns.tolist()
    all_cols = df_subset.columns.tolist()

    x_axis = st.selectbox(T["x_axis"], all_cols)
    y_axis = st.multiselect(T["y_axis"], numeric_cols)
    chart_type = st.selectbox(T["chart_type"], [T["chart_bar"], T["chart_line"], T["chart_pie"]])

    if st.button(T["draw_btn"]) and x_axis and y_axis:
        for y in y_axis:
            if chart_type == T["chart_bar"]:
                fig = px.bar(df_subset, x=x_axis, y=y, title=f"{y} vs {x_axis}")
            elif chart_type == T["chart_line"]:
                fig = px.line(df_subset, x=x_axis, y=y, title=f"{y} vs {x_axis}")
            elif chart_type == T["chart_pie"]:
                fig = px.pie(df_subset, names=x_axis, values=y, title=f"{y} by {x_axis}")
            st.plotly_chart(fig)
