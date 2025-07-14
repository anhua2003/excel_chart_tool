import streamlit as st
import pandas as pd
import plotly.express as px
import io
import xlsxwriter

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
        "download_excel": "⬇️ Tải Excel (biểu đồ thật)",
        "download_html": "⬇️ Tải biểu đồ HTML",
        "download_png": "⬇️ Tải ảnh biểu đồ PNG",
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
        "download_excel": "⬇️ Download Excel (real chart)",
        "download_html": "⬇️ Download HTML chart",
        "download_png": "⬇️ Download PNG chart",
    }
}

st.set_page_config(page_title="Excel Chart Tool", page_icon="📊")

lang = st.selectbox("🌐 Language / Ngôn ngữ", ["vi", "en"], key="lang")
T = LANG[lang]

st.title(T["title"])

if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None

uploaded = st.file_uploader(T["upload"], type=["xlsx"])
if uploaded:
    st.session_state.uploaded_file = uploaded

file = st.session_state.uploaded_file

if file:
    df = pd.read_excel(file)
    st.write(T["data"])
    st.dataframe(df)

    total_rows = df.shape[0]
    start_row = st.number_input(T["start_row"], min_value=1, max_value=total_rows, value=1)
    end_row = st.number_input(T["end_row"], min_value=start_row, max_value=total_rows, value=min(start_row+2, total_rows))

    df_subset = df.iloc[start_row - 1:end_row]
    st.write(T["subset"])
    st.dataframe(df_subset)

    numeric_cols = df_subset.select_dtypes(include=['number']).columns.tolist()
    all_cols = df_subset.columns.tolist()

    x_axis = st.selectbox(T["x_axis"], all_cols)
    y_axis = st.multiselect(T["y_axis"], numeric_cols)
    chart_type = st.selectbox(T["chart_type"], [T["chart_bar"], T["chart_line"], T["chart_pie"]])

    if st.button(T["draw_btn"]) and x_axis and y_axis:
        st.session_state.figures = []
        for y in y_axis:
            if chart_type == T["chart_bar"]:
                fig = px.bar(df_subset, x=x_axis, y=y, title=f"{y} vs {x_axis}")
            elif chart_type == T["chart_line"]:
                fig = px.line(df_subset, x=x_axis, y=y, title=f"{y} vs {x_axis}")
            elif chart_type == T["chart_pie"]:
                fig = px.pie(df_subset, names=x_axis, values=y, title=f"{y} by {x_axis}")
            st.session_state.figures.append((fig, y, df_subset, x_axis, chart_type))

if "figures" in st.session_state:
    for fig, y, df_subset, x_axis, chart_type in st.session_state.figures:
        st.plotly_chart(fig)

        # HTML download
        # fig.update_layout(template="plotly_white")  # Ép dùng nền trắng
        fig.update_traces(marker_color="#1f77b4")  # Xanh dương chuẩn
        fig.update_layout(
            template="plotly_white",
            paper_bgcolor='white',
            plot_bgcolor='white',
            font_color='black'
        )
        html_bytes = fig.to_html(full_html=True, include_plotlyjs='include').encode("utf-8")
        st.download_button(
            label=T["download_html"] + f" ({y})",
            data=html_bytes,
            file_name=f"chart_{y}.html",
            mime="text/html"
        )

        # Excel download (dùng BytesIO)
        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet("Data")

        for col_idx, col in enumerate(df_subset.columns):
            ws.write(0, col_idx, col)
            for row_idx, val in enumerate(df_subset[col]):
                ws.write(row_idx + 1, col_idx, val)

        col_idx = df_subset.columns.get_loc(y)
        x_idx = df_subset.columns.get_loc(x_axis)

        chart_type_excel = 'column' if chart_type != T["chart_pie"] else 'pie'
        chart = wb.add_chart({'type': chart_type_excel})

        chart.add_series({
            'name': y,
            'values': f'=Data!${chr(65+col_idx)}$2:${chr(65+col_idx)}${len(df_subset)+1}',
            'categories': f'=Data!${chr(65+x_idx)}$2:${chr(65+x_idx)}${len(df_subset)+1}'
        })
        ws.insert_chart("H2", chart)
        wb.close()
        output.seek(0)

        st.download_button(
            label=T["download_excel"] + f" ({y})",
            data=output,
            file_name=f"chart_{y}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
