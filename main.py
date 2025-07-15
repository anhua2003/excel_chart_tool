import streamlit as st
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
st.set_page_config(layout="wide")
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
        "chart_bar_horizontal": "Biểu đồ cột ngang",
        "chart_line": "Biểu đồ đường",
        "chart_pie": "Biểu đồ tròn",
        "download_excel": "⬇️ Tải Excel (biểu đồ thật)",
        "download_html": "⬇️ Tải biểu đồ HTML",
        "download_png": "⬇️ Tải ảnh biểu đồ PNG",
        "combined_one": "Gộp vào 1 biểu đồ",
        "combined_n" : "Tách từng biểu đồ",
        "combined_type": "📌 Cách vẽ",
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
        "chart_bar_horizontal": "Horizontal Bar",
        "chart_line": "Line Chart",
        "chart_pie": "Pie Chart",
        "download_excel": "⬇️ Download Excel (real chart)",
        "download_html": "⬇️ Download HTML chart",
        "download_png": "⬇️ Download PNG chart",
        "combined_one": "Combined",
        "combined_n" : "Non-combined",
        "combined_type": "📌 Colouration",
    }
}

st.set_page_config(page_title="Excel Chart Tool", page_icon="📊")

lang = st.selectbox("🌐 Language / Ngôn ngữ", ["vi", "en"], key="lang")

if "lang_old" not in st.session_state:
    st.session_state.lang_old = lang
elif st.session_state.lang_old != lang:
    st.session_state.lang_old = lang

    # ✅ Giữ lại uploaded_file, lang, lang_old
    keys_to_keep = {"uploaded_file", "lang", "lang_old"}
    keys_to_clear = [k for k in st.session_state.keys() if k not in keys_to_keep]
    for k in keys_to_clear:
        del st.session_state[k]

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
    df.index += 1
    st.write(T["data"])
    st.dataframe(df)

    total_rows = df.shape[0]
    start_row = st.number_input(T["start_row"], min_value=1, max_value=total_rows, value=1)
    end_row = st.number_input(T["end_row"], min_value=start_row, max_value=total_rows, value=min(start_row + 2, total_rows))

    df_subset = df.iloc[start_row - 1:end_row]
    st.write(T["subset"])
    st.dataframe(df_subset)

    numeric_cols = df_subset.select_dtypes(include=['number']).columns.tolist()
    all_cols = df_subset.columns.tolist()

    x_axis = st.selectbox(T["x_axis"], all_cols)
    y_axis = st.multiselect(T["y_axis"], numeric_cols)
    chart_type = st.selectbox(T["chart_type"], [T["chart_bar"], T["chart_bar_horizontal"], T["chart_line"], T["chart_pie"]])
    combine_chart = st.radio(T["combined_type"], [T["combined_one"], T["combined_n"]], key="combine_mode") if chart_type != T["chart_pie"] else None

    # 🧠 Gộp dữ liệu trùng `x_axis`
    grouped = df_subset.groupby(x_axis, as_index=False)[y_axis].sum(numeric_only=True)
    st.session_state.grouped = grouped
   
    if "figures" not in st.session_state:
        st.session_state.figures = {}

    if st.button(T["draw_btn"]) and x_axis and y_axis:
        st.session_state.figures.clear()

        if chart_type == T["chart_pie"]:
            for y in y_axis:
                if y in grouped.columns:
                    fig = px.pie(grouped, names=x_axis, values=y, title=f"{y} by {x_axis}", color_discrete_sequence=px.colors.qualitative.Plotly)
                    st.session_state.figures[y] = fig

        elif combine_chart == T["combined_one"]:
            if chart_type == T["chart_bar"]:
                fig = px.bar(grouped, x=x_axis, y=y_axis, title=" vs ".join(y_axis), color_discrete_sequence=px.colors.qualitative.Plotly)
                fig.update_layout(barmode='group')  # 🟢 Quan trọng
            elif chart_type == T['chart_bar_horizontal']:
                fig = px.bar(grouped, x=y_axis, y=x_axis, orientation='h', title=" vs ".join(y_axis), color_discrete_sequence=px.colors.qualitative.Plotly)
                fig.update_layout(barmode='group')
            else:
                fig = px.line(grouped, x=x_axis, y=y_axis, title=" vs ".join(y_axis), color_discrete_sequence=px.colors.qualitative.Plotly)
            st.session_state.figures["combined"] = fig

        elif combine_chart == T["combined_n"]:
            for y in y_axis:
                if chart_type == T["chart_bar"]:
                    fig = px.bar(grouped, x=x_axis, y=y, title=f"{y} vs {x_axis}", color_discrete_sequence=px.colors.qualitative.Plotly)
                elif chart_type == T["chart_bar_horizontal"]:
                    fig = px.bar(grouped, x=y, y=x_axis, orientation='h', title=f"{y} vs {x_axis}", color_discrete_sequence=px.colors.qualitative.Plotly)
                else:
                    fig = px.line(grouped, x=x_axis, y=y, title=f"{y} vs {x_axis}", color_discrete_sequence=px.colors.qualitative.Plotly)
                st.session_state.figures[y] = fig

    # Hiển thị và tải các biểu đồ đã lưu
    for y, fig in st.session_state.figures.items():
        st.subheader(f"📊 {y}")
        st.plotly_chart(fig)

        # HTML download
        html_bytes = fig.to_html(full_html=True, include_plotlyjs='include').encode("utf-8")
        st.download_button(
            label=T["download_html"] + f" ({y})",
            data=html_bytes,
            file_name=f"chart_{y}.html",
            mime="text/html"
        )

        # ✅ Chỉ xuất Excel nếu `y` là cột hợp lệ
        grouped = st.session_state.get("grouped", df_subset)
        if y in grouped.columns:
            output = io.BytesIO()
            wb = xlsxwriter.Workbook(output, {'in_memory': True})
            ws = wb.add_worksheet("Data")

            for col_idx, col in enumerate(grouped.columns):
                ws.write(0, col_idx, col)
                for row_idx, val in enumerate(grouped[col]):
                    ws.write(row_idx + 1, col_idx, val)

            col_idx = grouped.columns.get_loc(y)
            x_idx = grouped.columns.get_loc(x_axis)

            if chart_type == T["chart_pie"]:
                chart_type_excel = 'pie'
            elif chart_type == T["chart_line"]:
                chart_type_excel = 'line'
            elif chart_type == T["chart_bar_horizontal"]:
                chart_type_excel = 'bar'
            else:
                chart_type_excel = 'column'

            chart = wb.add_chart({'type': chart_type_excel})
            chart.add_series({
                'name': y,
                'values': f'=Data!${chr(65 + col_idx)}$2:${chr(65 + col_idx)}${len(grouped) + 1}',
                'categories': f'=Data!${chr(65 + x_idx)}$2:${chr(65 + x_idx)}${len(grouped) + 1}'
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
        elif y == "combined":
            output = io.BytesIO()
            wb = xlsxwriter.Workbook(output, {'in_memory': True})
            ws = wb.add_worksheet("Data")

            for col_idx, col in enumerate(grouped.columns):
                ws.write(0, col_idx, col)
                for row_idx, val in enumerate(grouped[col]):
                    ws.write(row_idx + 1, col_idx, val)

            x_idx = grouped.columns.get_loc(x_axis)
            if chart_type == T["chart_pie"]:
                chart_type_excel = 'pie'
            elif chart_type == T["chart_line"]:
                chart_type_excel = 'line'
            elif chart_type == T["chart_bar_horizontal"]:
                chart_type_excel = 'bar'
            else:
                chart_type_excel = 'column'
            chart = wb.add_chart({'type': chart_type_excel})

            for y_col in y_axis:
                y_idx = grouped.columns.get_loc(y_col)
                chart.add_series({
                    'name': y_col,
                    'values': f'=Data!${chr(65 + y_idx)}$2:${chr(65 + y_idx)}${len(grouped) + 1}',
                    'categories': f'=Data!${chr(65 + x_idx)}$2:${chr(65 + x_idx)}${len(grouped) + 1}'
                })

            ws.insert_chart("H2", chart)
            wb.close()
            output.seek(0)

            st.download_button(
                label=T["download_excel"] + " (combined)",
                data=output,
                file_name=f"chart_combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
