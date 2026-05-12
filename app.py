import streamlit as st
from PIL import Image
import os
import json
import uuid

st.set_page_config(page_title="Thêm đồ vật")

# Chụp ảnh
photo = st.camera_input("Hình ảnh")

# Nhập thông tin
name = st.text_input("Tên")
length = st.number_input("Chiều dài", min_value=0.0)
width = st.number_input("Chiều rộng", min_value=0.0)

# Nút lưu tạm
if st.button("Xác nhận"):
    if photo is not None:

        # tạo folder images
        os.makedirs("images", exist_ok=True)

        # mở ảnh
        image = Image.open(photo)

        # tạo id
        product_id = str(uuid.uuid4())

        # tên file ảnh
        filename = f"{product_id}.jpg"

        # đường dẫn ảnh
        filepath = f"images/{filename}"

        # lưu ảnh
        image.save(filepath)

        # dữ liệu sản phẩm
        new_data = {
            "id": product_id,
            "ten": name,
            "chieu_dai": length,
            "chieu_rong": width,
            "images": filepath
        }

        # đọc file json cũ
        data = []

        if os.path.exists("data.json"):
            with open("data.json", "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                except:
                    data = []

        # thêm dữ liệu mới
        data.append(new_data)

        # lưu lại json
        with open("data.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        st.success("Đã lưu dữ liệu!")

        st.image(filepath)

        st.json(new_data)

        st.rerun()