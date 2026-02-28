import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import os
import re
import pandas as pd

# Cấu hình đường dẫn Tesseract OCR trên Windows (nếu cài đặt ở thư mục mặc định)
tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(tesseract_path):
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

# Cấu hình giao diện trang
st.set_page_config(page_title="Trích xuất & Nhập liệu Đơn hàng", layout="wide")

st.title("� Trích xuất & Nhập liệu Đơn hàng vào File")
st.markdown("Hỗ trợ đọc các thông tin: **Tên người nhận**, **Địa chỉ**, và **Order ID** từ nhãn giao hàng (PDF/Hình ảnh) rồi xuất ra bảng dữ liệu.")

def extract_label_info(text):
    """Hàm trích xuất thông tin người nhận, địa chỉ, order ID dựa vào cấu trúc và vị trí văn bản"""
    info = {
        "Order ID": "",
        "Tên người nhận": "",
        "Địa chỉ": "",
        "Văn bản gốc": text.strip()
    }
    
    # 1. Trích xuất Order ID (Vị trí góc dưới)
    order_match = re.search(r'(?i)Order\s*ID\s*:\s*([A-Za-z0-9]+)', text)
    if order_match:
        info["Order ID"] = order_match.group(1).strip()
        
    # --- XỬ LÝ TEXT TRƯỚC KHI TRÍCH XUẤT TÊN & ĐỊA CHỈ ---
    # Cắt bỏ phần "Người gửi" để tránh nhầm lẫn
    text_without_sender = re.sub(r'(?i)Người\s*gửi.*?(?=Người\s*nhận)', '', text, flags=re.DOTALL)
    
    # 2. Trích xuất Tên người nhận
    # Lấy khoảng text nằm ngay sau "Người nhận" và trước SĐT
    receiver_block_match = re.search(r'(?i)Người\s*nhận(.*?)\n(.*)', text_without_sender, flags=re.DOTALL)
    
    if receiver_block_match:
        line_with_receiver = receiver_block_match.group(1).strip()
        
        # Nếu trên cùng dòng với chữ "Người nhận" có chứa Tên
        if len(line_with_receiver) > 1:
            # Lọc bỏ số điện thoại, mã vạch (dãy >10 số) và các khoảng trắng dài sinh ra từ text căn lề phải
            name = re.sub(r'(?:\(\+84\)|0)\d{6,}.*', '', line_with_receiver).strip()
            name = re.sub(r'\d{10,}', '', name).strip()
            # Cắt bỏ phần dư thừa nếu OCR tạo ra nhiều khoảng trắng liên tiếp giữa Tên và Số (VD: Hoàng        014)
            name = re.split(r'\s{3,}', name)[0].strip()
            info["Tên người nhận"] = name
        
        # Nếu dòng đó rỗng, tìm tên ở dòng kế tiếp
        if not info.get("Tên người nhận"):
            next_lines = receiver_block_match.group(2).split('\n')
            for line in next_lines[:3]:
                clean_line = line.strip()
                if clean_line and not re.search(r'(?:\(\+84\)|0)\d{6,}', clean_line):
                    name = re.sub(r'\d{10,}', '', clean_line).strip()
                    name = re.split(r'\s{3,}', name)[0].strip()
                    if name:
                        info["Tên người nhận"] = name
                        break

    # 3. Trích xuất Địa chỉ (Nằm dưới Số điện thoại người nhận)
    phone_pattern = r'(?:\(\+84\)[0-9\*]+|0[0-9\*]{5,})'
    
    receiver_area_match = re.search(r'(?i)Người\s*nhận(.*)', text_without_sender, flags=re.DOTALL)
    if receiver_area_match:
        receiver_area_text = receiver_area_match.group(1)
        
        # Tìm SĐT đầu tiên sau chữ "Người nhận"
        phone_match = re.search(f"({phone_pattern}.*?\n)(.*)", receiver_area_text, re.IGNORECASE | re.DOTALL)
        
        if phone_match:
            addr_block = phone_match.group(2).strip()
            # Từ khóa dừng lấy địa chỉ
            split_keywords = r'(?i)\n\s*(?:Từ chối|Không tiền mặt|KHÔNG TIỀN MẶT|COD|Trọng\s*lượng|N/A|Order\s*ID|Tiktok|SKU|Qty|K\d+|Product Name|Mặc định|In transit|Thời gian)'
            addr_text_only = re.split(split_keywords, addr_block)[0]
            
            addr_lines = []
            for line in addr_text_only.split('\n'):
                # Xóa mã vạch nằm rải rác
                clean_l = re.sub(r'\b\d{10,}\b', '', line).strip()
                # Cắt bỏ đuôi "Không tiền mặt" nếu nó nằm ở cuối dòng của địa chỉ
                clean_l = re.split(r'(?i)\s{2,}KHÔNG TIỀN MẶT', clean_l)[0]
                
                if clean_l and not re.search(r'(?i)Người\s*mua\s*không', clean_l):
                    addr_lines.append(clean_l.strip())
            
            info["Địa chỉ"] = ', '.join(addr_lines)
        else:
            # Fallback nếu không quét được SĐT
            fallback_match = re.search(r'(?i)Người\s*nhận.*\n(.*)', text, re.IGNORECASE | re.DOTALL)
            if fallback_match:
                addr_block = fallback_match.group(1).strip()
                split_keywords = r'(?i)\n\s*(?:Từ chối|Không tiền mặt|KHÔNG TIỀN MẶT|COD|Trọng\s*lượng|N/A|Order\s*ID|Tiktok|SKU|Qty|K\d+|Product Name|Mặc định|In transit|Thời gian)'
                addr_text_only = re.split(split_keywords, addr_block)[0]
                addr_lines = [re.sub(r'\b\d{10,}\b', '', line).strip() for line in addr_text_only.split('\n') if line.strip() and not re.match(phone_pattern, line)]
                # Làm sạch thêm lần nữa
                clean_addrs = [re.split(r'(?i)\s{2,}KHÔNG TIỀN MẶT', l)[0].strip() for l in addr_lines]
                info["Địa chỉ"] = ', '.join(filter(None, clean_addrs))

    # Đảm bảo dọn dẹp kết quả tránh bị lẫn lộn
    if info["Địa chỉ"] and info["Tên người nhận"] in info["Địa chỉ"]:
        info["Địa chỉ"] = info["Địa chỉ"].replace(info["Tên người nhận"], "").strip(", ")

    return info

# Sidebar để chọn tính năng
st.sidebar.header("Tùy chọn")
option = st.sidebar.selectbox("Chọn loại tệp muốn đọc:", ["PDF", "Hình ảnh"])

extracted_data = []

if option == "PDF":
    st.header("1. Nhập liệu hàng loạt từ file PDF")
    st.markdown("Hệ thống sẽ lấy tự động **ID**, **Tên**, và **Địa chỉ** cho từng trang có trong PDF.")
    uploaded_file = st.file_uploader("Tải lên file PDF chứa nhãn mã vạch", type=["pdf"])
    
    if uploaded_file is not None:
        try:
            with st.spinner("Đang đọc file PDF và phân tích từng trang..."):
                doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    # Sắp xếp các chữ theo thứ tự dọc để không bị lẫn "Người gửi" và "Người nhận"
                    text = page.get_text("text", sort=True)
                    
                    if len(text.strip()) < 50:
                        try:
                            pix = page.get_pixmap(dpi=300)
                            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            text = pytesseract.image_to_string(image, lang='vie')
                        except Exception as e:
                            st.warning(f"Không thể OCR trang {page_num + 1}: {e}")
                    
                    if text.strip() != "":
                        info = extract_label_info(text)
                        info["Trang"] = page_num + 1
                        extracted_data.append(info)
                
                if not extracted_data:
                    st.warning("Không tìm thấy chữ hoặc không thể nhận diện được chữ trong file PDF này.")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi khi đọc PDF: {e}")

elif option == "Hình ảnh":
    st.header("2. Trích xuất từ Hình ảnh (sử dụng OCR Tiếng Việt)")
    st.markdown("⚠️ Yêu cầu Tesseract OCR (có tiếng Việt) đã được cài đặt trên máy.")
    
    uploaded_files = st.file_uploader("Tải lên hình ảnh (chọn nhiều ảnh nếu cần)", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)
    
    if uploaded_files:
        if st.button("Bắt đầu trích xuất dữ liệu"):
            with st.spinner("Đang nhận diện văn bản (OCR) từng ảnh..."):
                try:
                    for idx, file in enumerate(uploaded_files):
                        image = Image.open(file)
                        # Thực hiện OCR tiếng Việt
                        text = pytesseract.image_to_string(image, lang='vie')
                        
                        if text.strip() != "":
                            info = extract_label_info(text)
                            info["Tên File"] = file.name
                            extracted_data.append(info)
                            
                except FileNotFoundError:
                    st.error("❌ Không tìm thấy Tesseract OCR. Kiểm tra lại đường dẫn cài đặt.")
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi khi xử lý hình ảnh: {e}")

# Hiển thị và Xuất dữ liệu
if extracted_data:
    st.success(f"Đã trích xuất xong {len(extracted_data)} bản ghi!")
    
    # Tạo DataFrame
    df = pd.DataFrame(extracted_data)
    
    # Đưa cột Trang / Tên File lên đầu nếu có
    cols = df.columns.tolist()
    if "Trang" in cols:
        cols.insert(0, cols.pop(cols.index("Trang")))
    elif "Tên File" in cols:
        cols.insert(0, cols.pop(cols.index("Tên File")))
    
    df = df[cols]
    
    st.markdown("### Kết quả Trích xuất")
    # Hiển thị bảng
    df_display = df.drop(columns=["Văn bản gốc"], errors="ignore")
    st.dataframe(df_display, use_container_width=True)
    
    # Phần chia cột tải xuống
    col1, col2 = st.columns(2)
    
    with col1:
        # Cho phép tải xuống CSV
        csv = df_display.to_csv(index=False, encoding='utf-8-sig') # Dùng utf-8-sig để Excel đọc không lỗi font tiếng Việt
        st.download_button(
            label="📥 Tải xuống CSV",
            data=csv,
            file_name='du_lieu_don_hang.csv',
            mime='text/csv',
        )

    with col2:
        # Cho phép tải xuống Excel (.xlsx)
        # Sử dụng BytesIO để cache file excel trên bộ nhớ RAM trực tiếp cho Streamlit
        import io
        excel_buffer = io.BytesIO()
        df_display.to_excel(excel_buffer, index=False, engine='openpyxl')
        
        st.download_button(
            label="📊 Tải xuống Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name='du_lieu_don_hang.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    
    # Tùy chọn xem text gốc
    with st.expander("Xem văn bản gốc để đối chiếu lỗi"):
        for row in extracted_data:
            key = f"Trang {row.get('Trang')}" if 'Trang' in row else row.get('Tên File')
            st.text(f"--- {key} ---")
            st.text(row["Văn bản gốc"])
