import streamlit as st
from pdf2docx import Converter
import docx
from docx.oxml.ns import qn
from docx.shared import Pt
import os
import tempfile
import re

# ==========================================
# 1. CÁC HÀM KIỂM TRA VÀ XÓA WATERMARK TRÊN WORD
# ==========================================
def is_dark(hex_color, threshold=80):
    """Coi là chữ đen nếu màu 'auto' hoặc ba kênh RGB đều thấp."""
    if not hex_color or hex_color.lower() == "auto":
        return True
    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6:
        return True
    r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return max(r, g, b) <= threshold

def remove_non_black_text_safely(root):
    """
    Xóa chữ watermark có màu sáng nhưng giữ lại khoảng trắng
    để các chữ đen đứng hai bên watermark không bị sập dính vào nhau.
    """
    for run in list(root.iter(qn('w:r'))):
        t_elements = run.findall(qn('w:t'))
        texts = [t.text for t in t_elements if t.text]
        full_text = "".join(texts)
        
        # Nếu run chỉ chứa khoảng trắng -> giữ nguyên và bảo toàn khoảng trắng
        if full_text and full_text.strip() == "":
            for t in t_elements:
                t.set(qn('xml:space'), 'preserve')
            continue
            
        rpr = run.find(qn('w:rPr'))
        color = rpr.find(qn('w:color')) if rpr is not None else None
        if color is not None and not is_dark(color.get(qn('w:val'))):
            # Thay vì xóa hẳn run (dễ làm mất khoảng cách giữa 2 từ đen hai bên),
            # nếu run màu này nằm giữa các chữ, ta kiểm tra xem có khoảng trắng không
            if full_text and (full_text[0].isspace() or full_text[-1].isspace()):
                for i, t in enumerate(t_elements):
                    if i == 0:
                        t.text = " "
                        t.set(qn('xml:space'), 'preserve')
                    else:
                        t.text = ""
            else:
                run.getparent().remove(run)

def remove_large_graphics(root, min_cx=2_000_000):
    """Xóa hình lớn (watermark dạng ảnh). 2.000.000 EMU ≈ 5,5 cm chiều rộng."""
    for d in list(root.iter(qn('w:drawing'))):
        ext = next(d.iter(qn('wp:extent')), None)
        if ext is not None and int(ext.get('cx', 0)) >= min_cx:
            d.getparent().remove(d)
    for p in list(root.iter(qn('w:pict'))):
        p.getparent().remove(p)


# ==========================================
# 2. CÁC HÀM SỬA LỖI DÍNH CHỮ & CHUẨN HÓA FONT TIMES NEW ROMAN
# ==========================================
def fix_stuck_vietnamese_words(text):
    """
    Tách các từ tiếng Việt bị dính liền nhau bên trong cùng 1 chuỗi text,
    bao gồm cả Họ Tên (ThịKim -> Thị Kim) và các thuật ngữ trong Biên bản.
    """
    if not text:
        return text

    # 1. Tự động tách chữ thường dính liền với Chữ Hoa (VD: "ThịKim" -> "Thị Kim", "ThanhTrúc" -> "Thanh Trúc")
    uppercase_vn = "A-ZĐÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ"
    lowercase_vn = "a-zđáàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ"
    text = re.sub(f'([{lowercase_vn}])([{uppercase_vn}])', r'\1 \2', text)

    # 2. Thêm dấu cách sau dấu phẩy, chấm phẩy nếu bị dính chữ (VD: "đầy đủ,chính xác")
    text = re.sub(r',(?=[^\s\d])', ', ', text)
    text = re.sub(r';(?=[^\s])', '; ', text)
    
    # 3. Thêm dấu cách trước dấu mở ngoặc đơn (VD: "chính xác(do")
    text = re.sub(r'([^\s\(])\(', r'\1 (', text)

    # 4. Từ điển tách các cặp từ chữ thường bị dính trong PDF gốc
    replacements = [
        (r'quảsố', 'quả số'),
        (r'sốhóa', 'số hóa'),
        (r'Sốhóa', 'Số hóa'),
        (r'sốVBCC', 'số VBCC'),
        (r'SốVBCC', 'Số VBCC'),
        (r'Sốlần', 'Số lần'),
        (r'sốlần', 'số lần'),
        (r'dữliệu', 'dữ liệu'),
        (r'Dữliệu', 'Dữ liệu'),
        (r'sổgốc', 'sổ gốc'),
        (r'từsổ', 'từ sổ'),
        (r'đủcác', 'đủ các'),
        (r'trừcác', 'trừ các'),
        (r'đủchính', 'đủ chính'),
        (r'rõtrong', 'rõ trong'),
        (r'bướccủa', 'bước của'),
        (r'củaquy', 'của quy'),
        (r'quytrình', 'quy trình'),
        (r'trìnhsố', 'trình số'),
        (r'hóadữ', 'hóa dữ'),
        (r'liệuVBCC', 'liệu VBCC'),
        (r'bảođảm', 'bảo đảm'),
        (r'đảmđầy', 'đảm đầy'),
        (r'đầyđủ', 'đầy đủ'),
        (r'chínhxác', 'chính xác'),
        (r'xácthông', 'xác thông'),
        (r'thôngtin', 'thông tin'),
        (r'đượcnêu', 'được nêu'),
        (r'đượcsố', 'được số'),
        (r'hóachưa', 'hóa chưa'),
        (r'chưađầy', 'chưa đầy'),
        (r'dothông', 'do thông'),
        (r'tinbị', 'tin bị'),
        (r'bịmất', 'bị mất'),
        (r'Chỉ tiết', 'Chi tiết'),
    ]
    
    for _ in range(3):
        for pattern, repl in replacements:
            text = re.sub(pattern, repl, text)
            
    # Thu gọn nếu có nhiều dấu cách thừa liên tiếp
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text

def process_all_paragraphs_and_set_font(root):
    """
    Quét trực tiếp mọi thẻ đoạn văn <w:p> trong toàn bộ tài liệu (kể cả trong bảng và khung) để:
    1. Xóa thuộc tính ép co chữ (w:spacing, w:w, w:kern).
    2. Nối các run cùng định dạng bị pdf2docx cắt vụn và sửa lỗi dính chữ.
    3. Ép toàn bộ font chữ về Times New Roman.
    """
    # 1. Gỡ bỏ các thẻ ép khoảng cách ký tự và đặt font Times New Roman ở cấp độ XML
    for rpr in list(root.iter(qn('w:rPr'))):
        for tag in ('w:spacing', 'w:w', 'w:kern', 'w:position'):
            el = rpr.find(qn(tag))
            if el is not None:
                rpr.remove(el)
        
        # Đổi toàn bộ khai báo font của từng run sang Times New Roman
        rFonts = rpr.find(qn('w:rFonts'))
        if rFonts is not None:
            rFonts.set(qn('w:ascii'), 'Times New Roman')
            rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            rFonts.set(qn('w:cs'), 'Times New Roman')
            rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    # 2. Xử lý từng đoạn văn <w:p> trong toàn bộ cấu trúc XML
    for p_elem in list(root.iter(qn('w:p'))):
        # Lấy tất cả các thẻ chữ <w:t> bên trong đoạn văn này
        t_elems = list(p_elem.iter(qn('w:t')))
        if not t_elems:
            continue

        # Bước 2a: Sửa dính chữ ngay bên trong từng thẻ <w:t>
        for t in t_elems:
            if t.text:
                t.text = fix_stuck_vietnamese_words(t.text)
                t.set(qn('xml:space'), 'preserve')

        # Bước 2b: Kiểm tra 2 thẻ <w:t> đứng cạnh nhau (do pdf2docx tách mỗi từ 1 thẻ)
        # Nếu thẻ trước + thẻ sau ghép lại tạo thành từ bị dính (VD: thẻ 1 là "đủ", thẻ 2 là "các") -> chèn dấu cách!
        for i in range(len(t_elems) - 1):
            t1 = t_elems[i]
            t2 = t_elems[i + 1]
            text1 = t1.text or ""
            text2 = t2.text or ""
            
            if not text1 or not text2:
                continue
                
            # Nếu giữa t1 và t2 chưa hề có dấu cách
            if not text1[-1].isspace() and not text2[0].isspace():
                combined = text1 + text2
                fixed_combined = fix_stuck_vietnamese_words(combined)
                
                # Nếu hàm fix phát hiện chuỗi ghép bị dính (hoặc sau dấu phẩy không có cách)
                if fixed_combined != combined:
                    t1.text = text1 + " "
                    t1.set(qn('xml:space'), 'preserve')
                # Hoặc nếu t1 kết thúc bằng dấu phẩy/hai chấm mà t2 bắt đầu bằng chữ cái
                elif text1[-1] in (',', ';', ':') and text2[0].isalpha():
                    t1.text = text1 + " "
                    t1.set(qn('xml:space'), 'preserve')
                # Hoặc nếu t1 kết thúc bằng dấu gạch ngang '-' ở đầu dòng (VD: "-Số")
                elif text1.strip() == "-" and text2[0].isalnum():
                    t1.text = text1 + " "
                    t1.set(qn('xml:space'), 'preserve')


# ==========================================
# 3. HÀM CHUYỂN ĐỔI VÀ XỬ LÝ CHÍNH
# ==========================================
def convert_pdf_to_word_and_process(pdf_path, docx_path):
    # 1. Chuyển đổi PDF sang Word bằng cấu hình chuẩn của pdf2docx (không can thiệp PDF gốc để bảo toàn bảng)
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()
    
    # 2. Mở file Word để xử lý
    doc = docx.Document(docx_path)
    
    # Thiết lập font mặc định của style Normal về Times New Roman
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # Làm sạch toàn bộ: thân văn bản, bảng, textbox, đầu/chân trang
    roots = [doc.element.body]
    for s in doc.sections:
        for part in (s.header, s.footer):
            if not part.is_linked_to_previous:
                roots.append(part._element)
                
    for root in roots:
        remove_non_black_text_safely(root)
        remove_large_graphics(root)
        process_all_paragraphs_and_set_font(root)
    
    # 3. Tìm và di chuyển danh sách tên ở Mục 3 xuống cuối file
    names_to_move = []
    is_extracting_names = False
    paragraphs_to_delete = []
    
    for p in doc.paragraphs:
        text = p.text.strip()
            
        if "3. Họ tên những người thực hiện số hóa" in text:
            is_extracting_names = True
            names_to_move.append(text)
            paragraphs_to_delete.append(p)
            continue
            
        if is_extracting_names:
            if text.startswith("4. Thời gian") or text.startswith("II."):
                is_extracting_names = False
            else:
                names_to_move.append(text)
                paragraphs_to_delete.append(p)
    
    # Xóa hẳn các dòng tên ở phần trên để không bị trống dòng
    for p in paragraphs_to_delete:
        p._element.getparent().remove(p._element)
        
    # Ghi danh sách tên xuống cuối file Word với font Times New Roman
    if names_to_move:
        sep_p = doc.add_paragraph("\n" + "="*40)
        sep_p.runs[0].font.name = 'Times New Roman'
        
        title_p = doc.add_paragraph("DANH SÁCH NHỮNG NGƯỜI THỰC HIỆN SỐ HÓA (Đã di chuyển từ trên xuống):")
        title_p.runs[0].bold = True
        title_p.runs[0].font.name = 'Times New Roman'
        
        for name_line in names_to_move[1:]:
            if name_line:
                name_p = doc.add_paragraph(name_line)
                for r in name_p.runs:
                    r.font.name = 'Times New Roman'
                
    doc.save(docx_path)


# ==========================================
# 4. GIAO DIỆN STREAMLIT
# ==========================================
st.set_page_config(page_title="App Chuyển Đổi & Xử Lý PDF", page_icon="📄")
st.title("📄 Ứng dụng Chuyển PDF sang Word")
st.markdown("Hỗ trợ: Chuyển đổi chuẩn Times New Roman, Xoá Watermark thông minh, Sửa lỗi dính chữ, Di chuyển danh sách tên.")

if 'processed_file' not in st.session_state:
    st.session_state.processed_file = None

uploaded_file = st.file_uploader("Vui lòng tải lên file PDF của bạn", type=["pdf"])

if uploaded_file is not None:
    if st.button("Bắt đầu chuyển đổi"):
        with st.spinner('Đang xử lý, vui lòng chờ...'):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                pdf_path = tmp_pdf.name
                
            docx_path = pdf_path.replace(".pdf", ".docx")
            
            try:
                convert_pdf_to_word_and_process(pdf_path, docx_path)
                
                with open(docx_path, "rb") as f:
                    st.session_state.processed_file = f.read()
                    
                st.success("Chuyển đổi và xử lý thành công!")
            except Exception as e:
                st.error(f"Có lỗi xảy ra: {e}")
            finally:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                if os.path.exists(docx_path):
                    os.remove(docx_path)

if st.session_state.processed_file:
    st.download_button(
        label="Tải file Word về máy",
        data=st.session_state.processed_file,
        file_name="BienBanSoHoa_DaXuLy.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    
    if st.button("Xóa dữ liệu cũ để làm file mới"):
        st.session_state.processed_file = None
        st.rerun()