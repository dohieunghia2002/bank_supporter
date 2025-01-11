import streamlit as st
import pandas as pd
import re

st.set_page_config(layout="wide")

# Tiêu đề ứng dụng
st.title("HỆ THỐNG HỖ TRỢ AGRIBANK")

# Cho phép người dùng upload nhiều file (accepts Excel formats)
# LP_T7-2024.xls
# DN_T7-2024.xls
# mau_template.xlsx
uploaded_files = st.file_uploader(
    "Upload Excel files (.xlsx, .xls)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

def load_branch_data(files):
    """
    Hàm xử lý và phân loại dữ liệu từ các file Excel đã tải lên.
    Trả về ba DataFrame: LP_dataframe, DN_dataframe, template_dataframe.
    """
    dataframes = {"LP": None, "DN": None, "template": None}
    
    for file in files:
        for key in dataframes:
            if key in file.name:
                dataframes[key] = pd.read_excel(file)
    
    return dataframes["LP"], dataframes["DN"], dataframes["template"]

def excel_column_to_index(col_name):
    """Chuyển đổi chỉ số cột Excel (A, B, C, ...) thành số thứ tự cột (0, 1, 2, ...)."""
    col_index = 0
    for char in col_name:
        col_index = col_index * 26 + (ord(char.upper()) - ord('A') + 1)
    return col_index - 1

def normalize_lookup_value(value):
    if isinstance(value, int):  # Nếu là số nguyên, giữ nguyên
        return str(value)
    return re.sub(r'\D', '', value)

def calculate(template_df, brand_col_name):
    access_col_name = template_df.filter(regex=r'(?i)ma\s*so').columns[0]
    template_df[access_col_name] = template_df[access_col_name].fillna('')
    print(template_df[access_col_name])
    start_index = template_df[access_col_name].astype(str).str.match(r'^[+-]\s*\d+$').idxmax()

    result = 0
    for i in range(start_index, len(template_df)):
        access_code = str(template_df.loc[i, access_col_name])
        ket_qua_lp = template_df.loc[i, brand_col_name]
        
        if pd.notna(ket_qua_lp) and access_code.startswith(('+', '-')):
            value = int(access_code.replace('+', '').replace('-', '').strip())
            
            # Cộng hoặc trừ tùy thuộc vào dấu
            if access_code.startswith('+'):
                result += ket_qua_lp
            elif access_code.startswith('-'):
                result -= ket_qua_lp
    
    last_index = len(template_df) - 1
    if pd.notna(template_df.loc[last_index, brand_col_name]):
        # Nếu có giá trị, thêm hàng mới
        result_df = pd.DataFrame({brand_col_name: [result]})
        template_df = pd.concat([template_df, result_df], ignore_index=True)
    else:
        # Nếu ô cuối cùng rỗng, cập nhật giá trị vào ô đó
        template_df.loc[last_index, brand_col_name] = result

    return template_df

def update_sumtax_difference(template_df):
    template_df.at[0, 'TONG THUE'] = int(template_df['Ket qua LP'].iloc[0] + template_df['Ket qua DN'].iloc[0])
    print(template_df['TONG THUE'].iloc[0])

    # Dò tìm ô có giá trị 'EV', 'ev', hoặc 'eV' trong cột 'TONG THUE'
    ev_row = template_df[template_df['TONG THUE'].astype(str).str.lower() == 'ev']
    if not ev_row.empty:
        ev_index = ev_row.index[0]
        ev_next_value = template_df.iloc[ev_index + 1, template_df.columns.get_loc('TONG THUE')]
    else:
        print("Không tìm thấy giá trị 'EV', 'ev', hoặc 'eV' trong cột 'TONG THUE'")
        return template_df

    # Dò tìm ô có giá trị 'CHENH LECH'
    chenh_lech_row = template_df[template_df['TONG THUE'].astype(str).str.lower() == 'chenh lech']
    if not chenh_lech_row.empty:
        chenh_lech_index = chenh_lech_row.index[0]
        # Cập nhật giá trị cho ô phía sau 'CHENH LECH'
        new_value = int(template_df.at[0, 'TONG THUE'] - ev_next_value)
        template_df.iloc[chenh_lech_index + 1, template_df.columns.get_loc('TONG THUE')] = new_value
    else:
        print("Không tìm thấy giá trị 'CHENH LECH' trong cột 'TONG THUE'")
    
    return template_df


def find_and_update_results(template_df, row_index, LP_df, DN_df, access_code, cot_moc, data_lp_col, data_dn_col):
    """ 
    Hàm tìm kiếm và cập nhật kết quả dựa trên mã số truy xuất, cột mốc và cột dữ liệu cần lấy từ LP và DN.
    """
    cot_moc_index = excel_column_to_index(cot_moc)
    data_lp_index = excel_column_to_index(data_lp_col)
    data_dn_index = excel_column_to_index(data_dn_col)

    lp_row_index = LP_df.index[LP_df.iloc[:, cot_moc_index] == access_code].tolist()
    dn_row_index = DN_df.index[DN_df.iloc[:, cot_moc_index] == access_code].tolist()

    if len(lp_row_index) > 0:
        ket_qua_lp = LP_df.iloc[lp_row_index[0], data_lp_index]
    else:
        ket_qua_lp = 0

    if len(dn_row_index) > 0:
        ket_qua_dn = DN_df.iloc[dn_row_index[0], data_dn_index]
    else:
        ket_qua_dn = 0
       
    template_df.at[row_index, 'Ket qua LP'] = ket_qua_lp
    template_df.at[row_index, 'Ket qua DN'] = ket_qua_dn

    return template_df

if uploaded_files:
    if len(uploaded_files) != 3:
        st.error("Lỗi tải lên file, bạn cần tải đúng 3 files: 1 file dữ liệu của chi nhánh LP, 1 file dữ liệu của chi nhánh ĐN, 1 file template.")
    else:
        
        # Gọi hàm để xử lý dữ liệu
        LP_dataframe, DN_dataframe, template_dataframe = load_branch_data(uploaded_files)
        
        # Hàm hiển thị dữ liệu
        def display_dataframe(df, branch_name):
            if df is not None:
                st.write(f"Dữ liệu chi nhánh {branch_name}:")
                st.dataframe(df)
            else:
                st.error(f"Không tìm thấy file dữ liệu chi nhánh {branch_name}.")
        
        # Hiển thị dữ liệu của từng chi nhánh
        display_dataframe(LP_dataframe, "LP")
        display_dataframe(DN_dataframe, "ĐN")
        display_dataframe(template_dataframe, "template")

        # Thêm nút "Thực hiện" để bắt đầu xử lý
        if st.button("Thực hiện"):
            # Kiểm tra và hiển thị mã số truy xuất từ file template
            if template_dataframe is not None and 'Ma so truy xuat' in template_dataframe.columns:
                valid_values = template_dataframe['Ma so truy xuat'].dropna().tolist()

                if len(valid_values) > 0:
                    for val in valid_values:
                        access_code = val
                        row_indices = template_dataframe.index[template_dataframe['Ma so truy xuat'] == access_code].tolist()
                        for row_index in row_indices:
                            cot_moc = template_dataframe.at[row_index, 'Cot moc']
                            data_lp_col = template_dataframe.at[row_index, 'DataLP']
                            data_dn_col = template_dataframe.at[row_index, 'DataDN']
                        
                            access_code = normalize_lookup_value(access_code)
                            template_dataframe = find_and_update_results(template_dataframe, row_index, LP_dataframe, DN_dataframe, float(access_code), cot_moc, data_lp_col, data_dn_col)

                
                template_dataframe = calculate(template_dataframe, 'Ket qua LP')
                template_dataframe = calculate(template_dataframe, 'Ket qua DN')
                template_dataframe = update_sumtax_difference(template_dataframe)
                st.write("Kết quả sau khi cập nhật:")
                st.dataframe(template_dataframe, use_container_width=True)

                # Tạo file Excel từ DataFrame kết quả
                output_file = 'Ket_qua_ho_tro_Agribank.xlsx'
                template_dataframe.to_excel(output_file, index=False)

                # Thêm nút tải xuống file Excel
                with open(output_file, "rb") as file:
                    btn = st.download_button(
                        label="Tải xuống kết quả",
                        data=file,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.write("Cột 'Ma so truy xuat' không có trong file template.")
