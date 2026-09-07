import streamlit as st
import pandas as pd
import io

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Stock Material & Daily Monitor", layout="wide")

st.title("📦 สรุปรายการวัสดุและสารเคมีต่ำกว่า Safety Stock (Daily Report)")

# ส่วน Upload ไฟล์ Excel
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (Other_Material_Thailand...)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 1. โหลดข้อมูล Stock Material (แถวที่ 3 เป็นหัวตารางจริง header=2)
        stock_df = pd.read_excel(uploaded_file, sheet_name="Stock Material", header=2)
        sap_df = pd.read_excel(uploaded_file, sheet_name="SAP_ZRMM0004")
        
        # ตัดช่องว่างชื่อคอลัมน์
        stock_df.columns = [str(c).strip() for c in stock_df.columns]
        sap_df.columns = [str(c).strip() for c in sap_df.columns]
        
        # แปลงตัวเลขคอลัมน์ Safety Stock (H) และ Warehouse Stock (I)
        stock_df['Safety Stock'] = pd.to_numeric(stock_df['Safety Stock'], errors='coerce').fillna(0)
        stock_df['Warehouse Stock'] = pd.to_numeric(stock_df['Warehouse Stock'], errors='coerce').fillna(0)
        
        # 2. กรองเฉพาะรายการที่ Warehouse Stock (I) < Safety Stock (H)
        low_stock = stock_df[stock_df['Warehouse Stock'] < stock_df['Safety Stock']].copy()
        
        # 3. เตรียมข้อมูลจากชีท SAP_ZRMM0004
        sap_df['Mat. Number'] = sap_df['Mat. Number'].astype(str).str.strip()
        
        # ทำความสะอาดเลข PR / PO (ลบ .0 ออก)
        sap_df['PR Number'] = sap_df['PR Number'].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        sap_df['PO Number'] = sap_df['PO Number'].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        # รวมกลุ่ม PR/PO ตามรหัสวัสดุ
        sap_summary = {}
        for mat, grp in sap_df.groupby('Mat. Number'):
            prs = [p for p in grp['PR Number'].dropna().unique() if p and p.lower() != 'nan']
            pos = [p for p in grp['PO Number'].dropna().unique() if p and p.lower() != 'nan']
            sap_summary[mat] = {
                'PR': ", ".join(prs),
                'PO': ", ".join(pos)
            }
            
        # 4. ประกอบข้อมูลและเปรียบเทียบตามเงื่อนไข
        report_data = []
        for _, row in low_stock.iterrows():
            mat_code = str(row.get('Material Code', '')).strip()
            spec = row.get('Specification', '-')
            desc = row.get('รายละเอียด', '-')
            unit = row.get('Unit', '-')
            mat_grp = row.get('Mat.Group', '-')
            safety_stock = row.get('Safety Stock', 0)
            wh_stock = row.get('Warehouse Stock', 0)
            
            # ตรวจสอบ PR ในคอลัมน์ K (PR (1)) ของชีท Stock Material
            k_val = row.get('PR (1)', 0)
            has_k_pr = pd.notna(k_val) and str(k_val).strip() not in ['0', '', '0.0', 'nan', '-']
            
            # ดึงข้อมูลจาก SAP
            sap_info = sap_summary.get(mat_code, {'PR': '', 'PO': ''})
            pr_val = sap_info['PR']
            po_val = sap_info['PO']
            
            # เช็คว่ามี PR หรือไม่ (ทั้งจาก SAP หรือจากคอลัมน์ K)
            has_pr = bool(pr_val) or has_k_pr
            
            # กำหนดสถานะ Remind to buy
            remind_to_buy = "Follow PR&PO" if has_pr else "Buy"
                
            report_data.append({
                'Material Code': mat_code,
                'Specification': spec,
                'รายละเอียด': desc,
                'Unit': unit,
                'Mat.Group': mat_grp,
                'Safety Stock': round(safety_stock, 2),
                'Warehouse Stock': round(wh_stock, 2),
                'Remind to buy': remind_to_buy,
                'PR': pr_val if pr_val else ("-" if not has_k_pr else f"มีแจ้งเปิดแล้ว (ยอด {k_val})"),
                'PO': po_val if po_val else "-"
            })
            
        final_df = pd.DataFrame(report_data)
        
        # 5. แสดงสถิติภาพรวม (Metric Cards)
        c1, c2, c3 = st.columns(3)
        c1.metric("📌 รายการสต็อกต่ำกว่า Safety ทั้งหมด", f"{len(final_df)} รายการ")
        c2.metric("🚨 ต้องเปิด PR ด่วน (Buy)", f"{(final_df['Remind to buy'] == 'Buy').sum()} รายการ")
        c3.metric("⏳ ติดตามสถานะ (Follow PR&PO)", f"{(final_df['Remind to buy'] == 'Follow PR&PO').sum()} รายการ")
        
        st.write("---")
        
        # 6. ฟังก์ชันแต่งสีเซลล์ (รองรับทั้ง Pandas map และ applymap)
        def highlight_status(val):
            if val == 'Buy':
                return 'background-color: #ff4d4d; color: white; font-weight: bold; text-align: center;'
            elif val == 'Follow PR&PO':
                return 'background-color: #ffa600; color: black; font-weight: bold; text-align: center;'
            return ''

        styler = final_df.style
        if hasattr(styler, 'map'):
            styled_df = styler.map(highlight_status, subset=['Remind to buy'])
        else:
            styled_df = styler.applymap(highlight_status, subset=['Remind to buy'])
            
        styled_df = styled_df.format({'Safety Stock': '{:,.2f}', 'Warehouse Stock': '{:,.2f}'})
        
        # แสดงผลตารางบน Streamlit
        st.dataframe(styled_df, use_container_width=True, height=550)
        
        # 7. ปุ่มดาวน์โหลด Daily Report (Excel)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Daily_Low_Stock_Report')
        buffer.seek(0)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Daily Report ประจำวัน (.xlsx)",
            data=buffer,
            file_name="Daily_Material_Safety_Stock_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"⚠️ เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")
