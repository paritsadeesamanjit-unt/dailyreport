import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Material Stock & PR/PO Daily Monitor", layout="wide")

st.title("📦 ระบบตรวจสอบสต็อกวัสดุและสถานะการสั่งซื้อ (Daily Report)")

# ส่วน Upload ไฟล์ Excel
uploaded_file = st.file_uploader("กรุณาอัปโหลดไฟล์ Excel (Material & SAP Data)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # โหลดข้อมูลจากทั้งสองชีท
        stock_df = pd.read_excel(uploaded_file, sheet_name="Stock Material")
        sap_df = pd.read_excel(uploaded_file, sheet_name="SAP_ZRMM0004")
        
        # ปรับชื่อคอลัมน์ตัดช่องว่าง
        stock_df.columns = [str(c).strip() for c in stock_df.columns]
        sap_df.columns = [str(c).strip() for c in sap_df.columns]
        
        # ค้นหาคอลัมน์รหัสวัสดุ (Material Code / Item Code)
        material_col_stock = next((c for c in stock_df.columns if any(k in c.lower() for k in ['material', 'item', 'code', 'part', 'รหัส'])), stock_df.columns[1])
        material_col_sap = next((c for c in sap_df.columns if any(k in c.lower() for k in ['material', 'item', 'code', 'part', 'รหัส'])), sap_df.columns[0])
        
        # ค้นหาคอลัมน์ PR และ PO ในชีท SAP
        pr_col_sap = next((c for c in sap_df.columns if 'pr' in c.lower() or 'purchase req' in c.lower() or 'banfn' in c.lower()), None)
        po_col_sap = next((c for c in sap_df.columns if 'po' in c.lower() or 'purchasing doc' in c.lower() or 'ebeln' in c.lower()), None)
        
        # เข้าถึงคอลัมน์ H (Index 7: Safety Stock), I (Index 8: Stock คงเหลือ), K (Index 10: PR Status/Qty) ตามลำดับ Index คอลัมน์ Excel
        col_h = stock_df.columns[7]   # คอลัมน์ H: Safety Stock
        col_i = stock_df.columns[8]   # คอลัมน์ I: Current Stock
        col_k = stock_df.columns[10]  # คอลัมน์ K: ยอดหรือสถานะ PR เดิม
        
        # แปลงข้อมูลตัวเลขเพื่อคำนวณ
        stock_df[col_h] = pd.to_numeric(stock_df[col_h], errors='coerce').fillna(0)
        stock_df[col_i] = pd.to_numeric(stock_df[col_i], errors='coerce').fillna(0)
        
        # 1. คัดกรองเฉพาะรายการที่วัสดุคงเหลือ (I) < Safety Stock (H)
        low_stock = stock_df[stock_df[col_i] < stock_df[col_h]].copy()
        
        # จัดการข้อมูล SAP เพื่อดึง PR / PO โดยกลุ่มตาม Material
        sap_df[material_col_sap] = sap_df[material_col_sap].astype(str).str.strip()
        
        sap_summary = {}
        for mat_id, group in sap_df.groupby(material_col_sap):
            prs = group[pr_col_sap].dropna().astype(str).unique().tolist() if pr_col_sap else []
            pos = group[po_col_sap].dropna().astype(str).unique().tolist() if po_col_sap else []
            sap_summary[mat_id] = {
                "PR_Numbers": ", ".join([p for p in prs if p != "" and p.lower() != "nan"]),
                "PO_Numbers": ", ".join([p for p in pos if p != "" and p.lower() != "nan"])
            }
            
        # 2. ฟังก์ชันประเมินสถานะและเชื่อมโยงข้อมูล SAP
        def evaluate_action(row):
            mat_id = str(row[material_col_stock]).strip()
            k_val = row[col_k]
            has_k_value = pd.notna(k_val) and str(k_val).strip() not in ["0", "", "0.0", "nan", "-"]
            
            sap_info = sap_summary.get(mat_id, {"PR_Numbers": "", "PO_Numbers": ""})
            pr_no = sap_info["PR_Numbers"]
            po_no = sap_info["PO_Numbers"]
            
            # ตรวจสอบการมีอยู่ของข้อมูลใน SAP
            if pr_no or po_no:
                if po_no:
                    status = "เปิด PO แล้ว (กำลังรอรับของ)"
                else:
                    status = "เปิด PR แล้ว (ยังไม่ได้เปิด PO)"
            else:
                if has_k_value:
                    status = "คอลัมน์ K มีการบันทึก PR แต่ไม่พบประวัติใน SAP"
                else:
                    status = "⚠️ แจ้งเปิด PR ด่วน (สต็อกต่ำกว่า Safety)"
                    
            return pd.Series([status, pr_no, po_no], index=["Action_Status", "SAP_PR_No", "SAP_PO_No"])
            
        # ประมวลผลสถานะ
        eval_results = low_stock.apply(evaluate_action, axis=1)
        report_df = pd.concat([low_stock, eval_results], axis=1)
        
        # จัดลำดับคอลัมน์ให้อ่านง่าย
        highlight_cols = [material_col_stock, col_i, col_h, col_k, "Action_Status", "SAP_PR_No", "SAP_PO_No"]
        remaining_cols = [c for c in report_df.columns if c not in highlight_cols]
        final_report = report_df[highlight_cols + remaining_cols]
        
        # แสดงผลสรุป Metric บน Streamlit
        st.subheader("📊 ภาพรวมรายการสต็อกต่ำกว่า Safety Stock")
        col1, col2, col3 = st.columns(3)
        col1.metric("รายการที่ต้องตรวจสอบทั้งหมด", len(final_report))
        col2.metric("รายการที่ต้องแจ้งเปิด PR", (final_report["Action_Status"].str.contains("แจ้งเปิด PR ด่วน")).sum())
        col3.metric("รายการที่มี PR/PO ใน SAP แล้ว", (final_report["Action_Status"].str.contains("เปิด")).sum())
        
        st.dataframe(final_report, use_container_width=True)
        
        # สร้างปุ่มสำหรับดาวน์โหลด Daily Report (Excel)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            final_report.to_excel(writer, index=False, sheet_name='Daily_Material_Alert')
        buffer.seek(0)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Daily Report ประจำวัน (.xlsx)",
            data=buffer,
            file_name="Daily_Material_Safety_Stock_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")