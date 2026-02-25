import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
# ==========================================
# القائمة الجانبية (Sidebar) - تقارير الإكسيل
# ==========================================
with st.sidebar:
    st.header("📥 تقارير الإدارة")
    st.info("حمل دراسة الجدوى الشاملة في صفحة إكسيل واحدة جاهزة للطباعة والعرض.")
    
    # 1. تجهيز بيانات الإكسيل
    df_invest_flexo = pd.DataFrame({
        "البند": ["ماكينة طباعة فلكسو CI (8 ألوان)", "ماكينة تركيب البليتات (Mounter)", "مبرد الهواء والكمبروسر", "إجمالي استثمار الفلكسو"],
        "التكلفة (ريال)": [8000000, 150000, 400000, 8550000]
    })

    df_invest_roto = pd.DataFrame({
        "البند": ["ماكينة طباعة روتوجرافيور (8 ألوان)", "غلاية الزيت الحراري (Thermal Boiler)", "معدات نقل وتخزين السلندرات", "إجمالي استثمار الروتو"],
        "التكلفة (ريال)": [9000000, 1500000, 300000, 10800000]
    })

    df_opex = pd.DataFrame({
        "بند التكلفة الشهرية": ["الرواتب والأجور", "الإيجار والمصاريف الإدارية", "فاتورة الطاقة (الماكينة + الغلاية)"],
        "التكلفة في الفلكسو (ريال)": [150000, 50000, 25000], 
        "التكلفة في الروتو (ريال)": [150000, 60000, 65000]  
    })

    df_scenario = pd.DataFrame({
        "عناصر تكلفة الطلبية (حجم 5 طن - 8 ألوان)": [
            "تكلفة المواد الخام", 
            "تكلفة التجهيز (بليتات مقابل سلندرات)", 
            "تكلفة هالك التشغيل والتجهيز", 
            "تكلفة المستهلكات (أنيلوكس/رول مطاطي)",
            "إجمالي تكلفة الطلبية",
            "تكلفة الطن الواحد"
        ],
        "تقنية الفلكسو (ريال)": [45000, 3200, 450, 200, 48850, 9770],
        "تقنية الروتو (ريال)": [45000, 12000, 2250, 150, 59400, 11880]
    })

    df_client_mix = pd.DataFrame({
        "الهيكل المطلوب للعميل": ["طبقة واحدة (38 mic label white / 40 mic clear)", "طبقتين (20 opp + 20 met)", "3 طبقات (12 pet + 7 alu + 50 pe)"],
        "النسبة من إجمالي الطلب": ["60%", "30%", "10%"],
        "سعر البيع المستهدف للعميل - فلكسو (ريال/كجم)": [12.0, 13.0, 15.0],
        "سعر البيع المستهدف للعميل - روتو (ريال/كجم)": [13.0, 13.5, 15.0]
    })

    # 2. إنشاء ملف الإكسيل في الذاكرة
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('دراسة الجدوى')
        worksheet.right_to_left() 
        
        # التنسيقات
        header_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        money_format = workbook.add_format({'num_format': '#,##0', 'border': 1, 'align': 'center'})
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#D9E1F2', 'align': 'center', 'border': 1})

        # كتابة الجداول
        worksheet.merge_range('A1:B1', '1. استثمار الفلكسو (CAPEX)', title_format)
        df_invest_flexo.to_excel(writer, sheet_name='دراسة الجدوى', startrow=1, startcol=0, index=False)
        
        worksheet.merge_range('D1:E1', '1. استثمار الروتو (CAPEX)', title_format)
        df_invest_roto.to_excel(writer, sheet_name='دراسة الجدوى', startrow=1, startcol=3, index=False)

        worksheet.merge_range('A8:C8', '2. التكاليف التشغيلية الشهرية للمصنع (OPEX)', title_format)
        df_opex.to_excel(writer, sheet_name='دراسة الجدوى', startrow=8, startcol=0, index=False)

        worksheet.merge_range('A14:C14', '3. سيناريو التكلفة (طلبية 5 طن - 8 ألوان)', title_format)
        df_scenario.to_excel(writer, sheet_name='دراسة الجدوى', startrow=14, startcol=0, index=False)

        worksheet.merge_range('A23:D23', '4. تحليل منتجات العميل (بناء على المعطيات)', title_format)
        df_client_mix.to_excel(writer, sheet_name='دراسة الجدوى', startrow=23, startcol=0, index=False)

        # ضبط عرض الأعمدة
        for col_num, value in enumerate(df_invest_flexo.columns.values):
            worksheet.write(1, col_num, value, header_format)
        for col_num, value in enumerate(df_invest_roto.columns.values):
            worksheet.write(1, col_num + 3, value, header_format)
        for col_num, value in enumerate(df_opex.columns.values):
            worksheet.write(8, col_num, value, header_format)
        for col_num, value in enumerate(df_scenario.columns.values):
            worksheet.write(14, col_num, value, header_format)
        for col_num, value in enumerate(df_client_mix.columns.values):
            worksheet.write(23, col_num, value, header_format)

        worksheet.set_column('A:A', 45)
        worksheet.set_column('B:B', 20, money_format)
        worksheet.set_column('D:D', 45)
        worksheet.set_column('E:E', 20, money_format)
        worksheet.set_column('C:C', 20, money_format)

    # 3. زر التحميل
    st.download_button(
        label="📥 تحميل الإكسيل (Excel)",
        data=buffer.getvalue(),
        file_name="مقارنة_فلكسو_روتو_رويان.xlsx",
        mime="application/vnd.ms-excel",
        use_container_width=True
    )
