import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import io
import os
import pandas as pd
import html  # Стандартный модуль для безопасного экранирования спецсимволов
from datetime import datetime
import urllib.request
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Импорт компонентов ReportLab для верстки отчетов
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="NovaHeat Simulator", layout="wide")
import matplotlib.font_manager as fm

if os.path.exists("assets/arial.ttf"):
    fm.fontManager.addfont("assets/arial.ttf")
st.title("Термодинамический симулятор NovaHeat")
st.subheader("Расчет термодинамики массива конструкции")

# --- ИНТЕРФЕЙС ВВОДА МЕТАДАННЫХ ПРОЕКТА С ДЕФОЛТАМИ ---
st.header("Информация о проекте")
meta_col1, meta_col2, meta_col3 = st.columns(3)
with meta_col1:
    project_name = st.text_input("Наименование проекта", value="Проект")
with meta_col2:
    customer_name = st.text_input("Заказчик", value="")
with meta_col3:
    done_by = st.text_input("Исполнитель", value="")
# --- БОКОВАЯ ПАНЕЛЬ С ИСХОДНЫМИ ДАННЫМИ ---
st.sidebar.header("Исходные данные")

st.sidebar.subheader("1. Параметры нагревательного кабеля")
cable_name = st.sidebar.text_input("Греющий кабель", value="Контур-СВ64Ф")
voltage = st.sidebar.number_input("Номинальное напряжение, В", min_value=110, max_value=400, value=230)
P_line = st.sidebar.number_input("Номинальная мощность кабеля, Вт/м", min_value=1.0, max_value=200.0, value=64.0)

# Высокоточная двусторонняя синхронизация шага укладки без сброса
if 'spacing' not in st.session_state:
    st.session_state['spacing'] = 150

def update_spacing_slider():
    st.session_state['spacing'] = st.session_state['sp_num']

def update_spacing_num():
    st.session_state['spacing'] = st.session_state['sp_slider']

sp_col1, sp_col2 = st.sidebar.columns(2)
with sp_col1:
    spacing = st.slider("Шаг укладки кабеля, мм", min_value=5, max_value=1000, step=5, key='sp_slider', on_change=update_spacing_num, value=st.session_state['spacing'])
with sp_col2:
    spacing = st.number_input("Значение, мм", min_value=5, max_value=1000, step=5, key='sp_num', on_change=update_spacing_slider, value=st.session_state['spacing'])

# Высокоточная двусторонняя синхронизация глубины монтажа без сброса
if 'depth' not in st.session_state:
    st.session_state['depth'] = 55

def update_depth_slider():
    st.session_state['depth'] = st.session_state['dp_slider']

def update_depth_num():
    st.session_state['depth'] = st.session_state['dp_num']

dp_col1, dp_col2 = st.sidebar.columns(2)
with dp_col1:
    depth = st.slider("Глубина монтажа, мм", min_value=5, max_value=1000, step=5, key='dp_slider', on_change=update_depth_num, value=st.session_state['depth'])
with dp_col2:
    depth = st.number_input("Значение глубины, мм", min_value=5, max_value=1000, step=5, key='dp_num', on_change=update_depth_slider, value=st.session_state['depth'])

power_multiplier = st.sidebar.number_input("Коэффициент запаса мощности", min_value=0.5, max_value=2.0, value=1.0, step=0.1)

st.sidebar.subheader("2. Параметры и условия окружающей среды")
st.sidebar.caption("Выберите режим для расчета теплопотерь:")

calc_mode = st.sidebar.radio("Режим расчета теплопотерь внешнего слоя массива конструкции:", ["По исходным климатическим условиям", "Принудительный ввод коэффициента теплопроводности верхнего слоя α"])

T_amb = st.sidebar.number_input("Температура окружающей среды (воздух, °C)", min_value=-60.0, max_value=10.0, value=-15.0)

if calc_mode == "По исходным климатическим условиям":
    wind_unit = st.sidebar.selectbox("Единица измерения скорости ветра:", ["км/ч", "м/с"])
    raw_wind = st.sidebar.number_input("Скорость ветра", min_value=0.0, max_value=150.0, value=8.0)
    
    if wind_unit == "км/ч":
        wind_speed_kmh = raw_wind
        wind_ms = raw_wind / 3.6
    else:
        wind_speed_kmh = raw_wind * 3.6
        wind_ms = raw_wind
        
    snow_rate = st.sidebar.number_input("Интенсивность снегопада, мм/ч", min_value=0.0, max_value=50.0, value=13.0)
    free_area_ratio = st.sidebar.number_input("Коэффициент свободной площади (FAR)", min_value=0.0, max_value=1.0, value=1.00)
    
    alpha_top = 5.8 + 4.1 * wind_ms
    q_snow = (snow_rate / (3600.0 * 1000.0)) * 100.0 * 334000.0 * free_area_ratio
    condition_text = "Наличие ветра и снегопада"
else:
    wind_speed_kmh = 0.0
    snow_rate = 0.0
    free_area_ratio = 1.00
    q_snow = 0.0
    alpha_top = st.sidebar.number_input("Коэффициент теплоотдачи наружной поверхности, Вт/(м²·°C)", min_value=1.0, max_value=100.0, value=19.476)
    condition_text = "Наличие ветра"
st.sidebar.subheader("3. Параметры нижнего слоя массива конструкции")
T_ground = st.sidebar.number_input("Температура грунта, °C", min_value=-5.0, max_value=15.0, value=2.0)

st.sidebar.subheader("4. Теплофизические свойства слоев массива конструкции")
st.sidebar.markdown("**Слой 1: Верхний (внешний) слой массива конструкции**")

material_choice_1 = st.sidebar.selectbox("Материал внешнего верхнего слоя:", ["Искусственный газон", "Натуральный газон", "Бетонная стяжка"])
if material_choice_1 == "Искусственный газон": def_h1, def_k1 = 15.0, 0.035
elif material_choice_1 == "Натуральный газон": def_h1, def_k1 = 30.0, 0.550
else: def_h1, def_k1 = 80.0, 1.400

h_carpet = st.sidebar.number_input("Толщина верхнего (внешнего) слоя массива конструкции, мм", min_value=1.0, max_value=500.0, value=def_h1, key='h1_input')
k_carpet = st.sidebar.number_input("Теплопроводность внешнего слоя массива конструкции, Вт/(м·°C)", value=def_k1, format="%.3f", key='k1_input')

st.sidebar.markdown("**Слой 2: Средний слой массива конструкции**")
material_choice_2 = st.sidebar.selectbox("Материал среднего расчетного слоя:", ["Песчаная подготовка", "Мелкофракционный гравий", "Экструдированный пенополистирол"])
if material_choice_2 == "Песчаная подготовка": def_h2, def_k2 = 40.0, 1.255
elif material_choice_2 == "Мелкофракционный гравий": def_h2, def_k2 = 50.0, 0.380
else: def_h2, def_k2 = 50.0, 0.032

h_sand = st.sidebar.number_input("Толщина среднего слоя массива конструкции, мм", min_value=5.0, max_value=500.0, value=def_h2, key='h2_input')
k_sand = st.sidebar.number_input("Теплопроводность среднего слоя массива конструкции, Вт/(м·°C)", value=def_k2, format="%.3f", key='k2_input')

st.sidebar.markdown("**Слой 3: Нижний слой массива конструкции**")
material_choice_3 = st.sidebar.selectbox("Материал нижнего слоя основания:", ["Щебеночное основание", "Естественный грунт (суглинок)", "Железобетонная плита"])
if material_choice_3 == "Щебеночное основание": def_h3, def_k3 = 100.0, 0.235
elif material_choice_3 == "Естественный грунт (суглинок)": def_h3, def_k3 = 200.0, 1.150
else: def_h3, def_k3 = 150.0, 1.650

h_gravel = st.sidebar.number_input("Толщина нижнего слоя массива конструкции, мм", min_value=10.0, max_value=1000.0, value=def_h3, key='h3_input')
k_gravel = st.sidebar.number_input("Теплопроводность нижнего слоя массива конструкции, Вт/(м·°C)", value=def_k3, format="%.3f", key='k3_input')

# Возвращена базовая физика полного шага W
W = spacing / 1000.0  
H = (h_carpet + h_sand + h_gravel) / 1000.0  

dx = dy = 0.0010

nx = int(W / dx) + 1
ny = int(H / dy) + 1

K = np.zeros((ny, nx))
y_carpet_end = h_carpet / 1000.0
y_sand_end = (h_carpet + h_sand) / 1000.0
for j in range(ny):
    y = j * dy
    if y <= y_carpet_end: K[j, :] = k_carpet
    elif y <= y_sand_end: K[j, :] = k_sand
    else: K[j, :] = k_gravel

# Возвращено положение кабеля строго по центру расчетной ячейки
cable_x_idx = int((W / 2.0) / dx)
cable_y_idx = int((depth / 1000.0) / dy)
cable_y_idx = min(max(1, cable_y_idx), ny - 2)
q_v = (P_line * power_multiplier) / (dx * dy)
if st.button("Выполнить термодинамический расчет"):
    with st.spinner("Моделирование установившегося теплового режима"):
        T_calc = np.zeros((ny, nx))
        
        K_n = 0.5 * (K[:-2, 1:-1] + K[1:-1, 1:-1])
        K_s = 0.5 * (K[2:, 1:-1] + K[1:-1, 1:-1])
        K_e = 0.5 * (K[1:-1, 2:] + K[1:-1, 1:-1])
        K_w = 0.5 * (K[1:-1, :-2] + K[1:-1, 1:-1])
        denom = K_n + K_s + K_e + K_w
        
        for it in range(15000):
            T_old = T_calc.copy()
            
            # Стабильный векторизованный пересчет всей матрицы конструкции
            T_calc[1:-1, 1:-1] = (K_n * T_calc[:-2, 1:-1] + K_s * T_calc[2:, 1:-1] + 
                                  K_e * T_calc[1:-1, 2:] + K_w * T_calc[1:-1, :-2]) / denom
            
            T_calc[cable_y_idx, cable_x_idx] += (q_v * dx * dy) / denom[cable_y_idx - 1, cable_x_idx - 1]
            
            T_calc[:, 0] = T_calc[:, 1]
            T_calc[:, -1] = T_calc[:, -2]
            
            T_calc[0, :] = (K[1, :] * T_calc[1, :] + alpha_top * dy * T_amb - q_snow * dy) / (K[1, :] + alpha_top * dy)
            T_calc[-1, :] = T_ground
            
            if np.max(np.abs(T_calc - T_old)) < 1e-4: 
                break
        
        st.session_state['calculated'] = True
        st.session_state['min_temp'] = np.min(T_calc[0, :])
        st.session_state['max_temp'] = np.max(T_calc[0, :])
        st.session_state['cable_temp'] = T_calc[cable_y_idx, cable_x_idx]
        st.session_state['T_array'] = T_calc.copy()
        
        # Высокоточный расчет плотностей тепловых потоков для стационарного режима
        f_top_calibrated = np.mean(k_carpet * (T_calc[1, :] - T_calc[0, :]) / dy) + q_snow
        f_bottom_calibrated = np.mean(k_gravel * (T_calc[-2, :] - T_calc[-1, :]) / dy)
        
        st.session_state['flux_top'] = f_top_calibrated
        st.session_state['flux_bottom'] = f_bottom_calibrated
        # Прецизионная сумма теплового баланса: система равна сумме уходящих потоков вверх и вниз
        st.session_state['flux_sys'] = f_top_calibrated + f_bottom_calibrated

        x_indices = np.linspace(0, cable_x_idx, 11, dtype=int)
        st.session_state['x_coords_mm'] = [float(i * dx * 1000) for i in x_indices]
        
        target_depths_mm = [0.00, 3.75, 7.50, 11.25, 15.00, 25.00, 35.00, 45.00, float(depth), 80.00, 105.00, 130.00, 155.00]
        target_depths_mm = sorted(list(set(target_depths_mm)))
        
        matrix_rows = []
        for d_mm in target_depths_mm:
            j_idx = int((d_mm / 1000.0) / dy)
            j_idx = min(j_idx, ny - 1)
            y_curr = j_idx * dy
            if y_curr <= y_carpet_end: mat_name = material_choice_1
            elif y_curr <= y_sand_end: mat_name = material_choice_2
            else: mat_name = material_choice_3
            
            if d_mm == depth: idx_str = "0"
            elif d_mm < depth:
                curr_above_idx = sum(1 for d in target_depths_mm if d_mm <= d < depth)
                idx_str = f"+{curr_above_idx}"
            else:
                curr_below_idx = sum(1 for d in target_depths_mm if depth < d <= d_mm)
                idx_str = f"-{curr_below_idx}"
            
            row = [f"{d_mm:.2f}", idx_str, mat_name]
            for i_idx in x_indices:
                val = T_calc[j_idx, i_idx]
                if d_mm == depth and i_idx == cable_x_idx: row.append(f">> [{val:.1f}]")
                else: row.append(f"{val:.1f}")
            matrix_rows.append(row)
        st.session_state['matrix_rows'] = matrix_rows
def generate_pdf_report(meta, img_buf1, img_buf2):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=50, bottomMargin=36, title="Тепловая модель Термо-Нова")
    story = []
    
    p_name_esc = html.escape(meta['p_name'])
    c_name_esc = html.escape(meta['c_name'])
    done_by_esc = html.escape(meta['done_by'])
    c_name_m_esc = html.escape(meta['c_name_m'])
    m1_esc = html.escape(meta['m1'])
    m2_esc = html.escape(meta['m2'])
    m3_esc = html.escape(meta['m3'])
    
    fn, fb = 'Helvetica', 'Helvetica-Bold'
    if os.path.exists("assets/arial.ttf") and os.path.exists("assets/arialbd.ttf"):
        try:
            pdfmetrics.registerFont(TTFont('CustomArial', "assets/arial.ttf"))
            pdfmetrics.registerFont(TTFont('CustomArial-Bold', "assets/arialbd.ttf"))
            fn, fb = 'CustomArial', 'CustomArial-Bold'
        except: pass


    styles = getSampleStyleSheet()
    s_title = ParagraphStyle('T', fontName=fb, fontSize=12, leading=16, alignment=1)
    s_norm = ParagraphStyle('N', fontName=fn, fontSize=8.5, leading=11)
    s_bold = ParagraphStyle('B', fontName=fb, fontSize=8.5, leading=11)
    s_warn = ParagraphStyle('W', fontName=fb, fontSize=8, leading=11, textColor=colors.HexColor('#B91C1C'))
    
    def add_later_headers(canvas, doc_obj):
        canvas.saveState()
        sm_logo = os.path.join("assets", "logo.png")
        if os.path.exists(sm_logo): canvas.drawImage(sm_logo, 36, 750, width=18, height=18)
        canvas.setFont(fn, 8)
        canvas.setFillColor(colors.HexColor('#475569'))
        canvas.drawString(60, 756, f"Моделирование двумерного температурного поля — {p_name_esc}")
        canvas.drawRightString(576, 756, f"Страница {doc_obj.page}")
        canvas.setStrokeColor(colors.HexColor('#CBD5E1'))
        canvas.line(36, 744, 576, 744)
        canvas.restoreState()

    lg_path = os.path.join("assets", "logotype.png")
    date_txt = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
    
    if os.path.exists(lg_path): hdr_cells = [[Image(lg_path, width=130, height=20), Paragraph(date_txt, ParagraphStyle('R', fontName=fn, fontSize=8.5, alignment=2))]]
    else: hdr_cells = [[Paragraph("<b>NovaHeat</b>", s_bold), Paragraph(date_txt, ParagraphStyle('R', fontName=fn, fontSize=8.5, alignment=2))]]
    
    hdr_table = Table(hdr_cells, colWidths=[270, 270])
    hdr_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,-1), 8)]))
    story.append(hdr_table)
    
    story.append(Paragraph("NovaHeat. Отчет о результатах моделирования двумерного температурного поля", s_title))
    story.append(Spacer(1, 10))
    
    w_txt = "<b>ВАЖНО!</b><br/>Пользователь программного комплекса NovaHeat несет полную ответственность за верификацию проектных решений по системе напольного электрообогрева. Программа NovaHeat производит исключительно расчет распределения теплоты; проверка корректности исходных данных, теплофизической модели или финальных выходных параметров не осуществляется."
    w_tbl = Table([[Paragraph(w_txt, s_warn)]], colWidths=[540])
    w_tbl.setStyle(TableStyle([('PADDING', (0,0), (-1,-1), 5), ('BACKGROUND', (0,0), (0,0), colors.HexColor('#FEF2F2')), ('BOX', (0,0), (0,0), 0.5, colors.HexColor('#FCA5A5'))]))
    story.append(w_tbl)
    story.append(Spacer(1, 10))
    
    meta_data = [
        [Paragraph("• Название проекта:", s_norm), Paragraph(f"<b>{p_name_esc}</b>", s_norm)],
        [Paragraph("• Заказчик:", s_norm), Paragraph(f"<b>{c_name_esc}</b>", s_norm)],
        [Paragraph("• Расчет выполнен:", s_norm), Paragraph(f"<b>{done_by_esc}</b>", s_norm)]
    ]
    mt = Table(meta_data, colWidths=[130, 410])
    mt.setStyle(TableStyle([('PADDING', (0,0), (-1,-1), 1.5)]))
    story.append(mt)
    
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.gray, spaceAfter=8))
    story.append(Paragraph(f"<i><b>{p_name_esc}</b></i>", ParagraphStyle('C', fontName=fb, fontSize=10, alignment=1)))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.gray, spaceBefore=6, spaceAfter=15))
    
    story.append(Paragraph("<b>ИСХОДНЫЕ ДАННЫЕ ДЛЯ МОДЕЛИРОВАНИЯ</b>", s_bold))
    story.append(Spacer(1, 8))
    
    design_cells = [
        [Paragraph("<b>Параметры нагревательного кабеля</b>", s_bold), Paragraph("<b>Параметры и условия окружающей среды</b>", s_bold)],
        [Paragraph(f"• Греющий кабель: {c_name_m_esc}<br/>• Номинальное напряжение: {meta['voltage']} В<br/>• Способ монтажа: Укладка в массив плиты (монолитный обогрев)<br/>• Шаг укладки кабеля: {meta['spacing']} мм<br/>• Глубина монтажа: {meta['depth']} мм<br/>• Коэффициент запаса мощности: {meta['p_mult']}<br/>• Тип расчета: Расчет установившегося термодинамического режима при фиксированном шаге укладки", s_norm),
         Paragraph(f"• Расчетные условия: {meta['cond_txt']}<br/>• Температура окружающей среды (воздух): {meta['T_amb']} °C<br/>• Скорость ветра: {meta['wind']:.1f} км/ч<br/>• Интенсивность снегопада: {meta['snow']} мм/ч<br/>• Коэффициент свободной площади: {meta['far']:.2f}<br/>• Коэффициент теплоотдачи наружной поверхности: {meta['alpha']:.3f} Вт/(м²·°C)", s_norm)],
        [Paragraph("", s_norm), Paragraph("<b>Расчетные параметры нижней поверхности массива конструкции</b>", s_bold)],
        [Paragraph("", s_norm), Paragraph(f"• Граничные условия: Грунтовое основание<br/>• Температура грунта: {meta['T_ground']} °C", s_norm)]
    ]
    dt = Table(design_cells, colWidths=[270, 270])
    dt.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 4)]))
    story.append(dt)
    
    story.append(Spacer(1, 25))
    story.append(Paragraph("<b>Теплофизические свойства материалов слоев конструкции</b> (Параметры разбивки слоев: Автоматическая)", s_norm))
    story.append(Spacer(1, 8))
    
    m_cells = [
        [Paragraph("Материал слоя массивов", s_bold), Paragraph("Толщина (мм)", s_bold), Paragraph("Теплопроводность (Вт/(м·°C))", s_bold), Paragraph("Плотность (кг/м³)", s_bold), Paragraph("Теплоемкость (кДж/(кг·°C))", s_bold)],
        [Paragraph(f"{m1_esc}", s_norm), Paragraph(f"{meta['h_c']:.2f}", s_norm), Paragraph(f"{meta['k_c']:.3f}", s_norm), Paragraph("1800.00", s_norm), Paragraph("0.840", s_norm)],
        [Paragraph(f"{m2_esc}", s_norm), Paragraph(f"{meta['h_s']:.2f}", s_norm), Paragraph(f"{meta['k_s']:.3f}", s_norm), Paragraph("1777.998", s_norm), Paragraph("0.837", s_norm)],
        [Paragraph(f"{m3_esc}", s_norm), Paragraph(f"{meta['h_g']:.2f}", s_norm), Paragraph(f"{meta['k_g']:.3f}", s_norm), Paragraph("1457.638", s_norm), Paragraph("0.837", s_norm)]
    ]
    mtbl = Table(m_cells, colWidths=[160, 80, 100, 100, 100])
    mtbl.setStyle(TableStyle([('LINEBELOW', (0,0), (-1,0), 0.5, colors.black), ('PADDING', (0,0), (-1,-1), 4), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(mtbl)
    
    story.append(PageBreak())
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>РЕЗУЛЬТАТЫ ТЕРМОДИНАМИЧЕСКОГО РАСЧЕТА</b>", s_bold))
    story.append(Spacer(1, 8))
    
    res_txt = f"• Линейная мощность греющего кабеля: {meta['P_line']*meta['p_mult']:.1f} Вт/м при температуре оболочки {meta['cable_temp']:.1f} °C (Коэффициент запаса мощности = {meta['p_mult']})<br/>" \
              f"• Плотность теплового потока, вырабатываемого системой (Удельная тепловая мощность): {meta['flux_sys']:.1f} Вт/м²<br/>" \
              f"• Плотность теплового потока через верхнюю поверхность: {meta['flux_top']:.1f} Вт/м²<br/>" \
              f"• Плотность теплового потока через нижнюю поверхность: {meta['flux_bottom']:.1f} Вт/м²<br/>" \
              f"• Минимальная температура наружной поверхности: {meta['min_temp']:.2f} °C<br/>" \
              f"• Максимальная температура наружной поверхности: {meta['max_temp']:.2f} °C"
    story.append(Paragraph(res_txt, s_norm))
    
    story.append(Spacer(1, 20))
    story.append(Image(img_buf1, width=540, height=260))
    
    story.append(Spacer(1, 25))
    story.append(Image(img_buf2, width=540, height=260))
    
    doc.build(story, onLaterPages=add_later_headers)
    buffer.seek(0)
    return buffer
def generate_excel_report(meta):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_rows = [
            ["Отчет о результатах термодинамического моделирования NovaHeat", ""],
            ["Дата генерации:", datetime.now().strftime('%d.%m.%Y %H:%M:%S')],
            ["Наименование проекта:", meta['p_name']],
            ["Заказчик:", meta['c_name']],
            ["Расчет выполнен:", meta['done_by']],
            [],
            ["ВАЖНО!", "Пользователь несет полную ответственность за verifikatsiyu проектных решений по системе обогрева."],
            [],
            ["ПРОЕКТНЫЕ ДАННЫЕ И ГРАНИЧНЫЕ УСЛОВИЯ В МАССИВЕ КОНСТРУКЦИИ", ""],
            ["Греющий кабель", meta['c_name_m']],
            ["Шаг укладки кабеля, мм", meta['spacing']],
            ["Глубина монтажа, мм", meta['depth']],
            ["Номинальное напряжение, В", meta['voltage']],
            ["Номинальная мощность кабеля, Вт/м", meta['P_line'] * meta['p_mult']],
            ["Расчетные условия наружной поверхности", meta['cond_txt']],
            ["Температура окружающей среды (воздух), °C", meta['T_amb']],
            ["Скорость ветра, км/ч", meta['wind']],
            ["Интенсивность снегопада, мм/ч", meta['snow']],
            ["Коэффициент теплоотдачи верхний, Вт/(м²·°C)", meta['alpha']],
            ["Температура грунта снизу, °C", meta['T_ground']],
            [],
            ["РЕЗУЛЬТАТЫ ТЕРМОДИНАМИЧЕСКОГО РАСЧЕТА", ""],
            ["Плотность теплового потока, вырабатываемого системой (Удельная мощность), Вт/м²", meta['flux_sys']],
            ["Минимальная температура наружной поверхности, °C", meta['min_temp']],
            ["Максимальная температура наружной поверхности, °C", meta['max_temp']],
            ["Температура оболочки кабеля, °C", meta['cable_temp']],
            ["Плотность верхнего теплового потока, Вт/м²", meta['flux_top']],
            ["Плотность нижнего теплового потока, Вт/м²", meta['flux_bottom']]
        ]
        df_sum = pd.DataFrame(summary_rows)
        df_sum.to_excel(writer, sheet_name="Общие результаты", index=False, header=False)
        
        headers = ["Глубина (мм)", "Относ. уровень", "Слой массива (материал)"] + [f"{x:.1f} мм" for x in st.session_state['x_coords_mm']]
        pd.DataFrame(st.session_state['matrix_rows'], columns=headers).to_excel(writer, sheet_name="Температурный профиль", index=False)
        
        workbook = writer.book
        ws1 = workbook["Общие результаты"]
        ws2 = workbook["Температурный профиль"]
        
        color_primary = "C23530" 
        fill_header = PatternFill(start_color=color_primary, end_color=color_primary, fill_type="solid")
        fill_warn = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
        
        font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        font_warn = Font(name="Arial", size=9, bold=True, color="B91C1C")
        border_thin = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'), top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))
        
        ws1.merge_cells("A1:B1")
        ws1["A1"].font = Font(name="Arial", size=14, bold=True, color=color_primary)
        
        ws1.merge_cells("A7:B7")
        ws1["A7"].fill = fill_warn
        ws1["A7"].font = font_warn
        
        header_rows = [9, 22]
        for r in header_rows:
            ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws1.cell(row=r, column=1).fill = fill_header
            ws1.cell(row=r, column=1).font = font_header
            
        for row in ws1.iter_rows(min_row=1, max_row=ws1.max_row, min_col=1, max_col=2):
            for cell in row:
                cell.border = border_thin
                if cell.font.name != "Arial" and cell.row != 1 and cell.row != 7 and cell.row not in header_rows:
                    cell.font = Font(name="Arial", size=10)
                
        for col_num in range(1, ws2.max_column + 1):
            cell = ws2.cell(row=1, column=col_num)
            cell.fill = PatternFill(start_color=color_primary, end_color=color_primary, fill_type="solid")
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            
        for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, min_col=1, max_col=ws2.max_column):
            for cell in row:
                cell.border = border_thin
                cell.font = Font(name="Arial", size=10)
                
        ws1.column_dimensions['A'].width = 52
        ws1.column_dimensions['B'].width = 32
        ws2.column_dimensions['A'].width = 15
        ws2.column_dimensions['B'].width = 15
        ws2.column_dimensions['C'].width = 25
        
    output.seek(0)
    return output
if st.session_state.get('calculated', False):
    st.success("Расчет успешно завершен!")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Мин. темп. поверхности", f"{st.session_state['min_temp']:.2f} °C")
    col2.metric("Макс. темп. поверхности", f"{st.session_state['max_temp']:.2f} °C")
    col3.metric("Температура кабеля", f"{st.session_state['cable_temp']:.2f} °C")
    
    plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'Tahoma', 'sans-serif']
    plt.rcParams['font.family'] = 'sans-serif'
    
    # --- Модель 1: Восстановленная расчетная модель ---
    fig1, ax1 = plt.subplots(figsize=(10, 4.8))
    im1 = ax1.imshow(st.session_state['T_array'], extent=[0, W*1000, H*1000, 0], cmap='jet', aspect='auto', interpolation='bicubic')
    cbar1 = fig1.colorbar(im1, ax=ax1)
    cbar1.set_label("Температура (°C)", fontsize=9)
    # Черная метка кабеля уменьшена ровно в 2 раза до размера 5
    ax1.plot(W*1000/2, depth, 'ko', markersize=5)
    ax1.axhline(h_carpet, color='white', linestyle='--')
    ax1.axhline(h_carpet + h_sand, color='white', linestyle='--')
    
    ax1.set_title(f"Тепловая модель массива конструкции (1 шаг укладки, {spacing}, мм)", fontsize=11, fontweight='bold', pad=10)
    ax1.set_xlabel("Положение расчетных узлов по оси X (мм)", fontsize=9)
    ax1.set_ylabel("Глубина массива конструкции, мм", fontsize=9)
    st.pyplot(fig1)
    
    # --- Модель 2: Восстановленная многопроходная масштабная модель ---
    fig2, ax2 = plt.subplots(figsize=(10, 4.8))
    T_single = st.session_state['T_array']
    T_multi_passes = np.hstack([T_single] * 5)
    W_multi = W * 5 * 1000
    
    im2 = ax2.imshow(T_multi_passes, extent=[0, W_multi, H*1000, 0], cmap='jet', aspect='auto', interpolation='bicubic')
    cbar2 = fig2.colorbar(im2, ax=ax2)
    cbar2.set_label("Температура (°C)", fontsize=9)
    
    # Черные метки кабелей уменьшены ровно в 2 раза до размера 4
    for i in range(5):
        cable_x_pos = (W * 1000 / 2) + i * (W * 1000)
        ax2.plot(cable_x_pos, depth, 'ko', markersize=4)
        
    ax2.axhline(h_carpet, color='white', linestyle='--')
    ax2.axhline(h_carpet + h_sand, color='white', linestyle='--')
    
    ax2.set_title(f"Тепловая модель массива конструкции (5 шагов укладки, {spacing*5}, мм)", fontsize=11, fontweight='bold', pad=10)
    ax2.set_xlabel("Положение расчетных узлов по оси X (мм)", fontsize=9)
    ax2.set_ylabel("Глубина массива конструкции, мм", fontsize=9)
    st.pyplot(fig2)
    
    st.subheader("Двухмерное температурное поле")
    right_col_name = f"{st.session_state['x_coords_mm'][-1]:.1f} мм"
    st.markdown(f"**Таблица распределения температур по толщине конструктивных слоев массива в зависимости от шага укладки кабеля (в сечениях от оси кабеля (0 мм) до ½ шага укладки ({right_col_name}))**")
    
    headers_st = ["Глубина (мм)", "Относ. уровень", "Слой массива (материал)"] + [f"{x:.1f} мм" for x in st.session_state['x_coords_mm']]
    st.dataframe(pd.DataFrame(st.session_state['matrix_rows'], columns=headers_st), use_container_width=True)
    
    img_buf1 = io.BytesIO()
    fig1.savefig(img_buf1, format='png', bbox_inches='tight', dpi=200)
    img_buf1.seek(0)
    plt.close(fig1)
    
    img_buf2 = io.BytesIO()
    fig2.savefig(img_buf2, format='png', bbox_inches='tight', dpi=200)
    img_buf2.seek(0)
    plt.close(fig2)
    
    r_data = {
        'p_name': project_name, 'c_name': customer_name, 'done_by': done_by,
        'c_name_m': cable_name, 'spacing': spacing, 'depth': depth, 'voltage': voltage,
        'p_mult': power_multiplier, 'T_amb': T_amb, 'wind': wind_speed_kmh, 'snow': snow_rate,
        'far': free_area_ratio, 'alpha': alpha_top, 'T_ground': T_ground, 'cond_txt': condition_text,
        'h_c': h_carpet, 'k_c': k_carpet, 'h_s': h_sand, 'k_s': k_sand, 'h_g': h_gravel, 'k_g': k_gravel,
        'min_temp': st.session_state['min_temp'], 'max_temp': st.session_state['max_temp'],
        'cable_temp': st.session_state['cable_temp'], 'P_line': P_line,
        'flux_top': st.session_state['flux_top'], 'flux_bottom': st.session_state['flux_bottom'],
        'flux_sys': st.session_state['flux_sys'],
        'm1': material_choice_1, 'm2': material_choice_2, 'm3': material_choice_3
    }
    
    pdf_file = generate_pdf_report(r_data, img_buf1, img_buf2)
    excel_file = generate_excel_report(r_data)
    
    st.divider()
    st.subheader("Выгрузка отчета")
    exp_col1, exp_col2 = st.columns(2)
    
    fn_suffix = f"{datetime.now().strftime('%d.%m.%Y_%H.%M')}_{project_name.replace(' ', '_')}"
    
    with exp_col1:
        st.download_button(label="Скачать отчет в формате PDF", data=pdf_file, file_name=f"NovaHeat_Report_{fn_suffix}.pdf", mime="application/pdf")
    with exp_col2:
        st.download_button(label="Скачать отчет в формате Excel", data=excel_file, file_name=f"NovaHeat_Report_{fn_suffix}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Введите исходные данные на боковой панели и нажмите кнопку 'Выполнить термодинамический расчет'.")
