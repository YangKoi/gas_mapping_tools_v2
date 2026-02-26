import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import matplotlib.colors as mcolors
import math
import io
import os
from PIL import Image

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from shapely.geometry import Point, LineString, Polygon
import shapely.affinity as affinity
from matplotlib.path import Path

# ==========================================
# CẤU HÌNH TRANG WEB & STATE
# ==========================================
st.set_page_config(page_title="Riken Viet - Enterprise Gas Mapping", layout="wide", initial_sidebar_state="expanded")

# --- ĐIỀU HƯỚNG BẰNG SIDEBAR (MENU TRÁI) ---
with st.sidebar:
    # Hiển thị Logo nếu có
    if os.path.exists("header_logo.png"):
        st.image("header_logo.png", use_container_width=True)
    else:
        st.markdown("### RIKEN VIET")
        
    st.header("🔄 BẢNG ĐIỀU KHIỂN")
    app_mode = st.radio("Chọn luồng công việc:", [
        "1️⃣ Thiết kế Không gian Đa lớp (3D)",
        "2️⃣ Rải nhanh trên Bản vẽ 2D (Overlay)"
    ])
    
    st.markdown("---")
    st.info("💡 **Mẹo:** Thu gọn thanh này bằng dấu 'X' hoặc phím tắt để mở rộng tối đa vùng thiết kế bản vẽ.")

st.title("🛡️ Riken Viet - Hệ thống Thiết kế & Dự toán Vùng phủ Khí")

# --- KHỞI TẠO DỮ LIỆU ---
# Dữ liệu Mode 1 (3D)
if 'room_data' not in st.session_state:
    st.session_state.room_data = pd.DataFrame({"X": [0, 15, 15, 0], "Y": [0, 0, 10, 10]}) 
if 'obs_data' not in st.session_state:
    st.session_state.obs_data = pd.DataFrame([{"Type": "Cylinder", "X": 7.5, "Y": 5.0, "Width_Radius": 1.5, "Length": 0.0, "Height": 4.0, "Angle": 0}])
if 'auto_config' not in st.session_state:
    st.session_state.auto_config = pd.DataFrame([
        {"Target Gas": "CH4", "Layer": "Khí Nhẹ (Sát trần)", "Model": "SD-1 (Catalytic)", "Radius": 5.0, "Color": "cyan"},
        {"Target Gas": "H2S", "Layer": "Khí Nặng (Sát sàn)", "Model": "GD-70D (Electro)", "Radius": 4.0, "Color": "magenta"}
    ])
if 'det_data' not in st.session_state:
    st.session_state.det_data = pd.DataFrame(columns=["ID", "Model", "Gas", "X", "Y", "Z", "Radius", "Color"])

# Dữ liệu Mode 2 (2D Overlay)
if 'det_data_2d' not in st.session_state:
    st.session_state.det_data_2d = pd.DataFrame(columns=["ID", "Model", "Gas", "X", "Y", "Radius", "Color"])
if 'obs_data_2d' not in st.session_state:
    st.session_state.obs_data_2d = pd.DataFrame([{"Type": "Box", "X": 15.0, "Y": 10.0, "Width_Radius": 5.0, "Length": 5.0, "Angle": 0}])
if 'auto_config_2d' not in st.session_state:
    st.session_state.auto_config_2d = pd.DataFrame([{"Target Gas": "CH4", "Model": "SD-1", "Radius": 5.0, "Color": "cyan"}])


# ==========================================
# CÁC HÀM XỬ LÝ LÕI VÀ ĐỒ HỌA
# ==========================================
def create_obstacle_polys(df_obs):
    obs_polys = []
    for _, row in df_obs.iterrows():
        if row['Type'] == 'Cylinder':
            obs_polys.append(Point(row['X'], row['Y']).buffer(row['Width_Radius']))
        elif row['Type'] == 'Box':
            w, l = row['Width_Radius'], row['Length']
            box = Polygon([(row['X']-w/2, row['Y']-l/2), (row['X']+w/2, row['Y']-l/2),
                           (row['X']+w/2, row['Y']+l/2), (row['X']-w/2, row['Y']+l/2)])
            box = affinity.rotate(box, row.get('Angle', 0), origin='center')
            obs_polys.append(box)
    return obs_polys

def check_collision_shapely(df_dets, obs_polys):
    collisions = []
    if not obs_polys or df_dets.empty: return collisions
    for _, det in df_dets.iterrows():
        pt = Point(det['X'], det['Y'])
        if any(poly.contains(pt) for poly in obs_polys): collisions.append(det['ID'])
    return collisions

def generate_2d_plot(room_poly, obs_polys, df_dets_group, gas_name, px, py, bg_img=None, b_w=0, b_h=0):
    minx, miny, maxx, maxy = room_poly.bounds
    res = 0.2
    xx, yy = np.meshgrid(np.arange(minx, maxx, res), np.arange(miny, maxy, res))
    pts_x, pts_y = xx.flatten(), yy.flatten()

    room_path = Path(list(room_poly.exterior.coords))
    in_room_mask = room_path.contains_points(np.c_[pts_x, pts_y])
    
    in_obs_mask = np.zeros(len(pts_x), dtype=bool)
    for obs in obs_polys:
        obs_path = Path(list(obs.exterior.coords))
        in_obs_mask |= obs_path.contains_points(np.c_[pts_x, pts_y])
    
    valid_points_mask = in_room_mask & ~in_obs_mask
    coverage_mask = np.zeros(len(pts_x), dtype=bool)

    for _, det in df_dets_group.iterrows():
        dx, dy, dr = det['X'], det['Y'], det['Radius']
        dist_sq = (pts_x - dx)**2 + (pts_y - dy)**2
        in_radius_mask = dist_sq <= dr**2
        
        check_mask = valid_points_mask & in_radius_mask
        det_pt = Point(dx, dy)
        
        for i in np.where(check_mask)[0]:
            if coverage_mask[i]: continue
            line = LineString([det_pt, Point(pts_x[i], pts_y[i])])
            if not any(line.crosses(obs) for obs in obs_polys): coverage_mask[i] = True

    diem_kha_dung = np.sum(valid_points_mask)
    ty_le = (np.sum(coverage_mask) / diem_kha_dung) * 100 if diem_kha_dung > 0 else 0

    fig, ax = plt.subplots(figsize=(10, 8))
    coverage_2d = coverage_mask.reshape(xx.shape)
    
    if bg_img is not None:
        ax.imshow(bg_img, extent=[0, b_w, 0, b_h])
        ax.contourf(xx, yy, coverage_2d, levels=[0.5, 1], colors=['#00FFAA'], alpha=0.4, zorder=1)
        rx, ry = room_poly.exterior.xy
        ax.plot(rx, ry, 'k-', lw=1, zorder=2)
    else:
        ax.contourf(xx, yy, coverage_2d, levels=[0.5, 1], colors=['#A8E6CF'], alpha=0.6, zorder=1)
        rx, ry = room_poly.exterior.xy
        ax.plot(rx, ry, 'k-', lw=3, zorder=2)
    
    ax.plot(px, py, 's', color='red', markersize=12, markeredgecolor='black', zorder=5)
    ax.text(px + 0.3, py + 0.3, "Control Panel", color='red', fontweight='bold', zorder=6, bbox=dict(facecolor='white', alpha=0.8, edgecolor='none', pad=1))

    for obs in obs_polys:
        ox, oy = obs.exterior.xy
        ax.fill(ox, oy, color='gray', alpha=0.8, zorder=4)

    valid_colors = mcolors.CSS4_COLORS
    for _, det in df_dets_group.iterrows():
        c = det['Color'].lower()
        use_c = c if c in valid_colors else 'blue'
        ax.add_patch(plt.Circle((det['X'], det['Y']), det['Radius'], color=use_c, fill=False, linestyle='--', lw=1.5, zorder=3))
        ax.text(det['X']+0.3, det['Y']+0.3, f"{det['ID']}", fontsize=8, color='black', zorder=6, bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', pad=1))
        ax.plot(det['X'], det['Y'], '^', color=use_c, markersize=12, markeredgecolor='black', zorder=5)

    ax.set_title(f"Bản đồ phân tích: {gas_name} | Mức an toàn: {ty_le:.1f}%", fontweight='bold')
    ax.axis('equal'); ax.grid(True, linestyle=':', alpha=0.5)
    return fig, ty_le

def generate_plotly_3d_complex(room_poly, rz, obs_polys, df_obs, df_dets, px, py, pz):
    fig = go.Figure()
    rx, ry = room_poly.exterior.xy
    rx, ry = list(rx), list(ry)
    
    fig.add_trace(go.Scatter3d(x=rx, y=ry, z=[0]*len(rx), mode='lines', line=dict(color='white', width=4), name='Đáy tường'))
    fig.add_trace(go.Scatter3d(x=rx, y=ry, z=[rz]*len(rx), mode='lines', line=dict(color='white', width=4), name='Đỉnh tường'))
    for x, y in zip(rx[:-1], ry[:-1]):
        fig.add_trace(go.Scatter3d(x=[x,x], y=[y,y], z=[0,rz], mode='lines', line=dict(color='white', width=2), showlegend=False))

    fig.add_trace(go.Scatter3d(x=[px], y=[py], z=[pz], mode='markers+text', marker=dict(symbol='square', size=8, color='red'), text=["TỦ TRUNG TÂM"], textposition="top center", textfont=dict(color="red", size=12, weight="bold"), name="Control Panel"))

    def get_sphere(x0, y0, z0, r):
        u, v = np.mgrid[0:2*np.pi:15j, 0:np.pi:10j]
        return r*np.cos(u)*np.sin(v)+x0, r*np.sin(u)*np.sin(v)+y0, r*np.cos(v)+z0

    for _, det in df_dets.iterrows():
        hover_label = f"Model: {det['ID']}<br>Mục tiêu: {det['Gas']}<br>Cao độ Z: {det['Z']}m"
        fig.add_trace(go.Scatter3d(x=[det['X']], y=[det['Y']], z=[det['Z']], mode='markers+text', marker=dict(size=6, color='white'), text=[f"{det['ID']}<br>({det['Gas']})"], textposition="top center", textfont=dict(color="white", size=10), name=det['ID'], hoverinfo="text", hovertext=hover_label))
        
        sx, sy, sz = get_sphere(det['X'], det['Y'], det['Z'], det['Radius'])
        fig.add_trace(go.Surface(x=sx, y=sy, z=sz, opacity=0.15, showscale=False, colorscale=[[0, det['Color']], [1, det['Color']]]))

    for i, (_, obs) in enumerate(df_obs.iterrows()):
        if obs['Type'] == 'Cylinder':
            z_grid, theta = np.mgrid[0:obs['Height']:2j, 0:2*np.pi:20j]
            fig.add_trace(go.Surface(x=obs['Width_Radius']*np.cos(theta)+obs['X'], y=obs['Width_Radius']*np.sin(theta)+obs['Y'], z=z_grid, opacity=1.0, showscale=False, colorscale='Greys', name="Bồn trụ"))
        elif obs['Type'] == 'Box':
            box_2d = obs_polys[i]
            bx, by = box_2d.exterior.xy
            bx, by = list(bx)[:-1], list(by)[:-1] 
            x_box, y_box = bx * 2, by * 2
            z_box = [0]*4 + [obs['Height']]*4
            ii, jj, kk = [7,0,0,0,4,4,6,6,4,0,3,2], [3,4,1,2,5,6,5,2,0,1,6,3], [0,7,2,3,6,7,1,1,5,5,7,6]
            fig.add_trace(go.Mesh3d(x=x_box, y=y_box, z=z_box, i=ii, j=jj, k=kk, color='gray', opacity=1.0, name="Tủ/Kệ"))

    minx, miny, maxx, maxy = room_poly.bounds
    fig.update_layout(
        scene=dict(xaxis=dict(range=[minx, maxx], title='X', backgroundcolor="rgb(30,30,30)", gridcolor="gray"),
                   yaxis=dict(range=[miny, maxy], title='Y', backgroundcolor="rgb(30,30,30)", gridcolor="gray"),
                   zaxis=dict(range=[0, max(rz, 5)], title='Z', backgroundcolor="rgb(30,30,30)", gridcolor="gray"),
                   aspectmode='data'),
        paper_bgcolor="rgb(15,15,15)", plot_bgcolor="rgb(15,15,15)", margin=dict(l=0, r=0, b=0, t=30), showlegend=False
    )
    return fig

def build_bom_df(det_df, p_x, p_y, r_z, p_z, waste):
    bom_items = [{"STT": 1, "Hạng mục thiết bị": "Tủ điều khiển trung tâm đo khí", "Đơn vị": "Bộ", "Khối lượng": 1}]
    stt = 2
    total_cable_length = 0
    if not det_df.empty and 'Model' in det_df.columns:
        counts = det_df['Model'].value_counts()
        for model, qty in counts.items():
            bom_items.append({"STT": stt, "Hạng mục thiết bị": f"Đầu dò đo khí rò rỉ - Model: {model}", "Đơn vị": "Bộ", "Khối lượng": qty})
            stt += 1
        for _, d in det_df.iterrows():
            cable_up = abs(r_z - p_z) if 'Z' in det_df.columns else 3.0
            cable_horizontal = abs(d['X'] - p_x) + abs(d['Y'] - p_y)
            cable_down = abs(r_z - d['Z']) if 'Z' in det_df.columns else 1.5 
            total_cable_length += (cable_up + cable_horizontal + cable_down)
        total_cable_length = math.ceil(total_cable_length * (1 + waste / 100))
    bom_items.append({"STT": stt, "Hạng mục thiết bị": "Cáp tín hiệu chống nhiễu chuyên dụng", "Đơn vị": "Mét", "Khối lượng": total_cable_length})
    bom_items.append({"STT": stt+1, "Hạng mục thiết bị": "Chuông đèn cảnh báo (Siren/Light)", "Đơn vị": "Bộ", "Khối lượng": len(det_df) if not det_df.empty else 1})
    return pd.DataFrame(bom_items)

def generate_full_word_report(figs_dict, img_3d_bytes, bom_df, p_name, c_name, author, r_date, r_num, is_3d=True):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)

    section = doc.sections[0]
    header = section.header
    h_para = header.paragraphs[0]
    h_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists("header_logo.png"): h_para.add_run().add_picture("header_logo.png", width=Inches(6.5))
    else: h_para.add_run("CÔNG TY TNHH CÔNG NGHỆ THIẾT BỊ DÒ KHÍ RIKEN VIET").bold = True
    
    footer = section.footer
    f_para = footer.paragraphs[0]
    f_para.text = "Bản quyền thuộc về Riken Việt, không sao chép và sử dụng sai mục đích."
    f_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in f_para.runs: run.font.name = 'Arial'; run.font.size = Pt(9); run.font.italic = True; run.font.color.rgb = RGBColor(128, 128, 128)

    doc.add_paragraph()
    title = doc.add_heading('BÁO CÁO THIẾT KẾ VÀ DỰ TOÁN HỆ THỐNG ĐO KHÍ', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs: run.font.name = 'Arial'
    
    p_num = doc.add_paragraph(f"Số/ No.: {r_num}")
    p_num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in p_num.runs: run.font.name = 'Arial'; run.font.italic = True
        
    doc.add_paragraph(f"Tên Dự án: {p_name}").runs[0].font.bold = True
    doc.add_paragraph(f"Đơn vị / Khách hàng: {c_name}").runs[0].font.bold = True
    doc.add_paragraph(f"Người lập báo cáo: {author}")
    doc.add_paragraph(f"Ngày xuất báo cáo: {r_date.strftime('%d/%m/%Y')}")
    doc.add_paragraph("_" * 60)
    
    section_num = 1
    if is_3d and img_3d_bytes:
        h1 = doc.add_heading(f'{section_num}. PHÂN BỔ KHÔNG GIAN TỔNG THỂ (MÔ PHỎNG 3D)', level=1)
        for run in h1.runs: run.font.name = 'Arial'
        doc.add_paragraph("Sơ đồ dưới đây thể hiện vị trí không gian 3 chiều của các thiết bị đo khí, tủ trung tâm và các vật cản thực tế tại hiện trường.")
        p = doc.add_paragraph()
        p.add_run().add_picture(img_3d_bytes, width=Inches(6.0))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        section_num += 1
        
    h2 = doc.add_heading(f'{section_num}. PHÂN TÍCH ĐIỂM MÙ (MẶT BẰNG 2D)', level=1)
    for run in h2.runs: run.font.name = 'Arial'
    for gas_name, fig_info in figs_dict.items():
        p_title = doc.add_paragraph()
        p_title.add_run(f"Bản đồ phân hệ: {gas_name} ").bold = True
        p_title.add_run(f"(Mức độ an toàn: {fig_info['coverage']:.1f}%)").italic = True
        img_stream = io.BytesIO()
        fig_info['fig'].savefig(img_stream, format='png', bbox_inches='tight', dpi=150)
        img_stream.seek(0)
        p_img = doc.add_paragraph()
        p_img.add_run().add_picture(img_stream, width=Inches(6.0))
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    section_num += 1
        
    h3 = doc.add_heading(f'{section_num}. BẢNG BÓC TÁCH KHỐI LƯỢNG VẬT TƯ (BOM)', level=1)
    for run in h3.runs: run.font.name = 'Arial'
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    headers = ['STT', 'Hạng mục Thiết bị & Vật tư', 'Đơn vị', 'Khối lượng']
    for i in range(4):
        hdr[i].text = headers[i]
        for run in hdr[i].paragraphs[0].runs: run.font.name = 'Arial'; run.font.bold = True
    for _, row in bom_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text, row_cells[1].text, row_cells[2].text, row_cells[3].text = str(row['STT']), str(row['Hạng mục thiết bị']), str(row['Đơn vị']), f"{int(row['Khối lượng']):,}"
        for cell in row_cells:
            for p in cell.paragraphs:
                for run in p.runs: run.font.name = 'Arial'
        
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream


# ==========================================
# LUỒNG CHÍNH 1: MÔ PHỎNG 3D CHUYÊN SÂU
# ==========================================
if app_mode == "1️⃣ Thiết kế Không gian Đa lớp (3D)":
    st.markdown("## 🚀 Chế độ Thiết kế Không gian 3D Chuyên sâu")
    
    col_input1, col_input2 = st.columns([1.2, 1.1])
    with col_input1:
        st.header("1. Không gian & Tủ Trung Tâm")
        room_z = st.number_input("Chiều cao trần nhà xưởng (Z) - mét", min_value=1.0, value=5.0, key="z_3d")

        st.subheader("🎛️ Vị trí Tủ Điều Khiển & Hao hụt")
        col_p1, col_p2, col_p3 = st.columns(3)
        panel_x = col_p1.number_input("Tọa độ X (Tủ)", value=0.0, key="px_3d")
        panel_y = col_p2.number_input("Tọa độ Y (Tủ)", value=0.0, key="py_3d")
        panel_z = col_p3.number_input("Cao độ Z (Tủ)", value=1.5, key="pz_3d")
        wastage_percent = st.number_input("Hệ số hao hụt cáp thi công (%)", min_value=0, value=20, step=5, key="wast_3d")

        st.subheader("📐 Định hình Không gian")
        col_t1, col_t2, col_t3 = st.columns(3)
        with col_t1:
            if st.button("🟩 Mẫu Chữ Nhật", use_container_width=True, key="btn_rect"): st.session_state.room_data = pd.DataFrame({"X": [0, 15, 15, 0], "Y": [0, 0, 10, 10]}); st.rerun()
        with col_t2:
            if st.button("▛ Mẫu Chữ L", use_container_width=True, key="btn_l"): st.session_state.room_data = pd.DataFrame({"X": [0, 15, 15, 6, 6, 0], "Y": [0, 0, 6, 6, 12, 12]}); st.rerun()
        with col_t3:
            if st.button("⨆ Mẫu Chữ U", use_container_width=True, key="btn_u"): st.session_state.room_data = pd.DataFrame({"X": [0, 20, 20, 15, 15, 5, 5, 0], "Y": [0, 0, 15, 15, 5, 5, 15, 15]}); st.rerun()

        edited_room = st.data_editor(st.session_state.room_data, num_rows="dynamic", use_container_width=True, key="ed_room_3d")
        
        if len(edited_room) >= 3:
            room_poly = Polygon(list(zip(edited_room['X'], edited_room['Y'])))
            fig_grid, ax_grid = plt.subplots(figsize=(6, 5))
            x_ext, y_ext = room_poly.exterior.xy
            ax_grid.plot(x_ext, y_ext, color='#333333', linewidth=2)
            ax_grid.fill(x_ext, y_ext, alpha=0.1, color='blue')
            for idx, row in edited_room.iterrows():
                ax_grid.plot(row['X'], row['Y'], 'ro', markersize=6, zorder=10)
                ax_grid.text(row['X'] + 0.3, row['Y'] + 0.3, f"P{idx}", color='red', fontweight='bold', fontsize=11, zorder=11)
            ax_grid.plot(panel_x, panel_y, 's', color='red', markersize=12, markeredgecolor='black', zorder=12)
            ax_grid.text(panel_x + 0.5, panel_y + 0.5, "TỦ TT", color='red', fontweight='bold', zorder=12)
            
            # Vẽ thử vật cản nháp
            for _, obs in st.session_state.obs_data.iterrows():
                if obs['Type'] == 'Cylinder': ax_grid.add_patch(plt.Circle((obs['X'], obs['Y']), obs['Width_Radius'], color='gray', alpha=0.5))
                elif obs['Type'] == 'Box':
                    box = Polygon([(obs['X']-obs['Width_Radius']/2, obs['Y']-obs['Length']/2), (obs['X']+obs['Width_Radius']/2, obs['Y']-obs['Length']/2),
                                   (obs['X']+obs['Width_Radius']/2, obs['Y']+obs['Length']/2), (obs['X']-obs['Width_Radius']/2, obs['Y']+obs['Length']/2)])
                    bx, by = affinity.rotate(box, obs.get('Angle', 0), origin='center').exterior.xy
                    ax_grid.fill(bx, by, color='gray', alpha=0.5)
            
            for _, det in st.session_state.det_data.iterrows():
                c = det['Color'].lower() if det['Color'].lower() in mcolors.CSS4_COLORS else 'blue'
                ax_grid.plot(det['X'], det['Y'], '^', color=c, markersize=8, markeredgecolor='black')

            ax_grid.set_aspect('equal'); ax_grid.grid(True, linestyle='--', alpha=0.5)
            st.pyplot(fig_grid)
        else:
            st.error("Phòng cần ít nhất 3 góc!")
            room_poly = None

    with col_input2:
        st.header("2. Bố trí Thiết bị & Khí (3D)")
        with st.expander("🚧 Danh sách Vật cản (Cylinder / Box)", expanded=True):
            edited_obs = st.data_editor(st.session_state.obs_data, num_rows="dynamic", use_container_width=True, column_config={"Type": st.column_config.SelectboxColumn("Loại", options=["Cylinder", "Box"])}, key="ed_obs_3d")
            st.session_state.obs_data = edited_obs

        with st.expander("⚙️ Cấu hình Các Phân hệ Khí", expanded=True):
            edited_auto_config = st.data_editor(st.session_state.auto_config, num_rows="dynamic", use_container_width=True,
                column_config={"Layer": st.column_config.SelectboxColumn("Mặt phẳng", options=["Khí Nhẹ (Sát trần)", "Khí Nặng (Sát sàn)"]), "Color": st.column_config.SelectboxColumn("Màu", options=["cyan", "magenta", "yellow", "lime", "red"])}, key="ed_cfg_3d")
            st.session_state.auto_config = edited_auto_config

            if st.button("🚀 Tự động Rải Đầu dò (Mặt bằng 3D)", type="primary"):
                if room_poly is not None:
                    new_dets = []
                    for _, row_cfg in edited_auto_config.iterrows():
                        z_val = max(room_z - 0.5, 0.5) if "Nhẹ" in row_cfg["Layer"] else 0.5
                        spacing = row_cfg["Radius"] * 1.5 
                        minx, miny, maxx, maxy = room_poly.bounds
                        nx, ny = max(1, math.ceil((maxx - minx)/spacing)), max(1, math.ceil((maxy - miny)/spacing))
                        count = 1
                        for x in np.linspace(minx + (maxx-minx)/(2*nx), maxx - (maxx-minx)/(2*nx), nx):
                            for y in np.linspace(miny + (maxy-miny)/(2*ny), maxy - (maxy-miny)/(2*ny), ny):
                                if room_poly.contains(Point(x, y)): 
                                    new_dets.append({"ID": f"{row_cfg['Model']} ({count:02d})", "Model": row_cfg['Model'], "Gas": f"{row_cfg['Target Gas']}", "X": round(x, 1), "Y": round(y, 1), "Z": z_val, "Radius": row_cfg["Radius"], "Color": row_cfg["Color"]})
                                    count += 1
                    st.session_state.det_data = pd.DataFrame(new_dets)
                    st.rerun()

        st.write("📋 **Bảng Tọa độ Đầu dò Thực tế:**")
        edited_dets = st.data_editor(st.session_state.det_data, num_rows="dynamic", use_container_width=True, key="ed_dets_3d")
        st.session_state.det_data = edited_dets

    # CHẠY PHÂN TÍCH VÀ HIỂN THỊ ĐỒ HỌA TRỰC TIẾP LÊN WEB (ĐÃ KHÔI PHỤC)
    st.markdown("---")
    st.header("3. 📊 Phân tích & Xuất Báo cáo")
    col_info1, col_info2, col_info3 = st.columns([1, 1, 1])
    with col_info1: project_name = st.text_input("Tên Dự án", value="Thiết kế Giám sát Rò rỉ", key="pn_1")
    with col_info2: client_name = st.text_input("Khách hàng", value="Nhà máy ABC", key="cn_1")
    with col_info3: report_number = st.text_input("Số Báo cáo", value="RKV_TE_001/BC", key="rn_1")
    
    author_name = st.text_input("Người lập báo cáo", value="Cao Minh Lợi - Giám đốc Kỹ thuật", key="au_1")
    report_date = st.date_input("Ngày lập báo cáo", key="rd_1")

    if st.button("🚀 CHẠY MÔ PHỎNG & TẠO FILE BÁO CÁO (3D)", type='primary', use_container_width=True):
        if room_poly is None or edited_dets.empty:
            st.error("Lỗi: Cần vẽ phòng và rải đầu dò trước khi chạy!")
        else:
            with st.spinner('Đang nội suy không gian 3D và kết xuất báo cáo...'):
                obs_polys = create_obstacle_polys(edited_obs)
                
                # Hiển thị biểu đồ 3D
                fig_3d = generate_plotly_3d_complex(room_poly, room_z, obs_polys, edited_obs, edited_dets, panel_x, panel_y, panel_z)
                st.plotly_chart(fig_3d, use_container_width=True)
                
                img_3d_bytes = io.BytesIO()
                try: fig_3d.write_image(img_3d_bytes, format='png', width=800, height=500); img_3d_bytes.seek(0)
                except: img_3d_bytes = None

                # ĐÃ KHÔI PHỤC: HIỂN THỊ TABS BẢN ĐỒ 2D NGAY TRÊN WEB
                gas_groups = edited_dets['Gas'].unique()
                ui_tabs = st.tabs([f"Bản đồ: {g}" for g in gas_groups])
                figs_dict = {} 
                
                for i, gas_name in enumerate(gas_groups):
                    with ui_tabs[i]:
                        f, c = generate_2d_plot(room_poly, obs_polys, edited_dets[edited_dets['Gas']==gas_name], gas_name, panel_x, panel_y)
                        st.pyplot(f) # Hiển thị hình ảnh
                        if c >= 80: st.success(f"✅ Độ phủ: {c:.1f}%")
                        else: st.warning(f"⚠️ Độ phủ: {c:.1f}%")
                        figs_dict[gas_name] = {'fig': f, 'coverage': c}

                # Tạo BOM và xuất Word
                bom_df = build_bom_df(edited_dets, panel_x, panel_y, room_z, panel_z, wastage_percent)
                st.write("📋 **Bảng Bóc tách Khối lượng (BOM):**")
                st.dataframe(bom_df) # Hiển thị BOM lên web
                
                word_stream = generate_full_word_report(figs_dict, img_3d_bytes, bom_df, project_name, client_name, author_name, report_date, report_number, is_3d=True)
                st.download_button("📥 TẢI BÁO CÁO WORD (3D)", word_stream, f"BaoCao_3D_{client_name}.docx", type="primary")


# ==========================================
# LUỒNG CHÍNH 2: RẢI NHANH TRÊN BẢN VẼ 2D
# ==========================================
elif app_mode == "2️⃣ Rải nhanh trên Bản vẽ 2D (Overlay)":
    st.markdown("## 🖼️ Chế độ Rải nhanh trên Bản vẽ 2D (Overlay)")
    
    col_2d_1, col_2d_2 = st.columns([1.2, 1.1])
    
    with col_2d_1:
        st.header("1. Upload Bản vẽ & Căn Tỷ lệ")
        bg_file = st.file_uploader("Tải lên mặt bằng (PNG/JPG)", type=['png', 'jpg', 'jpeg'])
        col_w, col_h = st.columns(2)
        bg_real_width = col_w.number_input("Chiều ngang thực tế (m)", value=30.0)
        bg_real_height = col_h.number_input("Chiều dọc thực tế (m)", value=20.0)
        
        st.subheader("🎛️ Tủ Điều Khiển & Hao hụt cáp")
        col_p2d1, col_p2d2 = st.columns(2)
        panel_x_2d = col_p2d1.number_input("Tọa độ X (Tủ)", value=0.0, key="px_2d")
        panel_y_2d = col_p2d2.number_input("Tọa độ Y (Tủ)", value=0.0, key="py_2d")
        wastage_percent_2d = st.number_input("Hệ số hao hụt cáp (%)", min_value=0, value=20, step=5, key="wast_2d")
        
    with col_2d_2:
        st.header("2. Cấu hình Thiết bị & Vật cản")
        with st.expander("⚙️ Danh sách Thiết bị", expanded=True):
            edited_auto_config_2d = st.data_editor(st.session_state.auto_config_2d, num_rows="dynamic", use_container_width=True, 
                column_config={"Color": st.column_config.SelectboxColumn("Màu", options=["cyan", "magenta", "yellow", "lime", "red"])}, key="ed_cfg_2d")
            st.session_state.auto_config_2d = edited_auto_config_2d
            
            if st.button("🚀 Tạo lưới phủ tự động", type="primary"):
                new_dets_2d = []
                for _, row_cfg in edited_auto_config_2d.iterrows():
                    spacing = row_cfg["Radius"] * 1.5 
                    nx, ny = max(1, math.ceil(bg_real_width / spacing)), max(1, math.ceil(bg_real_height / spacing))
                    count = 1
                    for x in np.linspace(spacing/2, bg_real_width - spacing/2, nx):
                        for y in np.linspace(spacing/2, bg_real_height - spacing/2, ny):
                            new_dets_2d.append({"ID": f"{row_cfg['Model']} ({count:02d})", "Model": row_cfg['Model'], "Gas": row_cfg['Target Gas'], "X": round(x, 1), "Y": round(y, 1), "Radius": row_cfg["Radius"], "Color": row_cfg["Color"]})
                            count += 1
                st.session_state.det_data_2d = pd.DataFrame(new_dets_2d)
                st.rerun()

        st.write("📋 **Tọa độ Thiết bị (Chỉnh tay nếu đè lên vật cản):**")
        edited_dets_2d = st.data_editor(st.session_state.det_data_2d, num_rows="dynamic", use_container_width=True, key="ed_dets_2d")
        st.session_state.det_data_2d = edited_dets_2d
        
        with st.expander("🚧 Vẽ Vật cản (Chặn bán kính trên ảnh)"):
            edited_obs_2d = st.data_editor(st.session_state.obs_data_2d, num_rows="dynamic", use_container_width=True, column_config={"Type": st.column_config.SelectboxColumn("Loại", options=["Cylinder", "Box"])}, key="ed_obs_2d")
            st.session_state.obs_data_2d = edited_obs_2d

    st.markdown("---")
    st.header("3. 📊 Phân tích & Xuất Báo cáo")
    col_info1, col_info2, col_info3 = st.columns([1, 1, 1])
    with col_info1: project_name = st.text_input("Tên Dự án", value="Thiết kế Giám sát Rò rỉ", key="pn_2")
    with col_info2: client_name = st.text_input("Khách hàng", value="Nhà máy ABC", key="cn_2")
    with col_info3: report_number = st.text_input("Số Báo cáo", value="RKV_TE_001/BC", key="rn_2")
    
    author_name = st.text_input("Người lập báo cáo", value="Nguyễn Đình Trường Giang", key="au_2")
    report_date = st.date_input("Ngày lập báo cáo", key="rd_2")

    if st.button("🚀 CHẠY MÔ PHỎNG & TẠO FILE BÁO CÁO (2D)", type='primary', use_container_width=True):
        if bg_file is None or edited_dets_2d.empty:
            st.error("Lỗi: Bạn cần Upload bản vẽ và rải đầu dò!")
        else:
            with st.spinner('Đang ép lớp nhiệt (Heatmap) lên bản vẽ...'):
                img = Image.open(bg_file)
                bg_poly = Polygon([(0,0), (bg_real_width,0), (bg_real_width, bg_real_height), (0, bg_real_height)])
                obs_polys_2d = create_obstacle_polys(edited_obs_2d)
                
                gas_groups = edited_dets_2d['Gas'].unique()
                ui_tabs = st.tabs([f"Bản đồ: {g}" for g in gas_groups])
                figs_dict_2d = {}
                
                for i, g in enumerate(gas_groups):
                    with ui_tabs[i]:
                        f, c = generate_2d_plot(bg_poly, obs_polys_2d, edited_dets_2d[edited_dets_2d['Gas']==g], g, panel_x_2d, panel_y_2d, bg_img=img, b_w=bg_real_width, b_h=bg_real_height)
                        st.pyplot(f)
                        if c >= 80: st.success(f"✅ Độ phủ: {c:.1f}%")
                        else: st.warning(f"⚠️ Độ phủ: {c:.1f}%")
                        figs_dict_2d[g] = {'fig': f, 'coverage': c}
                
                bom_df_2d = build_bom_df(edited_dets_2d, panel_x_2d, panel_y_2d, 3.0, 1.5, wastage_percent_2d)
                st.write("📋 **Bảng Bóc tách Khối lượng (BOM):**")
                st.dataframe(bom_df_2d)
                
                word_stream_2d = generate_full_word_report(figs_dict_2d, None, bom_df_2d, project_name, client_name, author_name, report_date, report_number, is_3d=False)
                st.download_button("📥 TẢI BÁO CÁO WORD (2D Overlay)", word_stream_2d, f"BaoCao_2D_{client_name}.docx", type="primary")

# ==========================================
# FOOTER BẢN QUYỀN
# ==========================================
st.markdown("""
    <hr style="border: 0; height: 1px; background-image: linear-gradient(to right, rgba(255, 255, 255, 0), rgba(255, 255, 255, 0.2), rgba(255, 255, 255, 0)); margin-top: 50px;">
    <div style="text-align: center; color: #888888; font-size: 14px; padding-bottom: 20px;">
        &copy; 2026 All Rights Reserved.<br>
        Designed and programmed by <b>trggiang</b>.
    </div>
""", unsafe_allow_html=True)
