import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import matplotlib.colors as mcolors
import math
import io
import os

from docx import Document
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from shapely.geometry import Point, LineString, Polygon
import shapely.affinity as affinity
from matplotlib.path import Path

# ==========================================
# CẤU HÌNH TRANG WEB & STATE
# ==========================================
st.set_page_config(page_title="Riken Viet - Enterprise Gas Mapping", layout="wide")

# --- THANH CÔNG CỤ CỨU HỘ MẪU WORD (SIDEBAR) ---
def create_clean_template():
    doc = Document()
    
    # Header
    section = doc.sections[0]
    header = section.header
    header.paragraphs[0].text = "CÔNG TY TNHH CÔNG NGHỆ THIẾT BỊ DÒ KHÍ RIKEN VIET\nSố/ No.: {{ report_number }}"
    
    doc.add_heading('BÁO CÁO THIẾT KẾ VÀ DỰ TOÁN HỆ THỐNG ĐO KHÍ', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Tên Dự án: {{ project_name }}")
    doc.add_paragraph("Đơn vị / Khách hàng: {{ client_name }}")
    doc.add_paragraph("Người lập báo cáo: {{ author_name }}")
    doc.add_paragraph("Ngày xuất báo cáo: {{ report_date }}")
    
    # 3D
    doc.add_heading('1. PHÂN BỔ KHÔNG GIAN TỔNG THỂ', level=1)
    doc.add_paragraph("Sơ đồ dưới đây thể hiện vị trí không gian 3 chiều của các thiết bị đo khí, tủ trung tâm và các vật cản thực tế.")
    p3d = doc.add_paragraph("{{ img_3d }}")
    p3d.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 2D
    doc.add_heading('2. PHÂN TÍCH ĐIỂM MÙ THEO LỚP KHÍ (MẶT BẰNG 2D)', level=1)
    doc.add_paragraph("{%p for map in gas_maps %}")
    doc.add_paragraph("Bản đồ phân hệ: {{ map.gas_name }} (Mức độ che phủ an toàn: {{ map.coverage }}%)")
    p2d = doc.add_paragraph("{{ map.img_2d }}")
    p2d.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("{%p endfor %}")
    
    # BOM Table
    doc.add_heading('3. BẢNG BÓC TÁCH KHỐI LƯỢNG VẬT TƯ (BOM)', level=1)
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = 'STT'
    hdr[1].text = 'Hạng mục Thiết bị & Vật tư'
    hdr[2].text = 'Đơn vị'
    hdr[3].text = 'Khối lượng'
    
    row = table.rows[1].cells
    row[0].text = "{%tr for item in bom_items %}{{ item.stt }}"
    row[1].text = "{{ item.name }}"
    row[2].text = "{{ item.unit }}"
    row[3].text = "{{ item.qty }}{%tr endfor %}"
    
    # Footer
    footer = section.footer
    footer.paragraphs[0].text = "Bản quyền thuộc về Riken Việt, không sao chép và sử dụng sai mục đích."
    
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

with st.sidebar:
    st.header("🛠️ Công cụ sửa lỗi File Word")
    st.info("Nếu báo cáo bị lỗi 'unknown tag', nguyên nhân do MS Word tự chèn mã ẩn làm gãy từ khóa. Hãy bấm nút dưới đây để tải File Mẫu Chuẩn (100% không dính lỗi).")
    clean_docx = create_clean_template()
    st.download_button("📥 TẢI FILE MẪU CHUẨN (Sạch lỗi XML)", clean_docx, "Mau_Bao_Cao.docx", type="primary")
    st.caption("**Hướng dẫn:** Tải file này về -> Mở lên chèn Logo của công ty vào -> Lưu lại -> Tải đè lên GitHub là xong!")


# --- TIẾP TỤC GIAO DIỆN CHÍNH ---
st.title("🛡️ Riken Viet - Hệ thống Thiết kế & Dự toán Vùng phủ Khí")

if 'room_data' not in st.session_state:
    st.session_state.room_data = pd.DataFrame({"X": [0, 15, 15, 0], "Y": [0, 0, 10, 10]}) 
if 'obs_data' not in st.session_state:
    st.session_state.obs_data = pd.DataFrame([
        {"Type": "Cylinder", "X": 7.5, "Y": 5.0, "Width_Radius": 1.5, "Length": 0.0, "Height": 4.0, "Angle": 0}
    ])
if 'auto_config' not in st.session_state:
    st.session_state.auto_config = pd.DataFrame([
        {"Target Gas": "CH4", "Layer": "Khí Nhẹ (Sát trần)", "Model": "SD-1 (Catalytic)", "Radius": 5.0, "Color": "cyan"},
        {"Target Gas": "H2S", "Layer": "Khí Nặng (Sát sàn)", "Model": "GD-70D (Electro)", "Radius": 4.0, "Color": "magenta"}
    ])
if 'det_data' not in st.session_state:
    st.session_state.det_data = pd.DataFrame(columns=["ID", "Model", "Gas", "X", "Y", "Z", "Radius", "Color"])

# ==========================================
# 1. GIAO DIỆN NHẬP LIỆU
# ==========================================
col_input1, col_input2 = st.columns([1.2, 1.1])

with col_input1:
    st.header("1. Không gian & Tủ Trung Tâm")
    room_z = st.number_input("Chiều cao trần nhà xưởng (Z) - mét", min_value=1.0, value=5.0)

    st.subheader("🎛️ Vị trí Tủ Điều Khiển & Hao hụt")
    col_p1, col_p2, col_p3 = st.columns(3)
    panel_x = col_p1.number_input("Tọa độ X (Tủ)", value=0.0)
    panel_y = col_p2.number_input("Tọa độ Y (Tủ)", value=0.0)
    panel_z = col_p3.number_input("Cao độ Z (Tủ)", value=1.5)
    wastage_percent = st.number_input("Hệ số hao hụt cáp thi công (%)", min_value=0, value=20, step=5)

    st.subheader("📐 Định hình Không gian")
    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        if st.button("🟩 Mẫu Chữ Nhật", use_container_width=True):
            st.session_state.room_data = pd.DataFrame({"X": [0, 15, 15, 0], "Y": [0, 0, 10, 10]})
            st.rerun()
    with col_t2:
        if st.button("▛ Mẫu Chữ L", use_container_width=True):
            st.session_state.room_data = pd.DataFrame({"X": [0, 15, 15, 6, 6, 0], "Y": [0, 0, 6, 6, 12, 12]})
            st.rerun()
    with col_t3:
        if st.button("⨆ Mẫu Chữ U", use_container_width=True):
            st.session_state.room_data = pd.DataFrame({"X": [0, 20, 20, 15, 15, 5, 5, 0], "Y": [0, 0, 15, 15, 5, 5, 15, 15]})
            st.rerun()

    edited_room = st.data_editor(st.session_state.room_data, num_rows="dynamic", use_container_width=True)
    
    if len(edited_room) >= 3:
        room_coords = list(zip(edited_room['X'], edited_room['Y']))
        room_poly = Polygon(room_coords)
        
        fig_grid, ax_grid = plt.subplots(figsize=(6, 5))
        x_ext, y_ext = room_poly.exterior.xy
        ax_grid.plot(x_ext, y_ext, color='#333333', linewidth=2)
        ax_grid.fill(x_ext, y_ext, alpha=0.1, color='blue')
        
        for idx, row in edited_room.iterrows():
            ax_grid.plot(row['X'], row['Y'], 'ro', markersize=6, zorder=10)
            ax_grid.text(row['X'] + 0.3, row['Y'] + 0.3, f"P{idx}", color='red', fontweight='bold', fontsize=11, zorder=11)
        
        ax_grid.plot(panel_x, panel_y, 's', color='red', markersize=12, markeredgecolor='black', zorder=12)
        ax_grid.text(panel_x + 0.5, panel_y + 0.5, "TỦ TT", color='red', fontweight='bold', zorder=12)

        for _, obs in st.session_state.obs_data.iterrows():
            if obs['Type'] == 'Cylinder':
                c = plt.Circle((obs['X'], obs['Y']), obs['Width_Radius'], color='gray', alpha=0.5)
                ax_grid.add_patch(c)
            elif obs['Type'] == 'Box':
                w, l = obs['Width_Radius'], obs['Length']
                box = Polygon([(obs['X']-w/2, obs['Y']-l/2), (obs['X']+w/2, obs['Y']-l/2),
                               (obs['X']+w/2, obs['Y']+l/2), (obs['X']-w/2, obs['Y']+l/2)])
                box = affinity.rotate(box, obs.get('Angle', 0), origin='center')
                bx, by = box.exterior.xy
                ax_grid.fill(bx, by, color='gray', alpha=0.5)
                
        valid_colors = mcolors.CSS4_COLORS
        for _, det in st.session_state.det_data.iterrows():
            c = det['Color'].lower()
            use_c = c if c in valid_colors else 'blue'
            ax_grid.plot(det['X'], det['Y'], '^', color=use_c, markersize=8, markeredgecolor='black')

        ax_grid.set_aspect('equal')
        ax_grid.grid(True, linestyle='--', alpha=0.5)
        st.pyplot(fig_grid)
    else:
        st.error("Phòng cần ít nhất 3 góc!")
        room_poly = None

with col_input2:
    st.header("2. Bố trí Thiết bị & Khí")
    
    with st.expander("🚧 Danh sách Vật cản (Cylinder / Box)", expanded=True):
        edited_obs = st.data_editor(st.session_state.obs_data, num_rows="dynamic", use_container_width=True,
                                    column_config={"Type": st.column_config.SelectboxColumn("Loại", options=["Cylinder", "Box"])})
        st.session_state.obs_data = edited_obs

    with st.expander("⚙️ Cấu hình Các Phân hệ Khí (Bấm '+' thêm)", expanded=True):
        edited_auto_config = st.data_editor(st.session_state.auto_config, num_rows="dynamic", use_container_width=True,
            column_config={
                "Layer": st.column_config.SelectboxColumn("Mặt phẳng", options=["Khí Nhẹ (Sát trần)", "Khí Trung bình (Vùng thở)", "Khí Nặng (Sát sàn)"]),
                "Color": st.column_config.SelectboxColumn("Màu bản đồ", options=["cyan", "magenta", "yellow", "lime", "red", "blue", "orange"])
            })
        st.session_state.auto_config = edited_auto_config

        if st.button("🚀 Tự động Rải Đầu dò", type="primary"):
            if room_poly is not None:
                new_dets = []
                for _, row_cfg in edited_auto_config.iterrows():
                    if "Nhẹ" in row_cfg["Layer"]: z_val = max(room_z - 0.5, 0.5)
                    elif "Nặng" in row_cfg["Layer"]: z_val = 0.5
                    else: z_val = 1.5
                    
                    spacing = row_cfg["Radius"] * 1.5 
                    minx, miny, maxx, maxy = room_poly.bounds
                    nx = max(1, math.ceil((maxx - minx) / spacing))
                    ny = max(1, math.ceil((maxy - miny) / spacing))
                    
                    x_steps = np.linspace(minx + (maxx-minx)/(2*nx), maxx - (maxx-minx)/(2*nx), nx)
                    y_steps = np.linspace(miny + (maxy-miny)/(2*ny), maxy - (maxy-miny)/(2*ny), ny)
                    
                    count = 1
                    for x in x_steps:
                        for y in y_steps:
                            pt = Point(x, y)
                            if room_poly.contains(pt): 
                                new_dets.append({
                                    "ID": f"{row_cfg['Model']} ({count:02d})", 
                                    "Model": row_cfg['Model'], 
                                    "Gas": f"{row_cfg['Target Gas']} ({row_cfg['Layer'].split(' ')[1]})", 
                                    "X": round(x, 1), "Y": round(y, 1),
                                    "Z": z_val, "Radius": row_cfg["Radius"], "Color": row_cfg["Color"]
                                })
                                count += 1
                
                st.session_state.det_data = pd.DataFrame(new_dets)
                st.success(f"Đã rải thành công {len(new_dets)} đầu dò!")
                st.rerun()

    st.write("📋 **Bảng Tọa độ Đầu dò Thực tế:**")
    edited_dets = st.data_editor(st.session_state.det_data, num_rows="dynamic", use_container_width=True)
    st.session_state.det_data = edited_dets

# ==========================================
# THÔNG TIN DỰ ÁN CHO TEMPLATE WORD
# ==========================================
st.markdown("---")
st.header("3. 📝 Thông tin Dự án & BOM")
col_info1, col_info2, col_info3 = st.columns([1, 1, 1])

with col_info1:
    project_name = st.text_input("Tên Dự án / Gói thầu", value="Thiết kế Hệ thống Giám sát Rò rỉ Khí")
    author_name = st.text_input("Người lập báo cáo", value="Cao Minh Lợi - Giám đốc Kỹ thuật")
with col_info2:
    client_name = st.text_input("Đơn vị / Khách hàng", value="Nhà máy ABC")
    report_date = st.date_input("Ngày lập báo cáo")
with col_info3:
    report_number = st.text_input("Số Báo cáo (No.)", value="RKV_TE_001/BC")

bom_items = []
bom_items.append({"STT": 1, "Hạng mục thiết bị": "Tủ điều khiển trung tâm đo khí", "Đơn vị": "Bộ", "Khối lượng": 1})

stt = 2
total_cable_length = 0

if not edited_dets.empty and 'Model' in edited_dets.columns:
    counts = edited_dets['Model'].value_counts()
    for model, qty in counts.items():
        bom_items.append({"STT": stt, "Hạng mục thiết bị": f"Đầu dò đo khí rò rỉ - Model: {model}", "Đơn vị": "Bộ", "Khối lượng": qty})
        stt += 1
    
    for _, d in edited_dets.iterrows():
        cable_up = abs(room_z - panel_z)
        cable_horizontal = abs(d['X'] - panel_x) + abs(d['Y'] - panel_y)
        cable_down = abs(room_z - d['Z'])
        total_cable_length += (cable_up + cable_horizontal + cable_down)
        
    total_cable_length = math.ceil(total_cable_length * (1 + wastage_percent / 100))
else:
    total_cable_length = 0

bom_items.append({"STT": stt, "Hạng mục thiết bị": "Cáp tín hiệu chống nhiễu chuyên dụng", "Đơn vị": "Mét", "Khối lượng": total_cable_length})
stt += 1
bom_items.append({"STT": stt, "Hạng mục thiết bị": "Chuông đèn cảnh báo (Siren/Light)", "Đơn vị": "Bộ", "Khối lượng": len(edited_dets) if not edited_dets.empty else 1})

edited_bom = st.data_editor(pd.DataFrame(bom_items), use_container_width=True, hide_index=True)


# ==========================================
# 4. HÀM XỬ LÝ TOÁN HỌC & ĐỒ HỌA
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
        if any(poly.contains(pt) for poly in obs_polys):
            collisions.append(det['ID'])
    return collisions

def generate_2d_plot(room_poly, obs_polys, df_dets_group, gas_name, px, py):
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
            shadowed = any(line.crosses(obs) for obs in obs_polys)
            if not shadowed: coverage_mask[i] = True

    diem_kha_dung = np.sum(valid_points_mask)
    ty_le = (np.sum(coverage_mask) / diem_kha_dung) * 100 if diem_kha_dung > 0 else 0

    fig, ax = plt.subplots(figsize=(8, 6))
    coverage_2d = coverage_mask.reshape(xx.shape)
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

    fig.add_trace(go.Scatter3d(x=[px], y=[py], z=[pz], mode='markers+text', 
                               marker=dict(symbol='square', size=8, color='red'), 
                               text=["TỦ TRUNG TÂM"], textposition="top center", textfont=dict(color="red", size=12, weight="bold"), name="Control Panel"))

    def get_sphere(x0, y0, z0, r):
        u, v = np.mgrid[0:2*np.pi:15j, 0:np.pi:10j]
        return r*np.cos(u)*np.sin(v)+x0, r*np.sin(u)*np.sin(v)+y0, r*np.cos(v)+z0

    for _, det in df_dets.iterrows():
        hover_label = f"Model: {det['ID']}<br>Mục tiêu: {det['Gas']}<br>Cao độ Z: {det['Z']}m"
        fig.add_trace(go.Scatter3d(x=[det['X']], y=[det['Y']], z=[det['Z']], mode='markers+text', 
                                   marker=dict(size=6, color='white'), 
                                   text=[f"{det['ID']}<br>({det['Gas']})"], textposition="top center", textfont=dict(color="white", size=10),
                                   name=det['ID'], hoverinfo="text", hovertext=hover_label))
        
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
        scene=dict(
            xaxis=dict(range=[minx, maxx], title='X', backgroundcolor="rgb(30,30,30)", gridcolor="gray"),
            yaxis=dict(range=[miny, maxy], title='Y', backgroundcolor="rgb(30,30,30)", gridcolor="gray"),
            zaxis=dict(range=[0, max(rz, 5)], title='Z', backgroundcolor="rgb(30,30,30)", gridcolor="gray"),
            aspectmode='data'
        ),
        paper_bgcolor="rgb(15,15,15)", plot_bgcolor="rgb(15,15,15)",
        margin=dict(l=0, r=0, b=0, t=30), showlegend=False
    )
    return fig

# CÔNG NGHỆ TEMPLATE DOCXTPL
def generate_word_template(template_path, figs_dict, img_3d_bytes, bom_df, p_name, c_name, author, r_date, r_num):
    doc = DocxTemplate(template_path)
    
    img_3d_obj = None
    if img_3d_bytes:
        img_3d_obj = InlineImage(doc, img_3d_bytes, width=Inches(6.0))
        
    bom_list = []
    for _, row in bom_df.iterrows():
        bom_list.append({
            'stt': row['STT'],
            'name': row['Hạng mục thiết bị'],
            'unit': row['Đơn vị'],
            'qty': f"{int(row['Khối lượng']):,}"
        })
        
    gas_maps_list = []
    for gas_name, fig_info in figs_dict.items():
        img_stream = io.BytesIO()
        fig_info['fig'].savefig(img_stream, format='png', bbox_inches='tight', dpi=150)
        img_stream.seek(0)
        gas_maps_list.append({
            'gas_name': gas_name,
            'coverage': f"{fig_info['coverage']:.1f}",
            'img_2d': InlineImage(doc, img_stream, width=Inches(6.0))
        })
        
    context = {
        'report_number': r_num,
        'project_name': p_name.upper(),
        'client_name': c_name,
        'author_name': author,
        'report_date': r_date.strftime('%d/%m/%Y'),
        'img_3d': img_3d_obj,
        'bom_items': bom_list,
        'gas_maps': gas_maps_list
    }
    
    doc.render(context)
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream


# ==========================================
# 5. TRIGGER KẾT XUẤT
# ==========================================
st.markdown("---")
if st.button("📊 Chạy Mô phỏng Đồ họa & Tải Báo cáo Kỹ thuật", use_container_width=True, type='primary'):
    if not os.path.exists("Mau_Bao_Cao.docx"):
        st.error("🚨 Không tìm thấy file `Mau_Bao_Cao.docx`. Vui lòng tải 'File Mẫu Chuẩn' ở cột bên trái, chèn logo và Upload lên GitHub!")
    elif room_poly is None or edited_dets.empty:
        st.warning("⚠️ Vui lòng nhập đủ tọa độ phòng và danh sách đầu dò!")
    else:
        try:
            obs_polys = create_obstacle_polys(edited_obs)
            collided = check_collision_shapely(edited_dets, obs_polys)
            
            if collided:
                st.error(f"⛔ LỖI VA CHẠM: Đầu dò **{', '.join(collided)}** đang bị đặt nằm bên trong vật cản! Vui lòng chỉnh lại tọa độ.")
            else:
                with st.spinner('Đang nhúng dữ liệu và hình ảnh AI vào Template Riken Viet...'):
                    st.header("4. Phân tích Kết quả Đồ họa")
                    
                    fig_3d = generate_plotly_3d_complex(room_poly, room_z, obs_polys, edited_obs, edited_dets, panel_x, panel_y, panel_z)
                    st.plotly_chart(fig_3d, use_container_width=True)
                    
                    img_3d_bytes = io.BytesIO()
                    try:
                        fig_3d.write_image(img_3d_bytes, format='png', width=800, height=500)
                        img_3d_bytes.seek(0)
                    except: img_3d_bytes = None

                    gas_groups = edited_dets['Gas'].unique()
                    tabs = st.tabs([f"Bản đồ: {g}" for g in gas_groups])
                    generated_figs = {} 
                    
                    for i, gas_name in enumerate(gas_groups):
                        with tabs[i]:
                            df_group = edited_dets[edited_dets['Gas'] == gas_name]
                            fig_2d, coverage = generate_2d_plot(room_poly, obs_polys, df_group, gas_name, panel_x, panel_y)
                            st.pyplot(fig_2d)
                            if coverage >= 80:
                                st.success(f"✅ Tỷ lệ bao phủ của {gas_name} đạt chuẩn: {coverage:.1f}%")
                            else:
                                st.warning(f"⚠️ Tỷ lệ bao phủ của {gas_name} chỉ đạt: {coverage:.1f}%")
                            generated_figs[gas_name] = {'fig': fig_2d, 'coverage': coverage}
                    
                    word_stream = generate_word_template("Mau_Bao_Cao.docx", generated_figs, img_3d_bytes, edited_bom, project_name, client_name, author_name, report_date, report_number)
                    st.download_button("📄 Tải Báo cáo chuẩn Form Công ty", word_stream, f"{report_number.replace('/','_')}_{client_name}.docx", type="primary")

        except Exception as e:
            st.error(f"Lỗi hệ thống: {e}. Vui lòng tải lại File Mẫu Chuẩn ở cột bên trái.")

st.markdown("""
    <hr style="border: 0; height: 1px; background-image: linear-gradient(to right, rgba(255, 255, 255, 0), rgba(255, 255, 255, 0.2), rgba(255, 255, 255, 0)); margin-top: 50px;">
    <div style="text-align: center; color: #888888; font-size: 14px; padding-bottom: 20px;">
        &copy; 2026 All Rights Reserved.<br>
        Designed and programmed by <b>trggiang</b>.
    </div>
""", unsafe_allow_html=True)
