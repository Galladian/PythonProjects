import streamlit as st
import streamlit.components.v1 as components
import numpy as np

st.set_page_config(page_title="Executive Spinner", layout="wide")

# --- 1. Initialize Session State ---
if 'items' not in st.session_state:
    st.session_state['items'] = [{"name": "Pizza", "percent": 50}, {"name": "Tacos", "percent": 50}]
if 'winner_to_show' not in st.session_state:
    st.session_state.winner_to_show = None

# --- 2. Sidebar Settings ---
with st.sidebar:
    st.title("⚙️ Settings")
    password = st.text_input("Password", type="password")
    
    for i, item in enumerate(st.session_state['items']):
        cols = st.columns([2, 1, 0.5])
        item['name'] = cols[0].text_input(f"n{i}", value=item['name'], key=f"name_input_{i}", label_visibility="collapsed")
        item['percent'] = cols[1].number_input(f"p{i}", value=float(item['percent']), key=f"perc_input_{i}", label_visibility="collapsed")
        if cols[2].button("🗑️", key=f"del_btn_{i}"):
            st.session_state['items'].pop(i)
            st.rerun()

    if st.button("➕ Add Item"):
        st.session_state['items'].append({"name": "New Item", "percent": 50})
        st.rerun()

    manual_rig = None
    if password == "mc2026":
        st.success("Admin Active")
        manual_rig = st.selectbox("Force Winner?", [None] + [x['name'] for x in st.session_state['items']])

# --- 3. Logic & SVG Math ---
items = st.session_state['items']
names = [x['name'] for x in items]
weights = [x['percent'] for x in items]
total_w = sum(weights) if sum(weights) > 0 else 1

# Calculate rig/result BEFORE rendering
if manual_rig and manual_rig in names:
    winner = manual_rig
else:
    norm_weights = [w/total_w for w in weights]
    winner = np.random.choice(names, p=norm_weights)

idx = names.index(winner)
prev_sum = sum(weights[:idx]) / total_w * 360
slice_width = weights[idx] / total_w * 360
target_slice_mid = prev_sum + (slice_width / 2)
final_angle = (270 - target_slice_mid) + (360 * 5) 

def build_svg():
    colors = ["#FF4B4B", "#1C83E1", "#00C781", "#FFBB00", "#7D3CFF", "#FF4B91"]
    svg_elements = ""
    curr_angle = 0
    for i, item in enumerate(items):
        size = (item['percent'] / total_w) * 360
        color = colors[i % len(colors)]
        rad_start, rad_end = np.radians(curr_angle), np.radians(curr_angle + size)
        x1, y1 = 150 + 100 * np.cos(rad_start), 150 + 100 * np.sin(rad_start)
        x2, y2 = 150 + 100 * np.cos(rad_end), 150 + 100 * np.sin(rad_end)
        large_arc = 1 if size > 180 else 0
        svg_elements += f'<path d="M150,150 L{x1},{y1} A100,100 0 {large_arc},1 {x2},{y2} Z" fill="{color}" stroke="white" stroke-width="1"/>'
        mid_rad = np.radians(curr_angle + size/2)
        tx, ty = 150 + 75 * np.cos(mid_rad), 150 + 75 * np.sin(mid_rad)
        svg_elements += f'<text x="{tx}" y="{ty}" fill="white" font-size="10" font-weight="bold" text-anchor="middle" transform="rotate({curr_angle + size/2}, {tx}, {ty})">{item["name"]}</text>'
        curr_angle += size
    return svg_elements

# --- 4. Main UI ---
st.title("🎡 Club Decision Wheel")
col_left, col_right = st.columns([1, 1])

with col_left:
    svg_content = build_svg()
    html_code = f"""
    <div style="display: flex; flex-direction: column; align-items: center; background: #0e1117; padding: 20px; border-radius: 15px;">
        <div style="width: 0; height: 0; border-left: 15px solid transparent; border-right: 15px solid transparent; border-top: 20px solid #FFBB00; margin-bottom: -10px; z-index: 10;"></div>
        <div id="wheel" style="transition: transform 4s cubic-bezier(0.15, 0, 0.15, 1); transform: rotate(0deg);">
            <svg width="300" height="300" viewBox="0 0 300 300">{svg_content}</svg>
        </div>
        <button id="spin_btn" style="margin-top: 30px; padding: 15px 40px; background: #FF4B4B; color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 18px; font-weight: bold;">SPIN</button>
    </div>
    <script>
        const btn = document.getElementById('spin_btn');
        const wheel = document.getElementById('wheel');
        btn.onclick = () => {{
            wheel.style.transform = "rotate({final_angle}deg)";
            btn.disabled = true;
            btn.style.opacity = "0.5";
            // No postMessage here to avoid the DeltaGenerator error. 
            // We'll show the result using a timed display in JS or a button.
            setTimeout(() => {{
                document.getElementById('result-text').style.display = 'block';
            }}, 4100);
        }};
    </script>
    <div id="result-text" style="display:none; text-align:center; margin-top:20px;">
        <h2 style="color: #FFBB00;">🎊 WINNER: {winner} 🎊</h2>
        <button onclick="window.parent.location.reload()" style="padding:10px; cursor:pointer;">Next Spin</button>
    </div>
    """
    # CRITICAL: We do NOT assign this to a variable. Just call it.
    components.html(html_code, height=600)

with col_right:
    st.subheader("Instructions")
    st.info("1. Set your items in the sidebar.\n2. Click SPIN on the wheel.\n3. The result will appear under the wheel once it stops.")
    
    # We removed the balloon code from here to prevent random spawning.