import streamlit as st
import streamlit.components.v1 as components
import numpy as np

st.set_page_config(page_title="Executive Spinner", layout="wide")

# --- 1. Initialize Session State ---
if 'items' not in st.session_state:
    st.session_state['items'] = [{"name": "Pizza", "percent": 50}, {"name": "Tacos", "percent": 50}]
if 'spin_trigger' not in st.session_state:
    st.session_state.spin_trigger = 0 # Counter to trigger JS updates

# --- 2. Sidebar Settings ---
with st.sidebar:
    st.title("⚙️ Settings")
    password = st.text_input("Password", type="password")
    
    # Editable list
    for i, item in enumerate(st.session_state['items']):
        cols = st.columns([2, 1, 0.5])
        item['name'] = cols[0].text_input(f"Name {i}", value=item['name'], key=f"name_{i}", label_visibility="collapsed")
        item['percent'] = cols[1].number_input(f"Weight {i}", value=float(item['percent']), key=f"p_{i}", label_visibility="collapsed")
        if cols[2].button("🗑️", key=f"del_{i}"):
            st.session_state['items'].pop(i)
            st.rerun()

    if st.button("➕ Add Item"):
        st.session_state['items'].append({"name": "New Item", "percent": 50})
        st.rerun()

    # Admin Rigging
    manual_rig = None
    if password == "mc2026":
        st.success("Admin Active")
        manual_rig = st.selectbox("Force Winner?", [None] + [x['name'] for x in st.session_state['items']])

# --- 3. Rigging Logic ---
def get_spin_data():
    items = st.session_state['items']
    names = [x['name'] for x in items]
    weights = [x['percent'] for x in items]
    total_w = sum(weights)
    
    # Determine Winner
    if manual_rig and manual_rig in names:
        winner = manual_rig
    else:
        norm_weights = [w/total_w for w in weights]
        winner = np.random.choice(names, p=norm_weights)
    
    # Calculate Angle
    idx = names.index(winner)
    # Calculate cumulative start and end angles for the winning slice
    prev_sum = sum(weights[:idx]) / total_w * 360
    slice_width = weights[idx] / total_w * 360
    
    # The pointer is at the top (90 deg in CSS circle logic or 0 deg depending on SVG)
    # To land on the slice, we rotate (360 - center_of_slice)
    target_slice_mid = prev_sum + (slice_width / 2)
    final_angle = (360 - target_slice_mid) + (360 * 5) # 5 full spins
    
    return winner, int(final_angle)

# --- 4. Visual Construction ---
def build_svg():
    items = st.session_state['items']
    total_p = sum(x['percent'] for x in items)
    wheel_colors = ["#FF4B4B", "#1C83E1", "#00C781", "#FFBB00", "#7D3CFF", "#FF4B91"]
    
    svg_elements = ""
    curr_angle = 0
    for i, item in enumerate(items):
        size = (item['percent'] / total_p) * 360
        color = wheel_colors[i % len(wheel_colors)]
        
        # SVG path math
        rad_start = np.radians(curr_angle)
        rad_end = np.radians(curr_angle + size)
        x1, y1 = 150 + 100 * np.cos(rad_start), 150 + 100 * np.sin(rad_start)
        x2, y2 = 150 + 100 * np.cos(rad_end), 150 + 100 * np.sin(rad_end)
        
        large_arc = 1 if size > 180 else 0
        svg_elements += f'<path d="M150,150 L{x1},{y1} A100,100 0 {large_arc},1 {x2},{y2} Z" fill="{color}" stroke="white" stroke-width="1"/>'
        
        # Labels
        mid_rad = np.radians(curr_angle + size/2)
        tx, ty = 150 + 70 * np.cos(mid_rad), 150 + 70 * np.sin(mid_rad)
        rot = curr_angle + size/2
        svg_elements += f'<text x="{tx}" y="{ty}" fill="white" font-size="10" font-weight="bold" text-anchor="middle" transform="rotate({rot}, {tx}, {ty})">{item["name"]}</text>'
        
        curr_angle += size
    return svg_elements

# --- 5. Main UI ---
st.title("🎡 Executive Decision Wheel")
col_left, col_right = st.columns([1, 1])

with col_left:
    winner_name, target_deg = get_spin_data()
    svg_content = build_svg()

    # The HTML/JS Component
    # We use a key tied to a counter to force the component to re-render only when we want a "new" spin
    html_code = f"""
    <div style="display: flex; flex-direction: column; align-items: center; font-family: sans-serif;">
        <div style="color: #FFBB00; font-size: 30px; margin-bottom: -10px; z-index: 10;">▼</div>
        <div id="wheel" style="transition: transform 4s cubic-bezier(0.15, 0, 0.15, 1); transform: rotate(0deg);">
            <svg width="300" height="300" viewBox="0 0 300 300">{svg_content}</svg>
        </div>
        <button id="spin_btn" style="margin-top: 20px; padding: 10px 30px; background: #FF4B4B; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;">SPIN WHEEL</button>
    </div>

    <script>
        const btn = document.getElementById('spin_btn');
        const wheel = document.getElementById('wheel');
        btn.onclick = () => {{
            wheel.style.transform = "rotate({target_deg}deg)";
            btn.disabled = true;
            setTimeout(() => {{
                window.parent.postMessage({{type: 'streamlit:setComponentValue', value: '{winner_name}'}}, '*');
            }}, 4100);
        }};
    </script>
    """
    
    result = components.html(html_code, height=450, key=f"wheel_{st.session_state.spin_trigger}")

with col_right:
    st.write("### Result")
    if result:
        st.balloons()
        st.success(f"The winner is: **{result}**")
        if st.button("Prepare Next Spin"):
            st.session_state.spin_trigger += 1
            st.rerun()
    else:
        st.info("Click the 'SPIN WHEEL' button on the left!")