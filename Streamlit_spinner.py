import streamlit as st
import streamlit.components.v1 as components
import numpy as np

st.set_page_config(page_title="Executive Spinner", layout="wide")

# --- 1. Initialize Session State ---
if 'items' not in st.session_state:
    st.session_state['items'] = [{"name": "Pizza", "percent": 50.0}, {"name": "Tacos", "percent": 50.0}]

# --- 2. Sidebar Settings ---
with st.sidebar:
    st.title("⚙️ Settings")
    password = st.text_input("Password", type="password")
    
    # Track items and weights
    current_items = []
    for i, item in enumerate(st.session_state['items']):
        cols = st.columns([2, 1, 0.5])
        name = cols[0].text_input(f"Name {i}", value=item['name'], key=f"n_in_{i}", label_visibility="collapsed")
        perc = cols[1].number_input(f"P {i}", value=float(item['percent']), key=f"p_in_{i}", label_visibility="collapsed")
        
        if cols[2].button("🗑️", key=f"del_{i}"):
            st.session_state['items'].pop(i)
            st.rerun()
        current_items.append({"name": name, "percent": perc})

    if st.button("➕ Add Item"):
        st.session_state['items'].append({"name": "New Item", "percent": 50.0})
        st.rerun()

    # Admin Rigging
    manual_rig = None
    if password == "mc2026":
        st.success("Admin Active")
        manual_rig = st.selectbox("Force Winner?", [None] + [x['name'] for x in current_items])
    
    if st.button("🔄 Reset Wheel State"):
        st.rerun()

# --- 3. Logic: Calculate Randomized Landing ---
def get_wheel_data(items_list, rigged_name):
    names = [x['name'] for x in items_list]
    weights = [x['percent'] for x in items_list]
    total_w = sum(weights) if sum(weights) > 0 else 1
    
    # Step 1: Determine the Winner
    if rigged_name and rigged_name in names:
        winner_name = rigged_name
    else:
        norm_weights = [w/total_w for w in weights]
        winner_name = np.random.choice(names, p=norm_weights)
    
    # Step 2: Find the angular boundaries of that winner
    winner_idx = names.index(winner_name)
    start_degree = (sum(weights[:winner_idx]) / total_w) * 360
    end_degree = (sum(weights[:winner_idx+1]) / total_w) * 360
    
    # Step 3: Pick a random spot WITHIN that slice (with 2-degree padding from edges)
    # This prevents the wheel from landing exactly on the same spot every time.
    padding = 2
    if (end_degree - start_degree) > (padding * 2):
        random_spot_in_slice = np.random.uniform(start_degree + padding, end_degree - padding)
    else:
        random_spot_in_slice = (start_degree + end_degree) / 2

    # Step 4: Calculate Rotation
    # 270 degrees is the "Top" of the SVG. 
    # We subtract our target spot from 270 to rotate it to the pointer.
    total_rotation = (270 - random_spot_in_slice) + (360 * 5) # 5 full spins
    
    return winner_name, int(total_rotation)

winner_name, target_angle = get_wheel_data(current_items, manual_rig)

# --- 4. Build the SVG ---
def build_svg_code(items_list):
    colors = ["#FF4B4B", "#1C83E1", "#00C781", "#FFBB00", "#7D3CFF", "#FF4B91"]
    total_p = sum(x['percent'] for x in items_list) if sum(x['percent'] for x in items_list) > 0 else 1
    
    svg_parts = ""
    current_angle = 0
    for i, item in enumerate(items_list):
        sweep = (item['percent'] / total_p) * 360
        color = colors[i % len(colors)]
        
        # Math for SVG paths
        rad_s, rad_e = np.radians(current_angle), np.radians(current_angle + sweep)
        x1, y1 = 150 + 100 * np.cos(rad_s), 150 + 100 * np.sin(rad_s)
        x2, y2 = 150 + 100 * np.cos(rad_e), 150 + 100 * np.sin(rad_e)
        
        large_arc = 1 if sweep > 180 else 0
        svg_parts += f'<path d="M150,150 L{x1},{y1} A100,100 0 {large_arc},1 {x2},{y2} Z" fill="{color}" stroke="white" stroke-width="1"/>'
        
        # Text labels
        mid_rad = np.radians(current_angle + sweep/2)
        tx, ty = 150 + 75 * np.cos(mid_rad), 150 + 75 * np.sin(mid_rad)
        svg_parts += f'<text x="{tx}" y="{ty}" fill="white" font-size="10" font-weight="bold" text-anchor="middle" transform="rotate({current_angle + sweep/2}, {tx}, {ty})">{item["name"]}</text>'
        
        current_angle += sweep
    return svg_parts

# --- 5. Main UI Rendering ---
st.title("🎡 Club Decision Wheel")
c1, c2 = st.columns([1.2, 0.8])

with c1:
    svg_html = build_svg_code(current_items)
    
    # We display result purely in JS to avoid the DeltaGenerator error
    wheel_ui = f"""
    <div style="display: flex; flex-direction: column; align-items: center; background: #0e1117; padding: 30px; border-radius: 20px;">
        <div style="width: 0; height: 0; border-left: 15px solid transparent; border-right: 15px solid transparent; border-top: 20px solid #FFBB00; margin-bottom: -10px; z-index: 10;"></div>
        
        <div id="wheel_container" style="transition: transform 4s cubic-bezier(0.15, 0, 0.15, 1); transform: rotate(0deg);">
            <svg width="300" height="300" viewBox="0 0 300 300">{svg_html}</svg>
        </div>
        
        <button id="spin_button" style="margin-top: 30px; padding: 15px 50px; background: #FF4B4B; color: white; border: none; border-radius: 10px; cursor: pointer; font-size: 20px; font-weight: bold; box-shadow: 0 4px #990000;">SPIN</button>
        
        <div id="result_display" style="display: none; margin-top: 30px; text-align: center; animation: fadeIn 1s;">
            <h1 style="color: #FFBB00; font-size: 40px; margin: 0;">🎊 WINNER 🎊</h1>
            <h2 style="color: white; font-size: 30px; margin: 10px 0;">{winner_name}</h2>
            <button onclick="window.parent.location.reload()" style="background: none; border: 1px solid #555; color: #888; cursor: pointer; padding: 5px 10px; border-radius: 5px;">Click to Reset for Next Spin</button>
        </div>
    </div>

    <style>
        @keyframes fadeIn {{ from {{ opacity: 0; }} to {{ opacity: 1; }} }}
        #spin_button:active {{ transform: translateY(4px); box-shadow: none; }}
    </style>

    <script>
        const btn = document.getElementById('spin_button');
        const wheel = document.getElementById('wheel_container');
        const res = document.getElementById('result_display');

        btn.onclick = () => {{
            wheel.style.transform = "rotate({target_angle}deg)";
            btn.style.display = "none";
            
            setTimeout(() => {{
                res.style.display = "block";
            }}, 4100);
        }};
    </script>
    """
    components.html(wheel_ui, height=650)

with c2:
    st.markdown("### 📝 How it works")
    st.write("1. Enter items and weights in the sidebar.")
    st.write("2. If you have the password, you can force a winner.")
    st.write("3. Click **SPIN** and wait for the wheel to stop.")
    
    st.divider()
    st.info("The wheel uses physics-based easing and randomized landing positions to ensure every spin feels unique!")