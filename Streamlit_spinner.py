import streamlit as st
import streamlit.components.v1 as components
import numpy as np

st.set_page_config(page_title="Executive Spinner", layout="wide")

# --- 1. Helper to reset visibility ---
def reset_result():
    st.session_state.reveal_winner = False

# --- 2. Initialize Session State ---
if 'items' not in st.session_state:
    st.session_state['items'] = [
        {"name": "Pizza", "percent": 50.0}, 
        {"name": "Tacos", "percent": 50.0},
        {"name": "Pasta", "percent": 50.0},
        {"name": "Chocolate", "percent": 50.0}
    ]
if 'reveal_winner' not in st.session_state:
    st.session_state.reveal_winner = False

# --- 3. Sidebar Settings ---
with st.sidebar:
    st.title("⚙️ Settings")
    password = st.text_input("Password", type="password", on_change=reset_result)
    
    current_items = []
    for i, item in enumerate(st.session_state['items']):
        cols = st.columns([2, 1, 0.5])
        name = cols[0].text_input(f"N{i}", value=item['name'], key=f"n_in_{i}", label_visibility="collapsed", on_change=reset_result)
        perc = cols[1].number_input(f"P{i}", value=float(item['percent']), key=f"p_in_{i}", label_visibility="collapsed", on_change=reset_result)
        
        if cols[2].button("🗑️", key=f"del_{i}"):
            st.session_state['items'].pop(i)
            reset_result()
            st.rerun()
        current_items.append({"name": name, "percent": perc})

    if st.button("➕ Add Item"):
        st.session_state['items'].append({"name": "New Item", "percent": 50.0})
        reset_result()
        st.rerun()

    manual_rig = None
    if password == "mc2026":
        st.success("Admin Active")
        manual_rig = st.selectbox("Force Winner?", [None] + [x['name'] for x in current_items], on_change=reset_result)
    
    if st.button("🔄 Reset Wheel State"):
        reset_result()
        st.rerun()

# --- 4. Logic: Randomized Landing ---
def get_wheel_data(items_list, rigged_name):
    names = [x['name'] for x in items_list]
    weights = [x['percent'] for x in items_list]
    total_w = sum(weights) if sum(weights) > 0 else 1
    
    if rigged_name and rigged_name in names:
        winner_name = rigged_name
    else:
        norm_weights = [w/total_w for w in weights]
        winner_name = np.random.choice(names, p=norm_weights)
    
    winner_idx = names.index(winner_name)
    start_degree = (sum(weights[:winner_idx]) / total_w) * 360
    end_degree = (sum(weights[:winner_idx+1]) / total_w) * 360
    
    padding = 5
    random_spot = np.random.uniform(start_degree + padding, end_degree - padding)
    total_rotation = (270 - random_spot) + (360 * 5)
    
    return winner_name, int(total_rotation)

winner, angle = get_wheel_data(current_items, manual_rig)

# --- 5. Main UI Rendering ---
st.title("🎡 Club Decision Wheel")
col_wheel, col_info = st.columns([1.2, 0.8])

with col_wheel:
    colors = ["#FF4B4B", "#1C83E1", "#00C781", "#FFBB00", "#7D3CFF", "#FF4B91"]
    total_p = sum(x['percent'] for x in current_items) if sum(x['percent'] for x in current_items) > 0 else 1
    svg_parts = ""
    current_angle = 0
    for i, item in enumerate(current_items):
        sweep = (item['percent'] / total_p) * 360
        color = colors[i % len(colors)]
        rad_s, rad_e = np.radians(current_angle), np.radians(current_angle + sweep)
        x1, y1 = 150 + 100 * np.cos(rad_s), 150 + 100 * np.sin(rad_s)
        x2, y2 = 150 + 100 * np.cos(rad_e), 150 + 100 * np.sin(rad_e)
        large_arc = 1 if sweep > 180 else 0
        svg_parts += f'<path d="M150,150 L{x1},{y1} A100,100 0 {large_arc},1 {x2},{y2} Z" fill="{color}" stroke="white" stroke-width="1"/>'
        mid_rad = np.radians(current_angle + sweep/2)
        tx, ty = 150 + 75 * np.cos(mid_rad), 150 + 75 * np.sin(mid_rad)
        svg_parts += f'<text x="{tx}" y="{ty}" fill="white" font-size="10" font-weight="bold" text-anchor="middle" transform="rotate({current_angle + sweep/2}, {tx}, {ty})">{item["name"]}</text>'
        current_angle += sweep

    wheel_html = f"""
    <div style="display: flex; flex-direction: column; align-items: center; background: #0e1117; padding: 20px; border-radius: 20px;">
        <div style="width: 0; height: 0; border-left: 15px solid transparent; border-right: 15px solid transparent; border-top: 20px solid #FFBB00; margin-bottom: -10px; z-index: 10;"></div>
        <div id="wheel" style="transition: transform 4s cubic-bezier(0.15, 0, 0.15, 1); transform: rotate(0deg);">
            <svg width="300" height="300" viewBox="0 0 300 300">{svg_parts}</svg>
        </div>
        <button id="spin_btn" style="margin-top: 30px; padding: 15px 50px; background: #FF4B4B; color: white; border: none; border-radius: 10px; cursor: pointer; font-size: 20px; font-weight: bold;">SPIN</button>
    </div>
    <script>
        const btn = document.getElementById('spin_btn');
        const wheel = document.getElementById('wheel');
        btn.onclick = () => {{
            wheel.style.transform = "rotate({angle}deg)";
            btn.disabled = true;
            btn.style.opacity = "0.5";
            setTimeout(() => {{
                window.parent.postMessage({{type: 'streamlit:setComponentValue', value: true}}, '*');
            }}, 4100);
        }};
    </script>
    """
    
    # Corrected method to call the HTML component
    spin_signal = components.html(wheel_html, height=450)
    
    if spin_signal:
        st.session_state.reveal_winner = True

with col_info:
    st.markdown("### 📝 Instructions")
    st.write("1. Configure items in the sidebar.")
    st.write("2. Click **SPIN** on the wheel.")
    
    st.divider()

    # The Logic Gate ensures the winner stays hidden until the signal is received
    if st.session_state.reveal_winner:
        st.markdown(f"""
            <div style="text-align: center; background: #1e2129; padding: 20px; border-radius: 15px; border: 2px solid #FFBB00;">
                <h2 style="color: #FFBB00; margin-bottom: 0;">🎊 WINNER 🎊</h2>
                <h1 style="color: white; margin-top: 10px; font-size: 45px;">{winner}</h1>
            </div>
        """, unsafe_allow_html=True)
        
        st.write("") 
        if st.button("🔄 Reset for Next Spin", use_container_width=True):
            reset_result()
            st.rerun()
    else:
        st.info("🎡 The winner will appear here after the wheel stops!")