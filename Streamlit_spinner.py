import streamlit as st
import streamlit.components.v1 as components
import numpy as np

st.set_page_config(page_title="Executive Spinner", layout="wide")

# --- Global CSS ---
st.markdown("""
<style>
.block-container {
    padding-top: 1rem !important;
}
/* Hide the signal button container immediately via CSS before JS even runs */
div[data-testid="stButton"]:has(> button p) button p {
    /* fallback */
}
.signal-btn-hide {
    display: none !important;
    visibility: hidden !important;
    height: 0 !important;
    overflow: hidden !important;
    margin: 0 !important;
    padding: 0 !important;
}
</style>
""", unsafe_allow_html=True)

# --- Helper to reset state ---
def reset_result():
    st.session_state.reveal_winner = False
    st.session_state.winner = None
    st.session_state.angle = None

# --- Initialize Session State ---
if 'items' not in st.session_state:
    st.session_state['items'] = [
        {"name": "Pizza", "percent": 50.0},
        {"name": "Tacos", "percent": 50.0},
        {"name": "Pasta", "percent": 50.0},
        {"name": "Chocolate", "percent": 50.0}
    ]
if 'reveal_winner' not in st.session_state:
    st.session_state.reveal_winner = False
if 'winner' not in st.session_state:
    st.session_state.winner = None
if 'angle' not in st.session_state:
    st.session_state.angle = None

# --- Sidebar ---
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

# --- Wheel logic ---
def get_wheel_data(items_list, rigged_name):
    names = [x['name'] for x in items_list]
    weights = [x['percent'] for x in items_list]
    total_w = sum(weights) if sum(weights) > 0 else 1

    if rigged_name and rigged_name in names:
        winner_name = rigged_name
    else:
        norm_weights = [w / total_w for w in weights]
        winner_name = np.random.choice(names, p=norm_weights)

    winner_idx = names.index(winner_name)
    start_degree = (sum(weights[:winner_idx]) / total_w) * 360
    end_degree = (sum(weights[:winner_idx + 1]) / total_w) * 360

    padding = 5
    random_spot = np.random.uniform(start_degree + padding, end_degree - padding)
    total_rotation = (270 - random_spot) + (360 * 5)
    return winner_name, int(total_rotation)

if st.session_state.winner is None:
    winner, angle = get_wheel_data(current_items, manual_rig)
    st.session_state.winner = winner
    st.session_state.angle = angle
else:
    winner = st.session_state.winner
    angle = st.session_state.angle

# --- Hidden signal button ---
# Wrapped in a div with a unique class that CSS sets to display:none immediately.
st.markdown('<div class="signal-btn-hide">', unsafe_allow_html=True)
signal_clicked = st.button("SPIN_COMPLETE_SIGNAL", key="spin_signal")
st.markdown('</div>', unsafe_allow_html=True)

if signal_clicked:
    st.session_state.reveal_winner = True
    st.rerun()

# --- Main UI ---
st.title("🎡 Club Decision Wheel")
col_wheel, col_info = st.columns([1.2, 0.8])

with col_wheel:
    colors = ["#FF4B4B", "#1C83E1", "#00C781", "#FFBB00", "#7D3CFF", "#FF4B91"]
    total_p = sum(x['percent'] for x in current_items) if current_items else 1

    # Larger wheel: 500x500 viewBox, radius 220, text radius 165
    W = 500
    CX, CY = 250, 250
    R = 220
    TR = 165

    svg_parts = ""
    current_angle = 0
    for i, item in enumerate(current_items):
        sweep = (item['percent'] / total_p) * 360
        color = colors[i % len(colors)]
        rad_s = np.radians(current_angle)
        rad_e = np.radians(current_angle + sweep)
        x1, y1 = CX + R * np.cos(rad_s), CY + R * np.sin(rad_s)
        x2, y2 = CX + R * np.cos(rad_e), CY + R * np.sin(rad_e)
        large_arc = 1 if sweep > 180 else 0
        svg_parts += f'<path d="M{CX},{CY} L{x1},{y1} A{R},{R} 0 {large_arc},1 {x2},{y2} Z" fill="{color}" stroke="white" stroke-width="2"/>'
        mid_rad = np.radians(current_angle + sweep / 2)
        tx, ty = CX + TR * np.cos(mid_rad), CY + TR * np.sin(mid_rad)
        rot = current_angle + sweep / 2
        svg_parts += f'<text x="{tx}" y="{ty}" fill="white" font-size="16" font-weight="bold" text-anchor="middle" dominant-baseline="middle" transform="rotate({rot},{tx},{ty})">{item["name"]}</text>'
        current_angle += sweep

    wheel_html = f"""
    <div style="display: flex; flex-direction: column; align-items: center; background: #0e1117; padding: 20px; border-radius: 20px;">
        <div style="width: 0; height: 0; border-left: 22px solid transparent; border-right: 22px solid transparent; border-top: 30px solid #FFBB00; margin-bottom: -15px; z-index: 10;"></div>
        <div id="wheel" style="transition: transform 4s cubic-bezier(0.15, 0, 0.15, 1); transform: rotate(0deg);">
            <svg width="{W}" height="{W}" viewBox="0 0 {W} {W}">{svg_parts}</svg>
        </div>
        <button id="spin_btn" style="margin-top: 30px; padding: 18px 60px; background: #FF4B4B; color: white; border: none; border-radius: 10px; cursor: pointer; font-size: 24px; font-weight: bold;">SPIN</button>
    </div>
    <script>
        const btn = document.getElementById('spin_btn');
        const wheel = document.getElementById('wheel');
        btn.onclick = () => {{
            wheel.style.transform = "rotate({angle}deg)";
            btn.disabled = true;
            btn.style.opacity = "0.5";
            setTimeout(() => {{
                const buttons = window.parent.document.querySelectorAll('button');
                for (const b of buttons) {{
                    if (b.innerText.trim() === 'SPIN_COMPLETE_SIGNAL') {{
                        b.click();
                        break;
                    }}
                }}
            }}, 4200);
        }};
    </script>
    """
    components.html(wheel_html, height=620)

with col_info:
    st.markdown("### 📝 Instructions")
    st.write("1. Configure items in the sidebar.")
    st.write("2. Click **SPIN** on the wheel.")

    st.divider()

    if st.session_state.reveal_winner:
        st.markdown(f"""
            <div style="text-align: center; background: #1e2129; padding: 20px; border-radius: 15px; border: 2px solid #FFBB00;">
                <h2 style="color: #FFBB00; margin-bottom: 0;">🎊 WINNER 🎊</h2>
                <h1 style="color: white; margin-top: 10px; font-size: 55px;">{winner}</h1>
            </div>
        """, unsafe_allow_html=True)

        st.write("")
        if st.button("🔄 Spin Again", use_container_width=True, type="primary"):
            reset_result()
            st.rerun()
    else:
        st.info("🎡 The winner will appear here after the wheel stops!")