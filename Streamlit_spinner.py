import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import time
import random

st.set_page_config(page_title="Executive Spinner", layout="wide")

# --- Initialize Session State ---
if 'items' not in st.session_state:
    st.session_state['items'] = [{"name": "Pizza", "percent": 50}, {"name": "Tacos", "percent": 50}]
if 'winner_name' not in st.session_state:
    st.session_state.winner_name = ""
if 'target_angle' not in st.session_state:
    st.session_state.target_angle = 0

# --- Secret Rigging ---
params = st.query_params
url_winner = params.get("win")

with st.sidebar:
    st.title("⚙️ Settings")
    password = st.text_input("Password", type="password")
    
    new_data = []
    total_p = 0
    for i, item in enumerate(st.session_state['items']):
        cols = st.columns([2, 1, 0.5])
        name = cols[0].text_input(f"n{i}", value=item['name'], key=f"n{i}", label_visibility="collapsed")
        perc = cols[1].number_input(f"p{i}", value=float(item['percent']), key=f"p{i}", label_visibility="collapsed")
        if cols[2].button("🗑️", key=f"d{i}"):
            st.session_state['items'].pop(i)
            st.rerun()
        new_data.append({"name": name, "percent": perc})
        total_p += perc

    if st.button("➕ Add Item"):
        st.session_state['items'].append({"name": "New Item", "percent": 0})
        st.rerun()

    manual_rig = None
    if password == "mc2026":
        st.success("Admin Active")
        manual_rig = st.selectbox("Force Winner?", [None] + [x['name'] for x in new_data])

# --- Logic to Determine Winner ---
def calculate_winner():
    names = [x['name'] for x in new_data]
    weights = [x['percent'] for x in new_data]
    sum_w = sum(weights) if sum(weights) > 0 else 1
    norm_weights = [w/sum_w for w in weights]
    
    if url_winner:
        winner = url_winner
    elif manual_rig:
        winner = manual_rig
    else:
        winner = np.random.choice(names, p=norm_weights)
    
    idx = names.index(winner)
    start_deg = sum(norm_weights[:idx]) * 360
    end_deg = start_deg + (norm_weights[idx] * 360)
    mid_deg = (start_deg + end_deg) / 2
    
    # Calculate stop angle (5 full spins + target)
    st.session_state.target_angle = (360 - mid_deg) + (360 * 5)
    st.session_state.winner_name = winner

# --- Build the SVG Wheel ---
st.title("🎡 Club Decision Wheel")

wheel_colors = ["#FF4B4B", "#1C83E1", "#00C781", "#FFBB00", "#7D3CFF", "#FF4B91"]
svg_elements = ""
current_angle = 0

for i, item in enumerate(new_data):
    size = (item['percent'] / total_p) * 360 if total_p > 0 else 0
    color = wheel_colors[i % len(wheel_colors)]
    
    # Path math
    x1 = 150 + 100 * np.cos(np.radians(current_angle))
    y1 = 150 + 100 * np.sin(np.radians(current_angle))
    x2 = 150 + 100 * np.cos(np.radians(current_angle + size))
    y2 = 150 + 100 * np.sin(np.radians(current_angle + size))
    
    large_arc = 1 if size > 180 else 0
    svg_elements += f'<path d="M150,150 L{x1},{y1} A100,100 0 {large_arc},1 {x2},{y2} Z" fill="{color}" stroke="#0e1117" stroke-width="2"/>'
    
    # Text math
    mid_angle = current_angle + (size / 2)
    tx = 150 + 60 * np.cos(np.radians(mid_angle))
    ty = 150 + 60 * np.sin(np.radians(mid_angle))
    
    svg_elements += f'<text x="{tx}" y="{ty}" fill="white" font-size="10" font-family="sans-serif" font-weight="bold" text-anchor="middle" alignment-baseline="middle" transform="rotate({mid_angle}, {tx}, {ty})">{item["name"]}</text>'
    
    current_angle += size

# --- Display Area ---
col_left, col_right = st.columns([1, 1])

# Initialize a flag to hide the winner until the spin is done
if 'show_result' not in st.session_state:
    st.session_state.show_result = False

with col_left:
    if st.button("🔄 Prepare New Spin", use_container_width=True):
        calculate_winner()
        st.session_state.show_result = False # Hide old winner
        st.rerun()

    html_code = f"""
    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; background: #0e1117; padding: 20px; border-radius: 10px;">
        <div style="width: 0; height: 0; border-left: 15px solid transparent; border-right: 15px solid transparent; border-top: 20px solid #FFBB00; margin-bottom: -10px; z-index: 10;"></div>
        <div id="wheel-container" style="transition: transform 4s cubic-bezier(0.15, 0, 0.15, 1); transform: rotate(0deg);">
            <svg width="300" height="300" viewBox="0 0 300 300">{svg_elements}</svg>
        </div>
        <button id="spin-btn" style="margin-top: 30px; padding: 15px 40px; background: #FF4B4B; color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 18px; font-weight: bold;">CLICK TO SPIN</button>
    </div>
    <script>
        const btn = document.getElementById('spin-btn');
        const wheel = document.getElementById('wheel-container');
        btn.onclick = () => {{
            wheel.style.transform = "rotate({st.session_state.target_angle}deg)";
            btn.disabled = true;
            btn.style.background = "#333";
            btn.innerText = "SPINNING...";
            // This tells Streamlit to update AFTER the 4-second animation
            setTimeout(() => {{ 
                window.parent.postMessage({{type: 'streamlit:setComponentValue', value: true}}, '*'); 
            }}, 4000);
        }};
    </script>
    """
    # This variable catches the message from the JavaScript above
    spin_finished_signal = components.html(html_code, height=500)
    
    # If the JS sends the signal, set show_result to True
    if spin_finished_signal:
        st.session_state.show_result = True

with col_right:
    # ONLY show the winner if the animation is actually finished
    if st.session_state.show_result and st.session_state.winner_name:
        st.balloons()
        st.markdown(f"<h1 style='text-align: center; color: #FFBB00;'>🎊 WINNER 🎊</h1>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='text-align: center;'>{st.session_state.winner_name}</h2>", unsafe_allow_html=True)
    else:
        st.write("### 🎡 Spin the wheel to see the result!")