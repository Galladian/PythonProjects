import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import time
import json

st.set_page_config(page_title="Executive Spinner", layout="wide")

# --- Initialize Session State ---
if 'items' not in st.session_state:
    st.session_state['items'] = [{"name": "Pizza", "percent": 50}, {"name": "Tacos", "percent": 50}]

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

# --- Logic to Determine Winner BEFORE Animation ---
if 'target_angle' not in st.session_state:
    st.session_state.target_angle = 0
    st.session_state.winner_name = ""

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
    
    # Calculate angle for the JS wheel
    # We find the center of the winner's slice
    idx = names.index(winner)
    start_deg = sum(norm_weights[:idx]) * 360
    end_deg = start_deg + (norm_weights[idx] * 360)
    mid_deg = (start_deg + end_deg) / 2
    
    # 360 - mid_deg because CSS rotation is clockwise, SVG starts at 3 o'clock
    st.session_state.target_angle = (360 - mid_deg) + (360 * 5) # 5 full spins
    st.session_state.winner_name = winner

# --- The Visual Wheel (HTML/JS) ---
st.title("🎡 Club Decision Wheel")

wheel_colors = ["#FF5733", "#33FF57", "#3357FF", "#F333FF", "#FF33A1", "#F3FF33"]

svg_elements = ""
current_angle = 0
for i, item in enumerate(new_data):
    # Calculate slice size
    size = (item['percent'] / total_p) * 360 if total_p > 0 else 0
    color = wheel_colors[i % len(wheel_colors)]
    
    # 1. Draw the Slice (Arc)
    x1 = 150 + 100 * np.cos(np.radians(current_angle))
    y1 = 150 + 100 * np.sin(np.radians(current_angle))
    x2 = 150 + 100 * np.cos(np.radians(current_angle + size))
    y2 = 150 + 100 * np.sin(np.radians(current_angle + size))
    
    large_arc = 1 if size > 180 else 0
    svg_elements += f'<path d="M150,150 L{x1},{y1} A100,100 0 {large_arc},1 {x2},{y2} Z" fill="{color}" stroke="white" stroke-width="1"/>'
    
    # 2. Draw the Text Label
    # We find the middle angle of the slice to place the text
    mid_angle = current_angle + (size / 2)
    # We place the text at 70% of the radius (70px out from center)
    text_x = 150 + 65 * np.cos(np.radians(mid_angle))
    text_y = 150 + 65 * np.sin(np.radians(mid_angle))
    
    # Rotate text so it points toward the center
    text_rotation = mid_angle 
    
    svg_elements = "" 
    current_angle = 0
    for i, item in enumerate(new_data):
        # ... (your math for path and text) ...
        svg_elements += f'<path d="..." ... />'
        svg_elements += f'<text ...>{item["name"]}</text>'
        current_angle += size

html_code = f"""
<div style="display: flex; flex-direction: column; align-items: center; background: #0e1117;">
    <div id="pointer" style="width: 0; height: 0; border-left: 15px solid transparent; border-right: 15px solid transparent; border-top: 20px solid #FFCC00; margin-bottom: -10px; z-index: 10;"></div>
    <div id="wheel-container" style="transition: transform 4s cubic-bezier(0.15, 0, 0.15, 1); transform: rotate(0deg);">
        <svg width="300" height="300" viewBox="0 0 300 300">
            {svg_elements}
        </svg>
    </div>
    <button id="spin-btn" style="margin-top: 20px; padding: 10px 30px; background: #FF4B4B; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;">SPIN!</button>
</div>

<script>
    const btn = document.getElementById('spin-btn');
    const wheel = document.getElementById('wheel-container');
    
    btn.onclick = () => {{
        const targetDeg = {st.session_state.target_angle};
        wheel.style.transform = `rotate(${{targetDeg}}deg)`;
        btn.disabled = true;
        btn.style.opacity = "0.5";
        
        // Tell Streamlit when done (optional flair)
        setTimeout(() => {{
            window.parent.postMessage({{type: 'streamlit:setComponentValue', value: true}}, '*');
        }}, 4000);
    }};
</script>
"""

# Display the Wheel
if st.button("Prepare Spin"):
    calculate_winner()
    st.rerun()

components.html(html_code, height=450)

if st.session_state.winner_name:
    st.write("---")
    st.subheader(f"Result will appear above! (Rigged for: {st.session_state.winner_name})")