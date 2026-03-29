import streamlit as st
import pandas as pd
import numpy as np
import time
import random

# --- Setup Page ---
st.set_page_config(page_title="Executive Spinner", layout="wide")

# --- Initialize Session State ---
if 'items' not in st.session_state:
    st.session_state.items = [{"name": "Pizza", "percent": 50}, {"name": "Tacos", "percent": 50}]

# --- SECRET RIGGING (The URL Method) ---
params = st.query_params
url_winner = params.get("win")

# --- Sidebar ---
with st.sidebar:
    st.title("⚙️ Settings")
    
    # NEW: Admin Password to hide the manual rig
    password = st.text_input("Admin Password", type="password")
    
    new_data = []
    total_p = 0
    
    st.write("Edit Items:")
    for i, item in enumerate(st.session_state.items):
        cols = st.columns([2, 1, 0.5])
        name = cols[0].text_input(f"n{i}", value=item['name'], key=f"n{i}", label_visibility="collapsed")
        perc = cols[1].number_input(f"p{i}", value=float(item['percent']), key=f"p{i}", label_visibility="collapsed")
        if cols[2].button("🗑️", key=f"d{i}"):
            st.session_state.items.pop(i)
            st.rerun()
        new_data.append({"name": name, "percent": perc})
        total_p += perc

    if st.button("➕ Add Item"):
        st.session_state.items.append({"name": "New Item", "percent": 0})
        st.rerun()

    # Manual Rigging (Only shows if password is correct)
    manual_rig = None
    if password == "club2026": # Change this to your secret password
        st.divider()
        st.success("Admin Access Granted")
        manual_rig = st.selectbox("Force Winner?", [None] + [x['name'] for x in new_data])

# --- Main Wheel Area ---
st.title("🎡 Club Decision Wheel")

if st.button("🔥 SPIN THE WHEEL", use_container_width=True, type="primary"):
    placeholder = st.empty()
    
    # Animation
    for _ in range(15):
        random_name = random.choice([x['name'] for x in new_data])
        placeholder.subheader(f"🌀 Spinning... {random_name}")
        time.sleep(0.1)
    
    # Winning Logic
    if url_winner:
        winner = url_winner
    elif manual_rig:
        winner = manual_rig
    else:
        names = [x['name'] for x in new_data]
        weights = [x['percent'] for x in new_data]
        # Avoid division by zero if weights are 0
        sum_weights = sum(weights) if sum(weights) > 0 else 1
        norm_weights = [w/sum_weights for w in weights]
        winner = np.random.choice(names, p=norm_weights)
    
    placeholder.balloons()
    st.header(f"🎊 Result: {winner}")
    st.session_state.last_result = winner

if 'last_result' in st.session_state:
    st.info(f"Previous Winner: {st.session_state.last_result}")