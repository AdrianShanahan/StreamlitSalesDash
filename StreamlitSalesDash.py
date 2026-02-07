import datetime
from random import choice
import pandas as pd
import numpy as np  
import streamlit as st
from numpy.polynomial import polynomial as P
import matplotlib.pyplot as plt

start_date = datetime.datetime(2025, 1, 1)
end_date = datetime.datetime(2026, 1, 1)
num_days = (end_date - start_date).days

def generate_data():
    num_points = 100
    days = np.arange(num_points)

    dates = [start_date + datetime.timedelta(days=int(day)) for day in days]

    # --- Price: slow drift + noise ---
    price = np.zeros(num_points)
    price[0] = np.random.randint(20, 40)

    for i in range(1, num_points):
        price[i] = price[i-1] + np.random.normal(0, 3)

    price = np.clip(price, 5, 100)

    # --- Quantity: similar but noisier ---
    quantity = np.zeros(num_points)
    quantity[0] = np.random.randint(5, 15)

    for i in range(1, num_points):
        quantity[i] = quantity[i-1] + .05 + np.random.normal(0, 1)

    quantity = np.clip(quantity, 1, 30)

    revenue = price * quantity

    df = pd.DataFrame({
        "Date": dates,
        "Price": price.round(2),
        "Quantity": quantity.round(0).astype(int),
        "Revenue": revenue.round(2)
    })

    return df


if "df" not in st.session_state:
    st.session_state.df = generate_data()

if st.button("Regenerate Data"):
    st.session_state.df = generate_data()

min_date = "2025-01-01"
max_date = "2026-01-01"

if "start_date" not in st.session_state:
    st.session_state.start_date = min_date

if "end_date" not in st.session_state:
    st.session_state.end_date = max_date

col1, col2 = st.columns(2)

with col1:
    st.session_state.start_date = st.date_input(
        "Start Date",
        min_value=min_date,
        max_value=max_date,
        value=st.session_state.start_date
    )

with col2:
    st.session_state.end_date = st.date_input(
        "End Date",
        min_value=min_date,
        max_value=max_date,
        value=st.session_state.end_date
    )

if st.session_state.start_date > st.session_state.end_date:
    st.error("Start date must be before end date")
    st.stop()


filtered_df = st.session_state.df[
    (st.session_state.df['Date'] >= pd.to_datetime(st.session_state.start_date)) &
    (st.session_state.df['Date'] <= pd.to_datetime(st.session_state.end_date))   
]

avg_price = filtered_df['Price'].mean()
total_revenue = filtered_df['Revenue'].sum()

num_points = len(filtered_df)
fig_width = max(6, min(16, num_points / 10))

fig, ax = plt.subplots(figsize=(fig_width, 5))  
ax.scatter(filtered_df['Date'], filtered_df['Revenue'], alpha=0.6)

x_numeric = np.arange(len(filtered_df))
coeffs = P.polyfit(x_numeric, filtered_df['Revenue'], 2)
trend_line = P.polyval(x_numeric, coeffs)
ax.plot(filtered_df['Date'], trend_line, color='red', linewidth=2, label='Trend')

ax.set_xlabel('Date')
ax.set_ylabel('Revenue')
ax.set_title('Revenue Trend Over Time')
ax.legend()
plt.xticks(rotation=45)

st.title("Sales Data Analysis")
st.pyplot(fig)

col1, col2 = st.columns(2)

col1.metric("Average Price", f"${avg_price:,.2f}")
col2.metric("Total Revenue", f"${total_revenue:,.2f}")
