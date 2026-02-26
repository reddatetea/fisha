df_grouped = df.groupby("order_status")[["payment_value", "freight_value"]].agg(
    total_payment="sum",
    total_freight="sum"
)


df_grouped = df.groupby(["order_status", "payment_type"])["payment_value"].sum()



df_grouped = df.groupby("order_status")["price"].mean()


df_grouped = df.groupby("order_status")[["payment_value", "freight_value"]].agg(
    total_payment="sum",
    total_freight="sum")