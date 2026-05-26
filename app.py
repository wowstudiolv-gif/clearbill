
import streamlit as st
import pandas as pd
import re
import sqlite3

#page setup
st.set_page_config(page_title="ClearBill", layout="wide")

st.title("ClearBill")
st.subheader("Invoice & Reimbursement Reconciliation")

#database setup
DB_FILE = "clearbill.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_type TEXT,
            vendor TEXT,
            property TEXT,
            house_number TEXT,
            street_name TEXT,
            unit_number TEXT,
            description TEXT,
            bill_month TEXT,
            amount REAL,
            UNIQUE(record_type, house_number, street_name, unit_number, bill_month, amount, description)
        )
    """)

    conn.commit()
    conn.close()

def save_to_db(df, record_type):
    existing = load_database()

    if record_type == "reimbursement":
        duplicate_check_cols = ["House Number", "Street Name", "Unit Number", "Bill Month", "Amount", "Description"]
    else:
        duplicate_check_cols = ["House Number", "Street Name", "Unit Number", "Bill Month", "Amount"]

    df = df.drop_duplicates(subset=duplicate_check_cols)

    if existing.empty:
        new_rows = df
        duplicate_rows = pd.DataFrame()
    else:
        existing_type = existing[existing["record_type"] == record_type]

        if record_type == "reimbursement":
            left_cols = ["House Number", "Street Name", "Unit Number", "Bill Month", "Amount", "Description"]
            right_cols = ["house_number", "street_name", "unit_number", "bill_month", "amount", "description"]
        else:
            left_cols = ["House Number", "Street Name", "Unit Number", "Bill Month", "Amount"]
            right_cols = ["house_number", "street_name", "unit_number", "bill_month", "amount"]

        check = df.merge(
            existing_type,
            left_on=left_cols,
            right_on=right_cols,
            how="left",
            indicator=True
        )

        duplicate_rows = check[check["_merge"] == "both"][df.columns]
        new_rows = check[check["_merge"] == "left_only"][df.columns]

    conn = sqlite3.connect(DB_FILE)

    for _, row in new_rows.iterrows():
        conn.execute("""
            INSERT INTO records (
                record_type, vendor, property, house_number, street_name,
                unit_number, description, bill_month, amount
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            record_type,
            row["Vendor"],
            row["Property"],
            row["House Number"],
            row["Street Name"],
            row["Unit Number"],
            row["Description"],
            row["Bill Month"],
            row["Amount"]
        ))

    conn.commit()
    conn.close()

    return len(new_rows), len(duplicate_rows), duplicate_rows

def load_database():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM records", conn)
    conn.close()
    return df

init_db()

# reset database button
if "confirm_reset" not in st.session_state:
    st.session_state.confirm_reset = False
if st.button("Reset Database"):
    st.session_state.confirm_reset = True
if st.session_state.confirm_reset:
    st.warning("Click again to confirm reset")

    if st.button("⚠️ Confirm Reset"):
        conn = sqlite3.connect(DB_FILE)
        conn.execute("DELETE FROM records")
        conn.commit()
        conn.close()

        st.session_state.processed_invoice_files = set()
        st.session_state.processed_reimbursement_files = set()

        st.session_state.confirm_reset = False
        st.success("Database cleared!")
        st.rerun()

#Reading files
def clean_money(value):
    return (
        pd.Series(value)
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .astype(float)
    )

def split_property(value):
    text = str(value).lower().strip()
    text = re.sub(r'[^\w\s#]', '', text)

    unit_match = re.search(r'#\s*([a-z0-9-]+)', text)
    unit = unit_match.group(1) if unit_match else ""

    text = re.sub(r'#\s*[a-z0-9-]+', '', text)

    parts = text.split()
    house_number = parts[0] if parts else ""
    street_parts = parts[1:]

    # if address ends with a unit number
    if not unit and len(street_parts) > 1 and street_parts[-1].isdigit():
        unit = street_parts[-1]
        street_parts = street_parts[:-1]

    street_name = " ".join(street_parts)

    # remove directional prefixes
    street_name = re.sub(r'\b(n|s|e|w|north|south|east|west)\b', '', street_name)

    # remove suffixes
    street_name = re.sub(
        r'\b(ave|avenue|st|street|rd|road|dr|drive|ln|lane|blvd|court|ct)\b',
        '',
        street_name
    )

    street_name = re.sub(r'\s+', ' ', street_name).strip()

    return pd.Series([house_number, street_name, unit])

def clean_invoice(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "Bill Amount": "Amount"
    })
    df["Amount"] = clean_money(df["Amount"])
    df["Bill Month"] = pd.to_datetime(df["Bill Month"]).dt.to_period("M").astype(str)
    df[["House Number", "Street Name", "Unit Number"]] = df["Property"].apply(split_property)

    return df[[
    "Vendor",
    "Property",
    "House Number",
    "Street Name",
    "Unit Number",
    "Description",
    "Bill Month",
    "Amount"
    ]]

def clean_reimbursement(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "Payee": "Vendor",
        "Bill month": "Bill Month",
        "Reimbursed Amount": "Amount"
    })

    df["Amount"] = clean_money(df["Amount"])
    df["Bill Month"] = pd.to_datetime(df["Bill Month"]).dt.to_period("M").astype(str)
    df[["House Number", "Street Name", "Unit Number"]] = df["Property"].apply(split_property)

    return df[[
    "Vendor",
    "Property",
    "House Number",
    "Street Name",
    "Unit Number",
    "Description",
    "Bill Month",
    "Amount"
    ]]

invoice_files = st.file_uploader(
    "Upload invoice Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

reimbursement_files = st.file_uploader(
    "Upload reimbursement Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if "processed_invoice_files" not in st.session_state:
    st.session_state.processed_invoice_files = set()

if "processed_reimbursement_files" not in st.session_state:
    st.session_state.processed_reimbursement_files = set()

if invoice_files:
    st.write("Uploaded invoice files:")

    for file in invoice_files:
        if file.name in st.session_state.processed_invoice_files:
            st.write(f"{file.name} — already processed")
        else:
            st.write(f"{file.name} — new")

    new_files = [
        f for f in invoice_files
        if f.name not in st.session_state.processed_invoice_files
    ]

    if new_files:
        df_invoice = pd.concat(
            [clean_invoice(file) for file in new_files],
            ignore_index=True
        )

        saved, skipped, skipped_df = save_to_db(df_invoice, "invoice")

        for f in new_files:
            st.session_state.processed_invoice_files.add(f.name)

        df_invoice["is_duplicate"] = df_invoice.apply(
            lambda row: (
                (skipped_df["House Number"] == row["House Number"]) &
                (skipped_df["Street Name"] == row["Street Name"]) &
                (skipped_df["Unit Number"] == row["Unit Number"]) &
                (skipped_df["Bill Month"] == row["Bill Month"]) &
                (skipped_df["Amount"] == row["Amount"])
            ).any() if not skipped_df.empty else False,
            axis=1
        )

        df_display = df_invoice.drop(columns=["is_duplicate"])

        styled_invoice = df_display.style.format({
            "Amount": "${:,.2f}"
        }).apply(
            lambda row: ["background-color: #fff3cd"] * len(row)
            if df_invoice.loc[row.name, "is_duplicate"] else [""] * len(row),
            axis=1
        )

        st.success("New invoice files processed!")
        st.dataframe(styled_invoice)
        st.info(f"Saved {saved} invoice rows. Skipped {skipped} duplicates.")

    else:
        st.info("No new invoice files to process.")

if reimbursement_files:
    st.write("Uploaded reimbursement files:")

    for file in reimbursement_files:
        if file.name in st.session_state.processed_reimbursement_files:
            st.write(f"{file.name} — already processed")
        else:
            st.write(f"{file.name} — new")

    new_files = [
        f for f in reimbursement_files
        if f.name not in st.session_state.processed_reimbursement_files
    ]

    if new_files:
        df_reimb = pd.concat(
            [clean_reimbursement(file) for file in new_files],
            ignore_index=True
        )

        saved, skipped, skipped_df = save_to_db(df_reimb, "reimbursement")

        for f in new_files:
            st.session_state.processed_reimbursement_files.add(f.name)

        df_reimb["is_duplicate"] = df_reimb.apply(
            lambda row: (
                (skipped_df["House Number"] == row["House Number"]) &
                (skipped_df["Street Name"] == row["Street Name"]) &
                (skipped_df["Unit Number"] == row["Unit Number"]) &
                (skipped_df["Bill Month"] == row["Bill Month"]) &
                (skipped_df["Amount"] == row["Amount"]) &
                (skipped_df["Description"] == row["Description"])
            ).any() if not skipped_df.empty else False,
            axis=1
        )

        df_display = df_reimb.drop(columns=["is_duplicate"])

        styled_reimb = df_display.style.format({
            "Amount": "${:,.2f}"
        }).apply(
            lambda row: ["background-color: #fff3cd"] * len(row)
            if df_reimb.loc[row.name, "is_duplicate"] else [""] * len(row),
            axis=1
        )

        st.success("New reimbursement files processed!")
        st.dataframe(styled_reimb)
        st.info(f"Saved {saved} reimbursement rows. Skipped {skipped} duplicates.")

    else:
        st.info("No new reimbursement files to process.")

# manual entry
with st.expander("Manual Entry"):

    entry_type = st.selectbox("Entry type", ["invoice", "reimbursement"])

    with st.form("manual_entry_form"):
        vendor = st.text_input("Vendor / Payee")
        property_value = st.text_input("Property")
        description = st.text_input("Description")
        bill_month = st.text_input("Bill Month or Date", placeholder="2026-05 or 5/8/26")
        amount = st.number_input("Amount", min_value=0.0, step=0.01)

        submitted = st.form_submit_button("Save Manual Entry")

    if submitted:
        try:
            bill_month_parsed = pd.to_datetime(bill_month).to_period("M").strftime("%Y-%m")
        except Exception:
            st.error("Invalid date. Use a format like 2026-05 or 5/8/26.")
            st.stop()

        house_number, street_name, unit_number = split_property(property_value)

        manual_df = pd.DataFrame([{
            "Vendor": vendor,
            "Property": property_value,
            "House Number": house_number,
            "Street Name": street_name,
            "Unit Number": unit_number,
            "Description": description,
            "Bill Month": bill_month_parsed,
            "Amount": amount
        }])

        saved, skipped, skipped_df = save_to_db(manual_df, entry_type)

        if saved:
            st.success(f"Manual {entry_type} saved!")
        else:
            st.warning("This entry already exists in the database.")

# reconciliation
db = load_database()

if not db.empty:
    invoices_db = db[db["record_type"] == "invoice"].copy()
    reimb_db = db[db["record_type"] == "reimbursement"].copy()

    match_cols = ["house_number", "street_name", "unit_number", "amount"]

    # sort oldest first so they get matched first
    invoices_db["bill_month_sort"] = pd.to_datetime(invoices_db["bill_month"])
    reimb_db["bill_month_sort"] = pd.to_datetime(reimb_db["bill_month"])

    invoices_db = invoices_db.sort_values("bill_month_sort")
    reimb_db = reimb_db.sort_values("bill_month_sort")

    invoices_db["amount"] = invoices_db["amount"].round(2)
    reimb_db["amount"] = reimb_db["amount"].round(2)

    # create match index per property+amount
    invoices_db["match_num"] = invoices_db.groupby(match_cols).cumcount()
    reimb_db["match_num"] = reimb_db.groupby(match_cols).cumcount()

    # merge for 1-to-1 matching
    result = invoices_db.merge(
        reimb_db,
        on=match_cols + ["match_num"],
        how="outer",
        suffixes=("_invoice", "_reimbursement"),
        indicator=True
    )

    # status logic
    result["Status"] = result["_merge"].map({
        "both": "Cleared",
        "left_only": "Open",
        "right_only": "Review"
    })

    # rename columns
    result = result.rename(columns={
        "id_invoice": "Invoice ID",
        "id_reimbursement": "Reimb. ID",
        "property_invoice": "Property",
        "vendor_invoice": "Invoice Vendor",
        "description_reimbursement": "Reimbursement Description",
        "bill_month_invoice": "Invoice Month",
        "bill_month_reimbursement": "Reimbursement Month",
        "amount": "Amount"
    })
    result["Property"] = result["Property"].fillna(result["property_reimbursement"]).fillna("")
    result["ID"] = result["Invoice ID"].fillna(result["Reimb. ID"]).astype(int)

    # select + order columns
    result = result[[
        "ID",
        "Property",
        "Invoice Vendor",
        "Reimbursement Description",
        "Invoice Month",
        "Reimbursement Month",
        "Amount",
        "Status"
    ]]
    result = result.sort_values("Status", ascending=False)

    # coloring
    def highlight_status(row):
        if row["Status"] == "Cleared":
            return ["background-color: #c6f7c3"] * len(row)
        elif row["Status"] == "Open":
            return ["background-color: #f7c6c6"] * len(row)
        elif row["Status"] == "Review":
            return ["background-color: #c6e2f7"] * len(row)
        else:
            return [""] * len(row)

    styled = result.style.format({
        "Amount": "${:,.2f}"
    }).apply(highlight_status, axis=1)

    st.subheader("Reconciliation Result")
    st.dataframe(styled)
    st.download_button(
        "Download CSV",
        result.to_csv(index=False),
        file_name="reconciliation.csv",
        mime="text/csv"
    )

with st.expander("Delete Entry"):
    delete_id = st.number_input("Enter ID to delete", min_value=1, step=1)

    if st.button("Delete Entry"):
        conn = sqlite3.connect(DB_FILE)
        conn.execute("DELETE FROM records WHERE id = ?", (delete_id,))
        conn.commit()
        conn.close()

        st.success(f"Deleted entry ID {delete_id}")
        st.rerun()