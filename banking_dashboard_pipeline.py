"""
Banking Analytics Pipeline
==========================
Connects to MySQL (banking_analytics DB), pulls data with pandas,
runs analysis, and auto-generates a formatted Excel dashboard
with charts using xlsxwriter.

Requirements:
    pip install mysql-connector-python pandas xlsxwriter

Before running:
    1. Make sure MySQL is running and mysql_setup.sql has been executed
       (database + tables created + CSVs loaded).
    2. Update DB_CONFIG below with your MySQL credentials.
"""

import mysql.connector
import pandas as pd
from datetime import datetime

# ----------------------------------------------------------------
# 1. DATABASE CONFIG  (edit these to match your MySQL setup)
# ----------------------------------------------------------------
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "Madhavi@9019",   # <-- change this
    "database": "bank_Project"
}

OUTPUT_FILE = "Banking_Dashboard.xlsx"


def get_connection():
    """Create and return a MySQL connection."""
    return mysql.connector.connect(**DB_CONFIG)


def fetch_dataframe(query, conn):
    """Run a SQL query and return the result as a pandas DataFrame."""
    return pd.read_sql(query, conn)


def load_all_tables(conn):
    """Pull the four core tables from MySQL into pandas DataFrames."""
    customers = fetch_dataframe("SELECT * FROM Customers", conn)
    accounts = fetch_dataframe("SELECT * FROM Accounts", conn)
    transactions = fetch_dataframe("SELECT * FROM Transactions", conn)
    loans = fetch_dataframe("SELECT * FROM Loans", conn)
    return customers, accounts, transactions, loans


# ----------------------------------------------------------------
# 2. ANALYSIS FUNCTIONS  (pandas transforms MySQL data into insights)
# ----------------------------------------------------------------
def analyze_churn_by_geography(customers):
    grp = customers.groupby("Geography").agg(
        Total_Customers=("CustomerID", "count"),
        Churned=("Exited", "sum"),
    )
    grp["Churn_Rate_%"] = round(grp["Churned"] / grp["Total_Customers"] * 100, 2)
    return grp.reset_index().sort_values("Churn_Rate_%", ascending=False)


def analyze_churn_by_gender(customers):
    grp = customers.groupby("Gender").agg(
        Total_Customers=("CustomerID", "count"),
        Churned=("Exited", "sum"),
    )
    grp["Churn_Rate_%"] = round(grp["Churned"] / grp["Total_Customers"] * 100, 2)
    return grp.reset_index()


def analyze_avg_balance_by_activity(customers):
    grp = customers.groupby("IsActiveMember").agg(
        Avg_Balance=("Balance", "mean"),
        Customer_Count=("CustomerID", "count"),
    ).reset_index()
    grp["IsActiveMember"] = grp["IsActiveMember"].map({1: "Active", 0: "Inactive"})
    grp["Avg_Balance"] = round(grp["Avg_Balance"], 2)
    return grp


def analyze_monthly_transactions(transactions):
    df = transactions.copy()
    df["TransactionDate"] = pd.to_datetime(df["TransactionDate"])
    df["Month"] = df["TransactionDate"].dt.to_period("M").astype(str)
    grp = df.groupby("Month").agg(
        Txn_Count=("TransactionID", "count"),
        Net_Amount=("Amount", "sum"),
    ).reset_index()
    grp["Net_Amount"] = round(grp["Net_Amount"], 2)
    return grp


def analyze_loans_by_status(loans):
    grp = loans.groupby("Status").agg(
        Num_Loans=("LoanID", "count"),
        Total_Principal=("PrincipalAmount", "sum"),
    ).reset_index()
    grp["Total_Principal"] = round(grp["Total_Principal"], 2)
    return grp


def analyze_top_customers(customers, n=10):
    cols = ["CustomerID", "Name", "Geography", "Balance"]
    return customers[cols].sort_values("Balance", ascending=False).head(n)


def analyze_products_distribution(customers):
    grp = customers.groupby("NumOfProducts").agg(
        Customer_Count=("CustomerID", "count")
    ).reset_index()
    return grp


# ----------------------------------------------------------------
# 3. EXCEL DASHBOARD BUILDER  (xlsxwriter for formatted sheets + charts)
# ----------------------------------------------------------------
def build_dashboard(analyses: dict, kpis: dict, output_path: str):
    """
    analyses: dict of {sheet_name: DataFrame}
    kpis: dict of {label: value} shown on the Dashboard summary sheet
    """
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book

        # ---- Formats ----
        title_fmt = workbook.add_format({
            "bold": True, "font_size": 16, "font_color": "#FFFFFF",
            "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"
        })
        kpi_label_fmt = workbook.add_format({
            "bold": True, "font_size": 11, "font_color": "#1F4E78",
            "border": 1, "align": "center", "bg_color": "#D9E1F2"
        })
        kpi_value_fmt = workbook.add_format({
            "bold": True, "font_size": 14, "font_color": "#000000",
            "border": 1, "align": "center"
        })
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#1F4E78", "font_color": "white",
            "border": 1, "align": "center"
        })

        # ---- Dashboard summary sheet ----
        dash = workbook.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dash
        dash.merge_range("A1:H1", "BANKING ANALYTICS DASHBOARD", title_fmt)
        dash.set_row(0, 30)

        col = 0
        row = 2
        for label, value in kpis.items():
            dash.write(row, col, label, kpi_label_fmt)
            dash.write(row + 1, col, value, kpi_value_fmt)
            col += 1
        dash.set_column(0, len(kpis) - 1, 20)

        # ---- Write each analysis to its own sheet + add a chart ----
        chart_anchor_row = 20
        for sheet_name, df in analyses.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
            ws = writer.sheets[sheet_name]

            # Header formatting
            for c, col_name in enumerate(df.columns):
                ws.write(0, c, col_name, header_fmt)
            ws.set_column(0, len(df.columns) - 1, 18)

            n_rows = len(df)

            # Pick a chart type based on the sheet.
            # build_chart() is a factory so we can create a fresh chart object
            # for the sheet AND a separate one for the Dashboard summary —
            # xlsxwriter does not allow inserting the same chart object twice.
            def build_chart():
                if sheet_name in ("Churn_by_Geography", "Churn_by_Gender"):
                    c = workbook.add_chart({"type": "column"})
                    c.add_series({
                        "name": "Churn Rate %",
                        "categories": [sheet_name, 1, 0, n_rows, 0],
                        "values": [sheet_name, 1, df.columns.get_loc("Churn_Rate_%"), n_rows, df.columns.get_loc("Churn_Rate_%")],
                        "fill": {"color": "#C0504D"},
                    })
                    c.set_title({"name": f"{sheet_name.replace('_',' ')}"})
                    c.set_x_axis({"name": df.columns[0]})
                    c.set_y_axis({"name": "Churn Rate %"})
                    return c

                elif sheet_name == "Avg_Balance_by_Activity":
                    c = workbook.add_chart({"type": "pie"})
                    c.add_series({
                        "name": "Avg Balance by Activity",
                        "categories": [sheet_name, 1, 0, n_rows, 0],
                        "values": [sheet_name, 1, df.columns.get_loc("Avg_Balance"), n_rows, df.columns.get_loc("Avg_Balance")],
                    })
                    c.set_title({"name": "Avg Balance: Active vs Inactive"})
                    return c

                elif sheet_name == "Monthly_Transactions":
                    c = workbook.add_chart({"type": "line"})
                    c.add_series({
                        "name": "Net Amount",
                        "categories": [sheet_name, 1, 0, n_rows, 0],
                        "values": [sheet_name, 1, df.columns.get_loc("Net_Amount"), n_rows, df.columns.get_loc("Net_Amount")],
                        "line": {"color": "#4472C4", "width": 2.5},
                    })
                    c.set_title({"name": "Monthly Net Transaction Amount"})
                    c.set_x_axis({"name": "Month"})
                    c.set_y_axis({"name": "Net Amount"})
                    return c

                elif sheet_name == "Loans_by_Status":
                    c = workbook.add_chart({"type": "column"})
                    c.add_series({
                        "name": "Total Principal",
                        "categories": [sheet_name, 1, 0, n_rows, 0],
                        "values": [sheet_name, 1, df.columns.get_loc("Total_Principal"), n_rows, df.columns.get_loc("Total_Principal")],
                        "fill": {"color": "#70AD47"},
                    })
                    c.set_title({"name": "Total Loan Principal by Status"})
                    c.set_x_axis({"name": "Status"})
                    c.set_y_axis({"name": "Total Principal"})
                    return c

                elif sheet_name == "Products_Distribution":
                    c = workbook.add_chart({"type": "pie"})
                    c.add_series({
                        "name": "Customers by Num Products",
                        "categories": [sheet_name, 1, 0, n_rows, 0],
                        "values": [sheet_name, 1, df.columns.get_loc("Customer_Count"), n_rows, df.columns.get_loc("Customer_Count")],
                    })
                    c.set_title({"name": "Customers by Number of Products"})
                    return c

                return None  # e.g. Top_Customers - table only, no chart needed

            sheet_chart = build_chart()
            if sheet_chart is not None:
                sheet_chart.set_size({"width": 500, "height": 300})
                ws.insert_chart(chart_anchor_row, 0, sheet_chart)

                # Key charts also get placed on the Dashboard summary sheet —
                # built as a fresh chart object, not the same instance.
                if sheet_name in ("Churn_by_Geography", "Monthly_Transactions",
                                  "Loans_by_Status", "Products_Distribution"):
                    dash_chart = build_chart()
                    dash_chart.set_size({"width": 480, "height": 280})
                    dash_row = {"Churn_by_Geography": 5,
                                "Monthly_Transactions": 5,
                                "Loans_by_Status": 22,
                                "Products_Distribution": 22}[sheet_name]
                    dash_col = {"Churn_by_Geography": 0,
                                "Monthly_Transactions": 8,
                                "Loans_by_Status": 0,
                                "Products_Distribution": 8}[sheet_name]
                    dash.insert_chart(dash_row, dash_col, dash_chart)

    print(f"Dashboard written to: {output_path}")


# ----------------------------------------------------------------
# 4. MAIN PIPELINE
# ----------------------------------------------------------------
def main():
    print("Connecting to MySQL...")
    conn = get_connection()

    print("Loading tables...")
    customers, accounts, transactions, loans = load_all_tables(conn)
    conn.close()

    print("Running analysis...")
    analyses = {
        "Churn_by_Geography": analyze_churn_by_geography(customers),
        "Churn_by_Gender": analyze_churn_by_gender(customers),
        "Avg_Balance_by_Activity": analyze_avg_balance_by_activity(customers),
        "Monthly_Transactions": analyze_monthly_transactions(transactions),
        "Loans_by_Status": analyze_loans_by_status(loans),
        "Products_Distribution": analyze_products_distribution(customers),
        "Top_Customers": analyze_top_customers(customers),
    }

    kpis = {
        "Total Customers": len(customers),
        "Active Members": int(customers["IsActiveMember"].sum()),
        "Churn Rate %": round(customers["Exited"].mean() * 100, 2),
        "Avg Credit Score": round(customers["CreditScore"].mean(), 0),
        "Total Balance (₹)": round(customers["Balance"].sum(), 2),
        "Total Loans": len(loans),
        "Avg Satisfaction": round(customers["SatisfactionScore"].mean(), 2),
    }

    print("Building Excel dashboard...")
    build_dashboard(analyses, kpis, OUTPUT_FILE)

    print("Done. Open", OUTPUT_FILE, "to view your dashboard.")


if __name__ == "__main__":
    main()
