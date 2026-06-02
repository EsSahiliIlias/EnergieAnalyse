import time

import pandas as pd
import pymssql
import requests


# -----------------------
# Credentials
# -----------------------

token_url = "https://app1pub.smappee.net/dev/v3/oauth2/token"

client_id = "165804"
client_secret = "ty2HNwm2AF"
username = "Ilias.EsSahili@oracdecor.be"
password = "Traverse4-Yonder-Skylight-Unsubtle-Epidermal!"


# -----------------------
# Azure SQL
# -----------------------

SQL_SERVER = "smappeedata.database.windows.net"
SQL_DATABASE = "free-sql-db-8161371"
SQL_USERNAME = "ilias"
SQL_PASSWORD = "Il1a58530!"

# -----------------------
# Auth
# -----------------------

def get_access_token():
    payload = {
        "grant_type": "password",
        "client_id": client_id,
        "client_secret": client_secret,
        "username": username,
        "password": password,
    }

    headers = {
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    }

    response = requests.post(
        token_url,
        data=payload,
        headers=headers,
    )

    response.raise_for_status()

    return response.json()["access_token"]


# -----------------------
# Time range
# -----------------------

def get_time_range():
    to_ts = int(time.time() * 1000)

    # 🔥 1 jaar terug
    from_ts = to_ts - (365 * 24 * 60 * 60 * 1000)

    return from_ts, to_ts


# -----------------------
# API call
# -----------------------

def get_consumption(
    service_location_id,
    access_token,
    aggregation,
    from_ts,
    to_ts,
):
    url = (
        f"https://app1pub.smappee.net/dev/v3/"
        f"servicelocation/{service_location_id}/consumption"
    )

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }

    params = {
        "aggregation": aggregation,
        "from": int(from_ts),
        "to": int(to_ts),
    }

    response = requests.get(
        url,
        headers=headers,
        params=params,
    )

    response.raise_for_status()

    return response.json()


# -----------------------
# Transform
# -----------------------

def transform_data(data, service_location_id):
    rows = data["consumptions"]

    df = pd.DataFrame(rows)

    df["service_location_id"] = service_location_id

    df["date"] = pd.to_datetime(
        df["timestamp"],
        unit="ms"
    )

    df["consumption_kwh"] = df["consumption"] / 1000
    df["solar_kwh"] = df["solar"] / 1000
    df["grid_import_kwh"] = df["gridImport"] / 1000
    df["grid_export_kwh"] = df["gridExport"] / 1000

    return df


# -----------------------
# Save to Azure SQL
# -----------------------

def save_to_db(df):
    conn = pymssql.connect(
        server=SQL_SERVER,
        user=SQL_USERNAME,
        password=SQL_PASSWORD,
        database=SQL_DATABASE,
    )

    cursor = conn.cursor()

    for _, row in df.iterrows():

        cursor.execute("""
            INSERT INTO energy_consumption (
                service_location_id,
                timestamp,
                date,
                consumption_kwh,
                solar_kwh,
                grid_import_kwh,
                grid_export_kwh
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (
            int(row["service_location_id"]),
            int(row["timestamp"]),
            row["date"].to_pydatetime(),
            float(row["consumption_kwh"]),
            float(row["solar_kwh"]),
            float(row["grid_import_kwh"]),
            float(row["grid_export_kwh"]),
        ))

    conn.commit()
    conn.close()


# -----------------------
# Main
# -----------------------

def main():

    access_token = get_access_token()

    from_ts, to_ts = get_time_range()

    for service_location_id in [75452, 80060]:

        data = get_consumption(
            service_location_id=service_location_id,
            access_token=access_token,
            aggregation=1,  # 🔥 kwartierdata
            from_ts=from_ts,
            to_ts=to_ts,
        )

        df = transform_data(
            data,
            service_location_id
        )

        save_to_db(df)

        print(
            f"Inserted {len(df)} rows "
            f"for service location {service_location_id}"
        )


if __name__ == "__main__":
    main()