import cx_Oracle

# Replace these with your actual credentials and connection details
username = "your_username"
password = "your_password"
hostname = "your_hostname"
port = "your_port"
service_name = "your_service_name"

# Establish connection
conn = cx_Oracle.connect(
    user=username,
    password=password,
    dsn=f"{hostname}:{port}/{service_name}"
)
