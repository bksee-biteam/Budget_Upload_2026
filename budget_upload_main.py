

from pathlib import Path
import pandas as pd
import os
import shutil
from sqlalchemy import create_engine, Table, MetaData, text
from azure.storage.blob import BlobServiceClient
import logging
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import pyodbc

import urllib



##---------------- Azure Blob Setup ----------------##
connection_string = 'DefaultEndpointsProtocol=https;AccountName=XXXXX;AccountKey=XXXXX;EndpointSuffix=core.windows.net'

local_download_path = r"C:\\Tasks_APP\\65_Budget_Upload_2026\\DownloadedFiles"


base_path = Path(local_download_path).parent
logs_dir = base_path / "Logs"
logs_dir.mkdir(parents=True, exist_ok=True)


app_status =  1
container_name = "aproved"
target_container_name = "processed"

number_of_processed_files = 0

recipient_email = 'support.bksee@andea.com'

def send_email(subject, message, to_email):
    from_email = "XXXX"  # Your email address
    email_password = "XXXX"  # Your email password
    smtp_server = "smtp.office365.com"  # Your SMTP server
    smtp_port = 587  # Port for your SMTP server
    to_email = 'XXXXX'  # Recipient's email address

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(from_email, email_password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")






try:

    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    container_client = blob_service_client.get_container_client(container_name)

    target_container_client  = blob_service_client.get_container_client(target_container_name)
 
    os.makedirs(local_download_path, exist_ok=True)

    # List and sort blobs by last modified time (newest first)
    blobs = sorted(container_client.list_blobs(), key=lambda b: b.last_modified, reverse=True)

# Take the 1000 newest blobs
    latest_blobs = blobs[:1000]

    downloaded_files = []
    for blob in latest_blobs:
        blob_client = container_client.get_blob_client(blob.name)
   
        download_file_path = os.path.join(local_download_path, os.path.basename(blob.name))
    
        with open(download_file_path, "wb") as f:
            f.write(blob_client.download_blob().readall())
        downloaded_files.append(download_file_path)

##---------------- DB Setup ----------------##
    # destination_server = 'KRK-L429' 
    # destination_database = 'BKSEEDB'
    # destination_username = 'test'
    # destination_password = 'test'
    # destination_connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={destination_server};DATABASE={destination_database};UID={destination_username};PWD={destination_password}; Trusted_Connection={{yes}};'


    params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=XXXXX"
    "DATABASE=BKSEEDB;"
    "UID=XXXXX;"
    "PWD=XXXX"
    )

    engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")


    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"Budget_Upload_2026_{timestamp}.log"

    logging.basicConfig(
        filename= logs_dir / log_filename,
        level=logging.INFO,# log level: INFO, WARNING, ERROR, DEBUG
        format="%(asctime)s - %(levelname)s - %(message)s",  # timestamp + level + message
        datefmt="%Y-%m-%d %H:%M:%S"
    )


    columns_to_insert = [
                        "Country",
                        "RestaurantCode",
                        "RestaurantName",
                        "SubBrand",
                        "Date",
                        "TotalSales",
                        "TotalTickets"
                        ]






    logging.info("Budget Upload 2025 started")

    for file_path in downloaded_files:
        if not file_path.endswith('.xlsx'):
            continue

        try:
            file_name = os.path.basename(file_path)


            logging.info("Started to process file \" %s\"", file_name)
           

            with pd.ExcelFile(file_path) as xls:

                found_sheet_name = next(
                        s for s in xls.sheet_names
                        if s.strip() == 'budget upload'
                        )




            df = pd.read_excel(file_path, sheet_name= found_sheet_name, header=0)
            ##df = pd.read_excel(file_path, sheet_name='budget upload', header=0)


  




            df = df.dropna(subset=["RestaurantCode", "Date"]) # Remove empty rows from dataframe


            df = df[columns_to_insert]

            df["RestaurantCode"] = pd.to_numeric(df["RestaurantCode"]).astype("Int64")
            #df['Date'] = pd.to_datetime(df['Date'].astype('Int64').astype(str),format='%Y%m%d')
            
            #df['Date'] = pd.to_datetime(df['Date'].astype(str).str.replace('-', '', regex=False).astype('Int64').astype(str),format='%Y%m%d')
           
           #df['Date'] = pd.to_datetime(df['Date'].astype(str).str.replace('-', '', regex=False).str.replace('.0', '', regex=False)   # <-- FIX.astype('Int64').astype(str),format='%Y%m%d')
           
            df['Date'] = pd.to_datetime(df['Date'].astype(str).str.replace('-', '', regex=False).str.split('.').str[0],format='%Y%m%d')

           
           
           
           
           
           
            df["T_FileName"] = file_name

            
            df["SubBrand"] = df["SubBrand"].replace("NSEE", "NS")


            

            logging.info("Starting to reload data in database.")
            with engine.begin() as conn:
                    conn.execute(text("TRUNCATE TABLE dbo.STAGING_BUDGET_UPLOAD_2026"))

                    df.to_sql(
                    "STAGING_BUDGET_UPLOAD_2026",
                    schema="dbo",
                    con=conn,  # Keep insert in the same transaction
                    if_exists="append",
                    index=False,
                    #chunksize=2000,
                    #method="multi"
                    )

                    conn.execute(text("EXEC dbo.sp_BUDGET_2026_MOVE_DATA_FROM_STAGING"))

            logging.info("Data from file \" %s\" was loaded to database", file_name)

            number_of_processed_files += 1




            logging.info("Started to move file \" %s\" to the Processed container", file_name)
        
            base_name, ext = os.path.splitext(os.path.basename(file_path))
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            blob_file_name = f"{base_name}_Processed_{timestamp}{ext}"

            with open(file_path, "rb") as data:
                target_container_client.upload_blob(
                name=blob_file_name, 
                data=data, 
                overwrite=False
                )
            
            logging.info("File \" %s\" was uploaded to Processed container", file_name)



            container_client.delete_blob(file_name)
            logging.info("File \" %s\" was deleted from Aproved container", file_name)


            # Move to archive
            archive_dir = os.path.join(local_download_path, 'archive')
            os.makedirs(archive_dir, exist_ok=True)


            shutil.move(file_path, os.path.join(archive_dir, blob_file_name))

        except Exception as e:
            app_status  = -1
            logging.error(f"Error processing file {file_path}: {e}", exc_info=True)


    logging.info(f"")   
    logging.info(f"Number of processed files:{number_of_processed_files}")

    if app_status < 0:
         raise Exception("Something went wrong with processing budget data.")



    logging.info("Application finished successfully.")
    send_email("[App_succeeded]Budget_Upload_2026_File_Upload",f'\n Number of processed files:{number_of_processed_files}',recipient_email)
       
except Exception as e:
         logging.error(f"Error: {e}")
         send_email("[App_failed]Budget_Upload_2026_File_Upload",f'Failed to load budget data.Check logs: {logs_dir} \nAn error occurred: {str(e)}',recipient_email)
