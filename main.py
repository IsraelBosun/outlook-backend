from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from typing import List
import pandas as pd
import os
import re
import extract_msg

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Dummy function to extract details from a .msg file
def extract_details_from_msg(file_path):
    try:
        msg = extract_msg.Message(file_path)
        body = msg.body

        # Extract dates using regex
        date_pattern = r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b'
        dates = re.findall(date_pattern, body)

        # Assign the dates if found
        sent_date = dates[0] if len(dates) > 0 else None
        effective_date = dates[1] if len(dates) > 1 else None

        # Extract other details
        sender = msg.sender
        receiver = msg.to

        # Extract email address from receiver
        email_pattern = r'<([^>]+)>'
        email_match = re.search(email_pattern, receiver)
        receiver_email = email_match.group(1) if email_match else receiver.strip()

        # Extract receiver's name
        receiver_name_match = re.search(r'Attention:\s*(.*)\n', body)
        receiver_name = receiver_name_match.group(1).strip() if receiver_name_match else None

        # Extract MPR and lending rates
        percentage_pattern = r'\b\d+(?:\.\d+)?%'
        all_percentages = re.findall(percentage_pattern, body)

        # Filter out percentages that are part of URLs or encoded sequences
        filtered_percentages = [
            match for match in all_percentages if not re.search(r'\bdata=|com%', match)
        ]

        # Debug: Print all filtered percentages found
        print("Filtered percentages found:", filtered_percentages)

        # Extract the third and fourth percentages as lending rates
        lending_old, lending_new = (filtered_percentages[2], filtered_percentages[3]) if len(filtered_percentages) >= 4 else (None, None)

        # Convert percentages to decimal
        lending_old = float(lending_old.strip('%')) / 100 if lending_old else None
        lending_new = float(lending_new.strip('%')) / 100 if lending_new else None

        # Formatted output
        details = {
            "Receiver": receiver_email,
            "Receiver Name": receiver_name,
            "Lending Rate Old": lending_old,
            "Lending Rate New": lending_new,
        }

        return details

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

@app.post('/cross_reference/')
async def cross_reference(email_files: List[UploadFile] = File(...), excel_file: UploadFile = File(...)):
    try:
        # Process the Excel file
        excel_contents = await excel_file.read()
        excel_path = f"/tmp/{excel_file.filename}"
        with open(excel_path, 'wb') as f:
            f.write(excel_contents)

        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path)
        else:
            raise HTTPException(status_code=400, detail="Excel file not found after upload.")

        # Process the email files
        all_details = {}
        for file in email_files:
            contents = await file.read()
            file_path = f"/tmp/{file.filename}"
            with open(file_path, 'wb') as f:
                f.write(contents)
            details = extract_details_from_msg(file_path)
            os.remove(file_path)  # Clean up the temporary file after processing
            if details:
                receiver_email = details['Receiver']
                if receiver_email not in all_details:
                    all_details[receiver_email] = []
                all_details[receiver_email].append(details)

        # Cross-check details with the Excel data
        results = []
        customer_counter = 1
        for email in df['Email']:
            if email in all_details:
                excel_data = df[df['Email'] == email].iloc[0]
                msg_data = all_details[email][0]  # Assuming there's only one relevant email per receiver

                if (excel_data['Receiver name'] == msg_data['Receiver Name'] and
                    excel_data['Old rate'] == float(msg_data['Lending Rate Old']) and
                    excel_data['New Rate'] == float(msg_data['Lending Rate New'])):

                    results.append({
                        "Customer": customer_counter,
                        "Email": email,
                        "Message": "Details are the same",
                        "Excel Data": {
                            "Receiver Name": excel_data['Receiver name'],
                            "Old Rate": excel_data['Old rate'],
                            "New Rate": excel_data['New Rate']
                        },
                        "Email Data": msg_data
                    })
                else:
                    differences = {}
                    if excel_data['Receiver name'] != msg_data['Receiver Name']:
                        differences["Name"] = f"Excel({excel_data['Receiver name']}) vs Email({msg_data['Receiver Name']})"
                    if excel_data['Old rate'] != float(msg_data['Lending Rate Old']):
                        differences["Old Rate"] = f"Excel({excel_data['Old rate']}) vs Email({msg_data['Lending Rate Old']})"
                    if excel_data['New Rate'] != float(msg_data['Lending Rate New']):
                        differences["New Rate"] = f"Excel({excel_data['New Rate']}) vs Email({msg_data['Lending Rate New']})"

                    results.append({
                        "Customer": customer_counter,
                        "Email": email,
                        "Differences": differences
                    })
            else:
                results.append({
                    "Customer": customer_counter,
                    "Email": email,
                    "Differences": "Customer with this email not found in Email data."
                })
            customer_counter += 1

        return results

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='0.0.0.0', port=8080)

