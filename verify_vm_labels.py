import pandas as pd
from googleapiclient import discovery
from google.oauth2 import service_account
import os
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\aruni\OneDrive\Desktop\python-program\vibrant-mantis-408011-50808ba855fd.json"


# Google Cloud Project ID
PROJECT_ID = 'vibrant-mantis-408011'  # Replace with your project ID

# Path to your service account key file
SERVICE_ACCOUNT_FILE = r'C:\Users\aruni\OneDrive\Desktop\python-program\vibrant-mantis-408011-50808ba855fd.json'

# Initialize Google Compute Engine API client
def get_compute_service():
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE)
    service = discovery.build('compute', 'v1', credentials=credentials)
    return service

# Function to get VM labels from a specified zone
def get_vm_labels(service, instance_name, zone):
    try:
        instance = service.instances().get(project=PROJECT_ID, zone=zone, instance=instance_name).execute()
        return instance.get('labels', {})
    except Exception as e:
        print(f"Error retrieving labels for {instance_name} in zone {zone}: {e}")
        return None

# Load VM label data from Excel
def load_vm_data(file_path):
    df = pd.read_excel(file_path)
    return df

# Save verification results to a new Excel file
def save_results(df, output_path):
    df.to_excel(output_path, index=False)

# Main function to verify VM labels with zone-specific retrieval
def verify_vm_labels(input_excel, output_excel):
    # Load data from Excel
    df = load_vm_data(input_excel)
    
    # Initialize Google Cloud Compute service
    service = get_compute_service()

    # Add a results column to DataFrame
    df['status'] = ''

    # Check each VM's labels
    for index, row in df.iterrows():
        vm_name = row['instance_name']
        zone = row['zone']
        
        # Get actual labels from the VM
        actual_labels = get_vm_labels(service, vm_name, zone)

        if actual_labels is None:
            df.at[index, 'status'] = 'VM not found or access error'
        else:
            # Initialize status for this row
            status = []
            
            # Compare each column label with the actual label value
            for label_key in row.index[2:]:  # Skip 'instance_name' and 'zone'
                expected_value = row[label_key]
                
                # Skip if the cell is empty (NaN) as it's not relevant for this VM
                if pd.isna(expected_value):
                    continue
                
                # Compare the actual value for the given label key
                actual_value = actual_labels.get(label_key)
              #  print(f"Comparing {label_key}: expected = '{expected_value}', actual = '{actual_value}'")
                
                if str(actual_value).strip().lower() == str(expected_value).strip().lower():
                  status.append(f"{label_key}: Match")
                    
                else:
                   # status.append(f"{label_key}: Mismatch (expected: {expected_value}, actual: {actual_value})")
                 status.append(f"{label_key}: Mismatch")

            
            # Combine status messages for this VM
            df.at[index, 'status'] = "; ".join(status)

    # Save the verification results to a new Excel file
    save_results(df, output_excel)
    print(f"Verification results saved to {output_excel}")

if __name__ == "__main__":
    # Input and output Excel files
    input_excel = "vm_labels.xlsx"  # Input Excel file with VM data
    output_excel = "verification_results.xlsx"  # Output file to save results

    # Run label verification
    verify_vm_labels(input_excel, output_excel)
