import os
from flask import Flask, request, render_template, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

# Ensure the temporary folder exists
UPLOAD_FOLDER = './uploads/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Function to compare two files and generate a single result sheet
def compare_files(master_file_path, input_file_path):
    # Load both files into pandas DataFrames
    master_df = pd.read_excel(master_file_path)
    input_df = pd.read_excel(input_file_path)

    # Convert 'Sales Date' to datetime format for both master and input files
    master_df['Sales Date'] = pd.to_datetime(master_df['Sales Date'], errors='coerce')
    input_df['Sales Date'] = pd.to_datetime(input_df['Sales Date'], errors='coerce')
    
    # Format the sales date in MM/DD/YYYY format
    master_df['Sales Date'] = master_df['Sales Date'].dt.strftime('%m/%d/%Y')
    input_df['Sales Date'] = input_df['Sales Date'].dt.strftime('%m/%d/%Y')

    # Merge the dataframes on 'Sheriff #' and 'County Name' to identify matches
    comparison_df = pd.merge(master_df[['County Name', 'Sheriff #', 'Sales Date']], 
                             input_df[['County Name', 'Sheriff #', 'Sales Date']], 
                             on=['County Name', 'Sheriff #'], 
                             how='outer', indicator=True, suffixes=('_master', '_input'))

    # Add a new 'Result' column to specify the comparison result
    comparison_df['Result'] = ''
    
    # Mark the results based on the comparison
    comparison_df.loc[comparison_df['_merge'] == 'left_only', 'Result'] = 'Not in the System'
    comparison_df.loc[comparison_df['_merge'] == 'right_only', 'Result'] = 'Newly Added'
    comparison_df.loc[(comparison_df['_merge'] == 'both') & 
                      (comparison_df['Sales Date_master'] != comparison_df['Sales Date_input']), 'Result'] = 'Date Changed'

    # Select the relevant columns to keep
    final_result_df = comparison_df[['County Name', 'Sheriff #', 'Sales Date_master', 'Sales Date_input', 'Result']]

    # Filter the rows where 'Result' has a value (i.e., not empty)
    final_result_df = final_result_df[final_result_df['Result'] != '']

    # Save the result to a new Excel file in memory (in-memory file object)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_result_df.to_excel(writer, sheet_name='Comparison Results', index=False)
    output.seek(0)

    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    # Get uploaded files
    master_file = request.files['master_file']
    input_file = request.files['input_file']
    
    if master_file and input_file:
        # Save the uploaded files temporarily
        master_file_path = os.path.join(UPLOAD_FOLDER, 'master_file.xlsx')
        input_file_path = os.path.join(UPLOAD_FOLDER, 'input_file.xlsx')

        master_file.save(master_file_path)
        input_file.save(input_file_path)

        # Compare the files and get the result in memory
        result_file = compare_files(master_file_path, input_file_path)

        # Send the result file back to the user
        return send_file(result_file, as_attachment=True, download_name="comparison_results.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True)
