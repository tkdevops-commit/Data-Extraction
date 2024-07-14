from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def index():
    return send_from_directory('', 'index.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = request.json
    field1 = data['field1']
    field2 = data['field2']
    field3 = data['field3']

    # Check if the Excel file exists
    file_exists = os.path.isfile('data.xlsx')

    # Load existing data or create a new DataFrame
    if file_exists:
        df = pd.read_excel('data.xlsx')
    else:
        df = pd.DataFrame(columns=['Field 1', 'Field 2', 'Field 3'])

    # Create a new DataFrame with the new data
    new_data = pd.DataFrame({'Field 1': [field1], 'Field 2': [field2], 'Field 3': [field3]})

    # Append the new data to the existing DataFrame
    df = pd.concat([df, new_data], ignore_index=True)

    # Save the data back to the Excel file
    df.to_excel('data.xlsx', index=False)

    return jsonify({'message': 'Data saved successfully!'}), 200

if __name__ == '__main__':
    app.run(debug=True)