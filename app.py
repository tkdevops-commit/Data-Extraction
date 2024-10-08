from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os

#New route allowss users to download the file directly
@app.route('/download', methods=['GET'])
def download_file():
    path = 'data.xlsx'
    return send_from_directory('', path, as_attachment=True)

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

    file_exists = os.path.isfile('data.xlsx')

    if file_exists:
        df = pd.read_excel('data.xlsx')
    else:
        df = pd.DataFrame(columns=['Field 1', 'Field 2', 'Field 3'])

    new_data = pd.DataFrame({'Field 1': [field1], 'Field 2': [field2], 'Field 3': [field3]})

    df = pd.concat([df, new_data], ignore_index=True)

    df.to_excel('data.xlsx', index=False)

    return jsonify({'message': 'Data saved successfully!'}), 200

if __name__ == '__main__':
    app.run(debug=True)

#New route displays contents of file in table format

@app.route('/view_data', methods=['GET'])
def view_data():
    df = pd.read_excel('data.xlsx')
    return df.to_html()

#Allows eidt to previously submitted data
@app.route('/edit', methods=['POST'])
def edit_data():
    data = request.json
    record_id = data['id']
    field1 = data['field1']
    field2 = data['field2']
    field3 = data['field3']
    
    df = pd.read_excel('data.xlsx')
    df.loc[df['ID'] == record_id, ['Field 1', 'Field 2', 'Field 3']] = [field1, field2, field3]
    df.to_excel('data.xlsx', index=False)
    
    return jsonify({'message': 'Data updated successfully!'}), 200

#Allows users to delete specific records
@app.route('/delete', methods=['POST'])
def delete_data():
    data = request.json
    record_id = data['id']
    
    df = pd.read_excel('data.xlsx')
    df = df[df['ID'] != record_id]
    df.to_excel('data.xlsx', index=False)
    
    return jsonify({'message': 'Data deleted successfully!'}), 200

#Allows users to query excel file
@app.route('/search', methods=['GET'])
def search_data():
    query = request.args.get('query')
    df = pd.read_excel('data.xlsx')
    filtered_df = df[df['Field 1'].str.contains(query)]
    return filtered_df.to_html()
