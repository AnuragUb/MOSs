import requests

url = 'http://localhost:5000/api/recognize-audio'
file_path = 'processed_hin.mp4'  # Change to your local file if needed
tcr_in = '00:00:00.000'
tcr_out = '00:00:10.000'

with open(file_path, 'rb') as f:
    files = {'file': f}
    data = {'tcrIn': tcr_in, 'tcrOut': tcr_out}
    response = requests.post(url, files=files, data=data)
    print('Status code:', response.status_code)
    print('Response:', response.text) 