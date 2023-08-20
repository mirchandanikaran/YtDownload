from flask import Flask, render_template, request, send_from_directory
from pytube import YouTube
import openpyxl
import os
import tempfile

app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'downloads'

def download_video(url, output_path):
    try:
        yt = YouTube(url)
        video_stream = yt.streams.get_highest_resolution()
        video_stream.download(output_path)
        return yt.title
    except Exception as e:
        return f"An error occurred: {e}"

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        excel_file = request.files['excel_file']

        # Save the uploaded Excel file to a temporary location
        excel_path = os.path.join(tempfile.gettempdir(), excel_file.filename)
        excel_file.save(excel_path)

        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active

        downloaded_videos = []

        # Create a temporary directory to save downloaded videos
        download_dir = tempfile.mkdtemp(prefix='yt_downloads_')

        for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True):
            url = row[0]
            if url:
                title = download_video(url, download_dir)
                downloaded_videos.append(os.path.join(download_dir, title + ".mp4"))

        # Clean up: Delete the temporary Excel file
        os.remove(excel_path)

        return render_template('index.html', message="Downloads completed!", download_links=downloaded_videos)

    return render_template('index.html', message="", download_links=None)

@app.route('/downloads/<filename>')
def downloads(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
