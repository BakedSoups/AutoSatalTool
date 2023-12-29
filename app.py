from flask import Flask, render_template, send_file
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download')
def download():
    filename = 'output_presentation.pptx'
    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
