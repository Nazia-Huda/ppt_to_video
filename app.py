from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from ppt_to_mp4 import ppt2pdf, ppt_presenter
import os 
import tempfile 
import argparse 
from subprocess import call,check_call 
import glob 
import tqdm 
from pptx import Presentation 
from gtts import gTTS 
import pyttsx3 
engine = pyttsx3.init('dummy')
from pdf2image import convert_from_path, convert_from_bytes
from moviepy.editor import VideoFileClip, concatenate_videoclips 


import ffmpeg 
import traceback 
from PIL import Image
import PyPDF2
from pathlib import Path
import win32com.client



app = Flask(__name__)
UPLOAD_FOLDER = 'E:/ffff/ppt_files'
app.config['UPLOAD_FOLDER']= UPLOAD_FOLDER

@app.route('/upload')
def upload_file():
   return render_template('index.html')
	
@app.route('/uploader', methods = ['GET', 'POST'])
def uploadd_file():
   if request.method == 'POST':
    f = request.files['file']
    file = f.filename
    ppt2pdf(file) 
    ppt_presenter(file, file.rsplit(".", 1 )[ 0 ]+".pdf", 
    file.rsplit(".", 1 )[ 0 ]+".mp4")
    # file.save(file.rsplit(".", 1 )[ 0 ]+".mp4")
    # f.download(file.rsplit(".", 1 )[ 0 ]+".mp4") 
    f.save(secure_filename(f.filename))
    print(file)
    return 'file uploaded successfully'



		
if __name__ == '__main__':
   app.run(debug = True)