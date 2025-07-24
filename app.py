from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
from docx.shared import Inches
import os
import tempfile
from werkzeug.utils import secure_filename
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def format_date(date_string):
    if date_string:
        try:
            date_obj = datetime.strptime(date_string, '%Y-%m-%d')
            return date_obj.strftime('%d/%m/%Y')
        except:
            return date_string
    return ''

def format_date_range(date_range_string):
    if date_range_string and ' - ' in date_range_string:
        try:
            start_date, end_date = date_range_string.split(' - ')
            formatted_start = format_date(start_date.strip())
            formatted_end = format_date(end_date.strip())
            return f"{formatted_start} - {formatted_end}"
        except:
            return date_range_string
    return date_range_string

def to_sentence_case(text):
    if text:
        return text.strip().capitalize()
    return text

def process_uploaded_images(request):
    uploaded_images = {}
    for i in range(1, 7):
        photo_key = f'photo{i}'
        if photo_key in request.files:
            file = request.files[photo_key]
            if file and file.filename != '' and allowed_file(file.filename):
                temp_dir = tempfile.gettempdir()
                filename = secure_filename(f"temp_photo_{i}_{file.filename}")
                filepath = os.path.join(temp_dir, filename)
                file.save(filepath)
                uploaded_images[photo_key] = filepath
            else:
                uploaded_images[photo_key] = None
        else:
            uploaded_images[photo_key] = None
    return uploaded_images

def create_image_placeholder(image_path, width_inches=2.5):
    if image_path and os.path.exists(image_path):
        from docxtpl import InlineImage
        return InlineImage(template, image_path, width=Inches(width_inches))
    return "[No Image]"

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate():
    global template
    
    uploaded_images = process_uploaded_images(request)
    
    context = {
        'dayssick': request.form.get('dayssick'),
        'claimdate': format_date(request.form['claimdate']),
        'date': format_date(request.form['date']),
        'tagnumber': request.form['tagnumber'],
        'cattletype': to_sentence_case(request.form['cattletype']),
        'ownername': to_sentence_case(request.form['ownername']),
        'ownercontact': request.form['ownercontact'],
        'location': to_sentence_case(request.form['location']),
        'taluka': to_sentence_case(request.form['taluka']),
        'lossdate': format_date(request.form['lossdate']),
        'losstime': request.form['losstime'],
        'invdate': format_date(request.form['invdate']),
        'policynumber': request.form['policynumber'],
        'policyperiod': format_date_range(request.form['policyperiod']),
        'intimdate': format_date(request.form['intimdate']),
        'insuredname': to_sentence_case(request.form['insuredname']),
        'loan': to_sentence_case(request.form['loan']),
        'disease': to_sentence_case(request.form['disease']),
        'gvc': request.form['gvc'],
        'taglocation': to_sentence_case(request.form['taglocation']),
        'cattlecolor': to_sentence_case(request.form['cattlecolor']),
        'lactation': request.form['lactation'],
        'pregnant': to_sentence_case(request.form['pregnant']),
        'cattletail': to_sentence_case(request.form['cattletail']),
        'milkey': to_sentence_case(request.form['milkey']),
        'specialmarks': to_sentence_case(request.form['specialmarks']),
        'treatment': to_sentence_case(request.form['treatment']),
        'deathtype': to_sentence_case(request.form['deathtype']),
        'visit': to_sentence_case(request.form['visit'])
    }

    template_path = os.path.join(os.getcwd(), 'template.docx')
    template = DocxTemplate(template_path)
    
    try:
        for i in range(1, 7):
            photo_key = f'photo{i}'
            image_path = uploaded_images.get(photo_key)
            if image_path:
                context[photo_key] = create_image_placeholder(image_path)
            else:
                context[photo_key] = "[No Image Uploaded]"
    except Exception as e:
        print(f"Error processing images: {e}")
        for i in range(1, 7):
            context[f'photo{i}'] = "[Image Processing Error]"

    template.render(context)

    location = context['location'].replace(' ', '-').upper()
    taluka = context['taluka'].replace(' ', '-').upper()
    ownername = context['ownername'].replace(' ', '-').upper()
    tag_number = context['tagnumber'][-6:]
    output_filename = f"{location}_{taluka}_{ownername}_{tag_number}.docx"
    output_path = os.path.join(tempfile.gettempdir(), output_filename)

    template.save(output_path)

    for image_path in uploaded_images.values():
        if image_path and os.path.exists(image_path):
            try:
                os.remove(image_path)
            except:
                pass

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))