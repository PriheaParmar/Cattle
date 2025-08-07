from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
import tempfile
from datetime import datetime

app = Flask(__name__)

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
            return f"from {formatted_start} to {formatted_end}"
        except:
            return date_range_string
    return date_range_string

def to_sentence_case(text):
    if text:
        return text.strip().title()
    return text

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate():
    # Using single date field for all date requirements
    single_date = format_date(request.form['date'])
    
    context = {
        'dayssick': request.form.get('dayssick'),
        'claimdate': request.form['claimdate'],
        'date': single_date,  # Single date field
        'intimdate': single_date,  # Same date for intimation
        'invdate': single_date,  # Same date for investigation
        'lossdate': format_date(request.form['lossdate']),  # Keep loss date separate as it's different
        'tagnumber': request.form['tagnumber'],
        'cattletype': to_sentence_case(request.form['cattletype']),
        'ownername': to_sentence_case(request.form['ownername']),
        'ownercontact': request.form['ownercontact'],
        'location': to_sentence_case(request.form['location']),
        'taluka': to_sentence_case(request.form['taluka']),
        'losstime': request.form['losstime'],
        'policynumber': request.form['policynumber'],
        'policyperiod': format_date_range(request.form['policyperiod']),
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

    template.render(context)

    location = context['location'].replace(' ', '-').upper()
    taluka = context['taluka'].replace(' ', '-').upper()
    ownername = context['ownername'].replace(' ', '-').upper()
    tag_number = context['tagnumber'][-6:]
    output_filename = f"{location}_{taluka}_{ownername}_{tag_number}.docx"
    output_path = os.path.join(tempfile.gettempdir(), output_filename)

    template.save(output_path)

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))