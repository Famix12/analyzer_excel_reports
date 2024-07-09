# import os
# from flask import Flask, request, redirect, url_for, render_template, send_file
# import pandas as pd
# import matplotlib.pyplot as plt
# from fpdf import FPDF
# import arabic_reshaper
# from bidi.algorithm import get_display
# from werkzeug.utils import secure_filename

# app = Flask(__name__)
# app.config['UPLOAD_FOLDER'] = './uploads'
# app.config['REPORT_FOLDER'] = './reports'
# app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

# # Create directories if they do not exist
# os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
# os.makedirs(app.config['REPORT_FOLDER'], exist_ok=True)

# def allowed_file(filename):
#     return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# def reshape_text(text):
#     reshaped_text = arabic_reshaper.reshape(text)
#     bidi_text = get_display(reshaped_text)
#     return bidi_text

# def generate_report(filepath, reportpath):
#     # Load the Excel file
#     df = pd.read_excel(filepath, sheet_name='Sheet1')

#     # Calculate the required data
#     count_of_cars = df.shape[0]

#     # Convert date column to datetime
#     df['التاريخ'] = pd.to_datetime(df['التاريخ'])
#     time_interval = df['التاريخ'].max() - df['التاريخ'].min()
#     time_interval_days = time_interval.days  # Extract only the number of days

#     # Analyze specified columns for the top consumed pieces
#     columns_to_analyze = [
#         'حالة المركبة [المكينة]', 'حالة المركبة [السيور]', 'حالة المركبة [الزيوت]', 
#         'حالة المركبة [البكرات]', 'حالة المركبة [ريديتر]', 'حالة المركبة [الكمبرسور]', 
#         'حالة المركبة [الكلتش]', 'حالة المركبة [البطارية]', 'حالة المركبة [مساعدات أمامي]', 
#         'حالة المركبة [القير]', 'حالة المركبة [هاند بريك]', 'حالة المركبة [البريك]', 
#         'حالة المركبة [الفحمات]', 'حالة المركبة [جلب]', 'حالة المركبة [مساعدات خلفي]', 
#         'حالة المركبة [اذرعه]', 'حالة المركبة [رمان امامي]', 'حالة المركبة [رمان خلفي]', 
#         'حالة المركبة [سست]', 'حالة المركبة [زجاج السيارة]', 'حالة المركبة [صدامات]', 
#         'حالة المركبة [البدي]', 'حالة المركبة [الاطارات + الاحتياطي]'
#     ]

#     df_analyze = df[columns_to_analyze].replace('سليم', pd.NA)
#     consumption_counts = df_analyze.notna().sum()
#     top_three_consumed = consumption_counts.sort_values(ascending=False).head(3)
#     top_ten_consumed = consumption_counts.sort_values(ascending=False).head(10)

#     # Set font properties for Matplotlib
#     from matplotlib import rcParams
#     rcParams['font.family'] = 'DejaVu Sans'
#     rcParams['font.sans-serif'] = ['DejaVu Sans']

#     # Plot top 10 consumed pieces
#     plt.figure(figsize=(10, 6))
#     top_ten_consumed.plot(kind='bar')
#     plt.title(reshape_text('أكثر عشر قطع استهلاكاً'))
#     plt.xlabel(reshape_text('القطع'))
#     plt.ylabel(reshape_text('العدد'))
#     plt.xticks(rotation=45, ha='right')
#     plt.tight_layout()
#     plt.savefig('top_ten_consumed.png')

#     # Plot monthly car registrations
#     df['Month_Year'] = df['التاريخ'].dt.to_period('M')
#     monthly_registrations = df['Month_Year'].value_counts().sort_index()

#     plt.figure(figsize=(10, 6))
#     monthly_registrations.plot(kind='line')
#     plt.title(reshape_text('تسجيلات السيارات الشهرية'))
#     plt.xlabel(reshape_text('الشهر والسنة'))
#     plt.ylabel(reshape_text('عدد التسجيلات'))
#     plt.tight_layout()
#     plt.savefig('monthly_registrations.png')

#     # Plot mileage distribution using a bar graph
#     plt.figure(figsize=(10, 6))
#     df['عداد الكيلوات'].dropna().plot(kind='hist', bins=20, edgecolor='black')
#     plt.title(reshape_text('توزيع عداد الكيلوات'))
#     plt.xlabel(reshape_text('عداد الكيلوات'))
#     plt.ylabel(reshape_text('عدد السيارات'))
#     plt.tight_layout()
#     plt.savefig('mileage_distribution.png')

#     # Create PDF
#     class PDF(FPDF):
#         def header(self):
#             self.set_font('DejaVu', 'B', 12)
#             self.cell(0, 10, reshape_text('تقرير السيارة'), 0, 1, 'C')

#         def chapter_title(self, title):
#             self.set_font('DejaVu', 'B', 12)
#             self.cell(0, 10, reshape_text(title), 0, 1, 'L')
#             self.ln(10)

#         def chapter_body(self, body):
#             self.set_font('DejaVu', '', 12)
#             self.multi_cell(0, 10, reshape_text(body))
#             self.ln()

#         def add_image(self, image_path, title):
#             self.add_page()
#             self.chapter_title(title)
#             self.image(image_path, x=10, y=30, w=180)
#             self.ln(85)  # Adjust the space after the image

#     # Instantiate the PDF
#     pdf = PDF()

#     # Add the DejaVu font
#     pdf.add_font('DejaVu', '', './static/fonts/DejaVuSans.ttf', uni=True)
#     pdf.add_font('DejaVu', 'B', './static/fonts/DejaVuSans-Bold.ttf', uni=True)

#     # Create the PDF content
#     pdf.add_page()

#     # Add count of cars
#     pdf.chapter_title('عدد السيارات')
#     pdf.chapter_body(f'إجمالي عدد السيارات: {count_of_cars}')

#     # Add time interval
#     pdf.chapter_title('الفترة الزمنية')
#     pdf.chapter_body(f'الفترة الزمنية بين السيارة الأولى والأخيرة المسجلة: {time_interval_days} days')

#     # Add top three most consumed pieces
#     pdf.chapter_title('أكثر ثلاث قطع استهلاكاً')
#     for piece, count in top_three_consumed.items():
#         pdf.chapter_body(f'{piece}: {count} occurrences')

#     # Add graph for top 10 consumed pieces
#     pdf.add_image('top_ten_consumed.png', 'أكثر عشر قطع استهلاكاً')

#     # Add graph for monthly car registrations
#     pdf.add_image('monthly_registrations.png', 'تسجيلات السيارات الشهرية')

#     # Add graph for mileage distribution
#     pdf.add_image('mileage_distribution.png', 'توزيع عداد الكيلوات')

#     # Output the PDF
#     pdf.output(reportpath)

#     # Clean up the images
#     os.remove('top_ten_consumed.png')
#     os.remove('monthly_registrations.png')
#     os.remove('mileage_distribution.png')

# @app.route('/', methods=['GET', 'POST'])
# def upload_file():
#     if request.method == 'POST':
#         # Check if the post request has the file part
#         if 'file' not in request.files:
#             return redirect(request.url)
#         file = request.files['file']
#         # If the user does not select a file, the browser submits an empty file without a filename
#         if file.filename == '':
#             return redirect(request.url)
#         if file and allowed_file(file.filename):
#             filename = secure_filename(file.filename)
#             filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
#             file.save(filepath)
#             reportpath = os.path.join(app.config['REPORT_FOLDER'], 'report_with_graphs.pdf')
#             generate_report(filepath, reportpath)
#             return send_file(reportpath, as_attachment=True)
#     return render_template('upload.html')

# if __name__ == '__main__':
#     app.run(debug=True)


import os
from flask import Flask, request, redirect, url_for, render_template, send_file
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import arabic_reshaper
from bidi.algorithm import get_display
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = './uploads'
app.config['REPORT_FOLDER'] = './reports'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

# Create directories if they do not exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORT_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def reshape_text(text):
    reshaped_text = arabic_reshaper.reshape(text)
    bidi_text = get_display(reshaped_text)
    return bidi_text

def generate_report(filepath, reportpath):
    # Load the Excel file
    df = pd.read_excel(filepath, sheet_name='Sheet1')

    # Calculate the required data
    count_of_cars = df.shape[0]

    # Convert date column to datetime
    df['التاريخ'] = pd.to_datetime(df['التاريخ'])
    time_interval = df['التاريخ'].max() - df['التاريخ'].min()
    time_interval_days = time_interval.days  # Extract only the number of days

    # Analyze specified columns for the top consumed pieces
    columns_to_analyze = [
        'حالة المركبة [المكينة]', 'حالة المركبة [السيور]', 'حالة المركبة [الزيوت]', 
        'حالة المركبة [البكرات]', 'حالة المركبة [ريديتر]', 'حالة المركبة [الكمبرسور]', 
        'حالة المركبة [الكلتش]', 'حالة المركبة [البطارية]', 'حالة المركبة [مساعدات أمامي]', 
        'حالة المركبة [القير]', 'حالة المركبة [هاند بريك]', 'حالة المركبة [البريك]', 
        'حالة المركبة [الفحمات]', 'حالة المركبة [جلب]', 'حالة المركبة [مساعدات خلفي]', 
        'حالة المركبة [اذرعه]', 'حالة المركبة [رمان امامي]', 'حالة المركبة [رمان خلفي]', 
        'حالة المركبة [سست]', 'حالة المركبة [زجاج السيارة]', 'حالة المركبة [صدامات]', 
        'حالة المركبة [البدي]', 'حالة المركبة [الاطارات + الاحتياطي]'
    ]

    df_analyze = df[columns_to_analyze].replace('سليم', pd.NA)
    consumption_counts = df_analyze.notna().sum()
    top_three_consumed = consumption_counts.sort_values(ascending=False).head(3)
    top_ten_consumed = consumption_counts.sort_values(ascending=False).head(10)

    # Set font properties for Matplotlib
    from matplotlib import rcParams
    rcParams['font.family'] = 'DejaVu Sans'
    rcParams['font.sans-serif'] = ['DejaVu Sans']

    # Plot top 10 consumed pieces
    plt.figure(figsize=(10, 6))
    top_ten_consumed.plot(kind='bar')
    plt.title(reshape_text('أكثر عشر قطع استهلاكاً'))
    plt.xlabel(reshape_text('القطع'))
    plt.ylabel(reshape_text('العدد'))
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig('top_ten_consumed.png')

    # Plot monthly car registrations
    df['Month_Year'] = df['التاريخ'].dt.to_period('M')
    monthly_registrations = df['Month_Year'].value_counts().sort_index()

    plt.figure(figsize=(10, 6))
    monthly_registrations.plot(kind='line')
    plt.title(reshape_text('تسجيلات السيارات الشهرية'))
    plt.xlabel(reshape_text('الشهر والسنة'))
    plt.ylabel(reshape_text('عدد التسجيلات'))
    plt.tight_layout()
    plt.savefig('monthly_registrations.png')

    # Plot mileage distribution using a bar graph
    plt.figure(figsize=(10, 6))
    df['عداد الكيلوات'].dropna().plot(kind='hist', bins=20, edgecolor='black')
    plt.title(reshape_text('توزيع عداد الكيلوات'))
    plt.xlabel(reshape_text('عداد الكيلوات'))
    plt.ylabel(reshape_text('عدد السيارات'))
    plt.tight_layout()
    plt.savefig('mileage_distribution.png')

    # Create PDF
    class PDF(FPDF):
        def header(self):
            self.set_font('DejaVu', 'B', 12)
            self.cell(0, 10, reshape_text('تقرير السيارة'), 0, 1, 'C')

        def chapter_title(self, title):
            self.set_font('DejaVu', 'B', 12)
            self.cell(0, 10, reshape_text(title), 0, 1, 'L')
            self.ln(10)

        def chapter_body(self, body):
            self.set_font('DejaVu', '', 12)
            self.multi_cell(0, 10, reshape_text(body))
            self.ln()

        def add_image(self, image_path, title):
            self.add_page()
            self.chapter_title(title)
            self.image(image_path, x=10, y=30, w=180)
            self.ln(85)  # Adjust the space after the image

    # Instantiate the PDF
    pdf = PDF()

    # Add the DejaVu font
    pdf.add_font('DejaVu', '', './static/fonts/DejaVuSans.ttf', uni=True)
    pdf.add_font('DejaVu', 'B', './static/fonts/DejaVuSans-Bold.ttf', uni=True)

    # Create the PDF content
    pdf.add_page()

    # Add count of cars
    pdf.chapter_title('عدد السيارات')
    pdf.chapter_body(f'إجمالي عدد السيارات: {count_of_cars}')

    # Add time interval
    pdf.chapter_title('الفترة الزمنية')
    pdf.chapter_body(f'الفترة الزمنية بين السيارة الأولى والأخيرة المسجلة: {time_interval_days} days')

    # Add top three most consumed pieces
    pdf.chapter_title('أكثر ثلاث قطع استهلاكاً')
    for piece, count in top_three_consumed.items():
        pdf.chapter_body(f'{piece}: {count} occurrences')

    # Add graph for top 10 consumed pieces
    pdf.add_image('top_ten_consumed.png', 'أكثر عشر قطع استهلاكاً')

    # Add graph for monthly car registrations
    pdf.add_image('monthly_registrations.png', 'تسجيلات السيارات الشهرية')

    # Add graph for mileage distribution
    pdf.add_image('mileage_distribution.png', 'توزيع عداد الكيلوات')

    # Output the PDF
    pdf.output(reportpath)

    # Clean up the images
    os.remove('top_ten_consumed.png')
    os.remove('monthly_registrations.png')
    os.remove('mileage_distribution.png')

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an empty file without a filename
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            reportpath = os.path.join(app.config['REPORT_FOLDER'], 'report_with_graphs.pdf')
            generate_report(filepath, reportpath)
            return send_file(reportpath, as_attachment=True)
    return render_template('upload.html')

if __name__ != '__main__':
    app.run(debug=True)
