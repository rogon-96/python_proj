from flask import Flask,render_template,url_for,request,redirect,flash
from flask_mail import Mail, Message
import csv,os,pandas as pd
from werkzeug.utils import secure_filename
import proj1
app = Flask(__name__)
app.config['MAIL_SERVER']='smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = 'motukurumidhun@gmail.com'
app.config['MAIL_PASSWORD'] = 'Midhunreddy134@'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)
bool_dict = {
    'master_response': '',
    'responses_response': '',
    'positive_response':'',
    'negative_response': '',
    'positive':'',
    'negative': ''
}
sample_input_path = "./sample_input"
sample_output_path = "./sample_output"
crt_marks,wrg_marks = 0,0

def handle_cases(request):
    if not os.path.exists(os.path.join(sample_input_path,"master_roll.csv")):
        bool_dict["master_response"] = "You have not Uploaded master_roll.csv"
        return False
    if not os.path.exists(os.path.join(sample_input_path,"responses.csv")):
        bool_dict["responses_response"] = "You have not Uploaded responses.csv"
        return False
    if request.form.get('positive')=='':
        bool_dict["positive_response"] = "This field is required"
        return False
    else :
        bool_dict["positive_response"] = ""
        bool_dict['positive'] = request.form.get('positive')
    if request.form.get('negative')=='':
        bool_dict["negative_response"] = "This field is required"
        return False
    else:
        bool_dict["negative_response"] = ""
        bool_dict['negative'] = request.form.get('negative')
    return True
    
def handle_file_save(FileObject,req_file_name,req_resp_name):
    if not FileObject.filename:
        bool_dict[f"{req_resp_name}"] = "Didn't upload any file."
        return
    file_name = FileObject.filename
    if file_name!=req_file_name:
        bool_dict[f"{req_resp_name}"] = "Uploaded a wrong file..plz Upload {}".format(req_file_name)
        return 
    sample_input= os.path.join(os.getcwd(),"sample_input")
    os.makedirs(sample_input,exist_ok = True)
    if os.path.exists(f"./sample_input/{file_name}"):
        os.remove(f"./sample_input/{file_name}")
    filename = secure_filename(file_name)
    FileObject.save(os.path.join(sample_input_path,filename))  
    bool_dict[f"{req_resp_name}"] = "Uploaded Successfully"
    return 

@app.route('/',methods=['GET'])
def index():
    return render_template('index.html',data = bool_dict)

@app.route('/master',methods=['GET','POST'])
def master():
    if request.method=="POST":
        FileObject = request.files.get("master")
        handle_file_save(FileObject,"master_roll.csv","master_response")
    return redirect(url_for('index'))

@app.route('/response',methods=['GET','POST'])
def response():
    FileObject = request.files.get("response")
    handle_file_save(FileObject,"responses.csv","responses_response")
    return redirect(url_for('index')) 


@app.route('/RollNo',methods=['GET','POST'])
def rollno():
    global crt_marks,wrg_marks
    if not handle_cases(request):
        return redirect(url_for('index')) 
    crt_marks = float(request.form.get('positive'))  
    wrg_marks = float(request.form.get('negative'))
    
    exists = proj1.generateMarksheet(crt_marks,wrg_marks)
    if exists!=None:
        bool_dict["rollno_response"] = exists
        return redirect(url_for('index'))
    bool_dict["rollno_response"] = "Generated Successfully"  
    return redirect(url_for('index'))


@app.route('/concise',methods=['GET','POST'])
def concise():
    global crt_marks,wrg_marks
    if not os.path.exists(os.path.join(sample_input_path,"master_roll.csv")):
        bool_dict["master_response"] = "You have not Uploaded master_roll.csv"
        return redirect(url_for('index'))
    if not os.path.exists(os.path.join(sample_input_path,"responses.csv")):
        bool_dict["responses_response"] = "You have not Uploaded responses.csv"
        return redirect(url_for('index'))
    if crt_marks==0 and wrg_marks==0:
        bool_dict["concise_response"] = "Plz fill the crt and wrng options"  
        return redirect(url_for('index'))
    exists = proj1.generate_concise(crt_marks,wrg_marks)
    if exists!=None:
        bool_dict["concise_response"] = exists
        return redirect(url_for('index'))
    bool_dict["concise_response"] = "Generated Successfully"
    return redirect(url_for('index'))


@app.route('/sendemail',methods=['GET','POST'])
def send_email():
    if not os.path.exists(os.path.join(sample_input_path,"master_roll.csv")):
        bool_dict["master_response"] = "You have not Uploaded master_roll.csv"
        return redirect(url_for('index'))
    if not os.path.exists(os.path.join(sample_input_path,"responses.csv")):
        bool_dict["responses_response"] = "You have not Uploaded responses.csv"
        return redirect(url_for('index'))
    if crt_marks==0 and wrg_marks==0:
        bool_dict["concise_response"] = "Plz fill the crt and wrng options"  
        return redirect(url_for('index'))   
    if not os.path.exists('sample_output\marksheet'):
        bool_dict['email_response'] = "Plz generate marksheets"  
        return redirect(url_for('index'))        
    files = os.listdir('sample_output\marksheet')
    roll_email_lst = proj1.get_roll_email()
    for roll,email1,email2 in roll_email_lst:
        msg = Message('Your marks',sender ='motukurumidhun@gmail.com',recipients = [email1,email2])
        msg.body = 'Your Quiz Marks'
        with app.open_resource("sample_output\marksheet\{}.xlsx".format(roll)) as fp:  
            msg.attach(f"{roll}", "application/xlsx", fp.read()) 
            mail.send(msg)
            print(f"Success mail sent {roll}")
    bool_dict["email_response"] = "Email Sent Successfully"
    return redirect(url_for('index')) 


if __name__ == "__main__":
    app.run(debug=True)
