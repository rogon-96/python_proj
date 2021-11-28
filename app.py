from flask import Flask,render_template,url_for,request,redirect,flash
import csv,os
from flask_mail import Mail, Message
from werkzeug.utils import secure_filename
# from func import *
app = Flask(__name__)
app.config['MAIL_SERVER']='smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = 'YOUR_GMAIL_ID'
app.config['MAIL_PASSWORD'] = 'YOUR_GMAIL_PASSWORD'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)
bool_dict = {
    'master_response': '',
    'responses_response': ''
}
sample_input_path = "./sample_input"
sample_output_path = "./sample_output"


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
    print(os.getcwd())
    FileObject = request.files.get("response")
    handle_file_save(FileObject,"responses.csv","responses_response")
    return redirect(url_for('index')) 


@app.route('/RollNo',methods=['GET','POST'])
def rollno():
    bool_dict['positive_response'] = bool_dict['negative_response']=''
    if not os.path.exists(os.path.join(sample_input_path,"master_roll.csv")):
        bool_dict["rollno_response"] = "You have not Uploaded master_roll.csv"
        return redirect(url_for('index'))
    if not os.path.exists(os.path.join(sample_input_path,"responses.csv")):
        bool_dict["rollno_response"] = "You have not Uploaded responses.csv"
        return redirect(url_for('index'))
    if request.form.get('positive')=='':
        bool_dict["positive_response"] = "This field is required"
        return redirect(url_for('index'))
    if request.form.get('negative')=='':
        bool_dict["negative_response"] = "This field is required"
        return redirect(url_for('index'))
    #Python Code must be inserted here
    bool_dict["rollno_response"] = "Generated Successfully"
    return redirect(url_for('index'))


@app.route('/concise',methods=['GET','POST'])
def concise():
    #check if the files are existed then
    #Python Code must be inserted here
    bool_dict["concise_response"] = "Generated Successfully"
    return redirect(url_for('index'))

@app.route('/sendemail',methods=['GET','POST'])
def send_email():
    files = os.listdir('sample_output\marksheet')
    for file in files:
        msg = Message('Hello',sender ='motukurumidhun@gmail.com',recipients = ['venkat.abhilashreddy@gmail.com'])
        msg.body = 'Hello Flask message sent from Flask-Mail'
        with app.open_resource(f"./sample_output/marksheet/{file}") as fp:  
            msg.attach(f"{file}", "application/xlsx", fp.read()) 
            mail.send(msg)
            print(f"Success mail sent{file}")
    bool_dict["email_response"] = "Email Sent Successfully"
    return redirect(url_for('index')) 
if __name__ == "__main__":
    app.run(debug=True)