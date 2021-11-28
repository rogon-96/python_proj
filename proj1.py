import os,math,csv,openpyxl,pandas as pd

slist = pd.read_csv('./sample_input/master_roll.csv')
response = pd.read_csv('./sample_input/responses.csv')
sample_output_path = "./sample_output/marksheet"
def get_roll_email():
    lst = []
    for index,row in response.iterrows():
        lst.append([row["Roll Number"],row["Email address"],row["IITP webmail"]])
    return lst
def get_answer():
    for index,row in response.iterrows():
        if row['Roll Number'] == 'ANSWER':
            return [key for key in row][7:]
    return []

def calculate(mrks,wmrks,crt_opts_list,curr_opts):
    crt_opts,not_attempted,wrg_opts = 0,0,0
    for j in range(len(crt_opts_list)):
        if(curr_opts[j]=='nan'): not_attempted+=1
        elif(curr_opts[j] == crt_opts_list[j]): crt_opts+=1
        else : wrg_opts+=1
    result = crt_opts*mrks + wrg_opts*wmrks
    return result,[crt_opts,not_attempted,wrg_opts]

def generate_concise(mrks,wmrks):
    concise_marksheet = response
    crt_opts_list = get_answer()
    if crt_opts_list == []:
        return "Error!!!Answer is not exist in responses"
    last_list,score_af_neg= [],[]
    for index,row in response.iterrows():
        curr_opts = [key for key in row][7:]
        result,lst = calculate(mrks,wmrks,crt_opts_list,curr_opts)
        last_list.append(lst)
        score_af_neg.append(str(result)+"/"+str(mrks*len(crt_opts_list)))
    concise_marksheet.insert(loc =6,column ="Score_After_Negative",value =score_af_neg)
    concise_marksheet["Options"] = last_list
    os.makedirs("sample_output",exist_ok = True)
    os.makedirs(sample_output_path,exist_ok = True)
    concise_marksheet.to_csv(sample_output_path+"/concise_marksheet.csv", index=False)
    return  
   
def generateMarksheet(mrks,wmrks):
    crt = get_answer()
    if crt == []:
        return "Error!!!Answer is not exist in responses"
    os.makedirs("sample_output",exist_ok = True)
    os.makedirs(sample_output_path,exist_ok = True)
    for i in range(len(response)):       
        wb =  openpyxl.Workbook() #creation and initialising a workbook
        sheet = wb.active
        sheet.title = response.at[i,'Roll Number']
        sheet.merge_cells('A1:E4')
        # add image here
        sheet.merge_cells('A5:E5')
        sheet.cell(row=5,column=1).value = "Mark Sheet"
        sheet.cell(row=6,column=1).value = "Name:"
        sheet.merge_cells('B6:C6')
        sheet.cell(row=6,column=2).value = slist.at[i,'name']
        sheet.cell(row=6,column=4).value = "Exam:"
        sheet.cell(row=6,column=5).value = "quiz"
        sheet.cell(row=7,column=1).value = "Roll Number:"
        sheet.cell(row=7,column=2).value = response.at[i,'Roll Number']

        sheet.cell(row=9,column=2).value = "Right"
        sheet.cell(row=9,column=3).value = "Wrong"
        sheet.cell(row=9,column=4).value = "Not Attempt"
        sheet.cell(row=9,column=5).value = "Max"
        sheet.cell(row=10,column=1).value = "No."
        sheet.cell(row=11,column=1).value = "Marking"
        sheet.cell(row=12,column=1).value = "Total"

        r,c = 16,1
        rgt,mrks,wrg,wmrks,na = 0,5,0,-1,0
        sheet.cell(row=15,column=1).value = "Student Ans"
        sheet.cell(row=15,column=2).value = "Correct Ans"
        sheet.cell(row=15,column=4).value = "Student Ans"
        sheet.cell(row=15,column=5).value = "Correct Ans"

        ## create empty files when they are not submitted
        lt = []
        for j in range(7,35) : #Checking the Options
            if(j==32): r,c=16,4
            sheet.cell(row=r,column=c).value = response.iat[i,j]
            sheet.cell(row=r,column=c+1).value = crt[j-7]
            if(response.iat[i,j]==crt[j-7]):#counting right,wrong and not attempted answers 
                rgt+=1
            else: wrg+=1
            r+=1
            lt.append(response.iat[i,j])
            # print(response.iat[i,j])

        na = pd.Series(lt).isna().sum()
        wrg-=na
        sheet.cell(row=10,column=2).value = rgt
        sheet.cell(row=11,column=2).value = mrks
        sheet.cell(row=12,column=2).value = rgt*mrks
        sheet.cell(row=10,column=3).value = wrg
        sheet.cell(row=11,column=3).value = wmrks
        sheet.cell(row=12,column=3).value = wrg*wmrks
        sheet.cell(row=10,column=4).value = na
        sheet.cell(row=11,column=4).value = 0
        sheet.cell(row=10,column=5).value = rgt+wrg+na
        sheet.cell(row=12,column=5).value = str(rgt*mrks+wrg*wmrks)+'/'+str(140)
        wb.save("./sample_output/marksheet/"+sheet.title+'.xlsx')
    return
# generateMarksheet(5,-1)
# generate_concise(5,-1)
