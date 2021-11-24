from tkinter import *
import openpyxl as op
import os
import tkinter.messagebox as tmsg
import pandas as pd
import docx
from datetime import date
import shutil
global bu

def homepage_once_again():
    f4.destroy()
def homepage_once_again1():
    sd=tmsg.askyesno("Warning","You will be redirected to homepage\n Press YES to go to homepage")
    if sd==True:
        f4.destroy()




def back_once_again():
    f13.destroy()
    ##bu.destroy()
def open_excel():
    qa=tmsg.askyesno("WARNNG","make sure not even a single excel file is opened\n Press YES to view records in excel file")
    if qa==True:
        os.startfile(f"E:\\{classname}_with_rank.xlsx")
def back_back():
    if int(total_no_of_students.get())<=25:
        bu.destroy()
    else:
        bu1.destroy()
    f45.destroy()

    ##bu.destroy()
def print_marksheet():
    wsw11 = tmsg.askyesno("Warning", "Make Sure that not even a single WORD file is opened \nPress YES to print marksheet")
    if wsw11 == True:
        os.startfile(f"E:\\{classname}_marksheet.docx","print")
def preview_marksheet():
    ld0=op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
    s0=ld0["Sheet"]
    total_marks_list35 = []
    for i in range(1, int(no_of_subject.get()) + 1):
        total_marks_list35.append(int(s0.cell(row=i + 1, column=2).value) + int(s0.cell(row=i + 1, column=3).value))
    #print(total_marks_list35)
    ld01 = op.load_workbook(f"E:\\{classname}.xlsx")
    s01 = ld01["Sheet"]
    obtained_total_marks1=[]
    obtained_total_marks= []
    for i in range(1,int(total_no_of_students.get())+1):
        for j in range(1, int(no_of_subject.get()) + 1):
            obtained_total_marks1.append(int(s01.cell(row=i + 1, column=3+3*j).value))

        obtained_total_marks.append(obtained_total_marks1)
        obtained_total_marks1 = []
    #print(obtained_total_marks)
    #print(len(obtained_total_marks))
    #print(obtained_total_marks[3])
    percentage_list=[]
    percentage_list_chahiyeko=[]
    for i in range(1,int(total_no_of_students.get())+1):
        x1=obtained_total_marks[i-1]
        for j in range(1, int(no_of_subject.get()) + 1):
            percentage_list.append(int(x1[j-1])/int(total_marks_list35[j-1])*100)
        percentage_list_chahiyeko.append(percentage_list)
        percentage_list=[]
    #print(percentage_list_chahiyeko)
    document = docx.Document()
    if int(total_no_of_students.get())>25:
        theory_total_list_w=total_grade_th+total_grade_th1
        practical_total_list_w=total_grade_pr+total_grade_pr1
        #print(practical_total_list)
        final_total_grade_list_w=total_grade+total_grade1
        final_total_gpa_list_w=total_gpa+total_gpa1
    elif int(total_no_of_students.get()) <= 25:
        theory_total_list_w = total_grade_th2
        practical_total_list_w = total_grade_pr2
        #print(practical_total_list_w)
        #print(practical_total_list)
        final_total_grade_list_w = total_grade2
        final_total_gpa_list_w = total_gpa2

    for r in range(1,int(total_no_of_students.get())+1):
        required_theory_grade=theory_total_list_w[r-1]
        required_practical_grade = practical_total_list_w[r - 1]
        required_total_grade=final_total_grade_list_w[r - 1]
        required_total_gpa=final_total_gpa_list_w[r - 1]
        required_per_chahiyeko_list=percentage_list_chahiyeko[r-1]
        #document.add_picture("logo.PNG")
        heading = document.add_heading("", 0)
        heading.add_run('SHREE MALIKA BASIC SCHOOL').bold = True
        heading.alignment = 1

        para = document.add_paragraph("")
        para.add_run('SINDHUPALCHOWK,NEPAL\n014758745,014587412\nxyz@gmail.com').bold = True
        para.alignment = 1

        head = document.add_paragraph("")
        head.add_run(f'  {exam_type_variable.get()} EXAMINATION GRADESHEET').bold = True

        head.alignment = 1
        ld90=op.load_workbook(f"E:\\{classname}.xlsx")
        s90=ld90["Sheet"]
        """for i in range(1,int(total_no_of_students.get())):
            if s90.cell(row=i+1,column=2).value==variable.get():
                d=i"""
        lda = op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
        sa = lda["Sheet"]



        info=document.add_paragraph("")
        info.add_run(f"Grades secured by ")
        info.add_run(f"{s90.cell(row=r+1,column=1).value}, ").bold=True
        info.add_run(" roll no:-")
        info.add_run(f"{s90.cell(row=r+1,column=2).value}").bold=True
        info.add_run(" with address :-")
        info.add_run(f"{s90.cell(row=r+1,column=3).value}").bold=True
        info.add_run(" of class:-")
        info.add_run(f"{str(classvalue).upper()}").bold=True
        info.add_run(f" in the {exam_type_variable.get().lower()} examination of ")
        info.add_run(f"SHREE MALIKA BASIC SCHOOL ").bold=True
        info.add_run(f"are given below:-")

        table = document.add_table(rows=1, cols=8)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'S.No'
        hdr_cells[1].text = 'NAME OF THE SUBJECTS'
        hdr_cells[2].text = 'GRADE OBTAINED\n(TH)'
        hdr_cells[3].text = 'GRADE OBTAINED\n(PR)'
        hdr_cells[4].text = 'MARKS %'
        hdr_cells[5].text = 'TOTAL GRADE'
        hdr_cells[6].text = 'TOTAL GPA'
        hdr_cells[7].text = 'REMARKS'
        relist=['','','','','','','','','','','','','','','','','','','','']
        #a0=1

        for i in range(1, int(no_of_subject.get()) + 1):
            #y1 = percentage_list_chahiyeko[i - 1]
            row_cells = table.add_row().cells
            row_cells[0].text=f'{i}'
            row_cells[1].text=str(sa.cell(row=i+1,column=1).value).upper()
            row_cells[2].text=required_theory_grade[i-1]
            row_cells[3].text = required_practical_grade[i - 1]
            #row_cells[4].text=str(y1[i-1])
            row_cells[4].text = str(required_per_chahiyeko_list[i - 1])[0:5]
            row_cells[5].text = required_total_grade[i - 1]
            row_cells[6].text = str(required_total_gpa[i - 1])
            row_cells[7].text = relist[i - 1]
            #a0=a0+1
        """for i in range(1,int(total_no_of_students.get())+1):
            y1=percentage_list_chahiyeko[i-1]
            for j in range(1, int(no_of_subject.get()) + 1):
                row_cells[4].text = str(y1[j - 1])"""


        grade_para=document.add_paragraph(f"FINAL GRADE:-{s90.cell(row=r+1,column=3+3*int(no_of_subject.get())+3).value}")
        grade_para.alignment=2
        gpa_para = document.add_paragraph(f"FINAL GPA:-{s90.cell(row=r + 1, column=3 + 3 * int(no_of_subject.get()) + 4).value}")
        gpa_para.alignment = 2
        eva_para=document.add_paragraph("")
        eva_para.add_run("Evaluation of students has been done as follows").bold=True

        eva_para1=document.add_paragraph("A+ - Excellent     A - Very Good")



        eva_para3 = document.add_paragraph("B+ - Good            B - Satisfactory")



        eva_para5 = document.add_paragraph("C+ - Average      C - Needs Improvement")




        eva_para9 = document.add_paragraph("D+,D,E+ and E are considered as a very poor grade")


        date1=str(date.today())
        hello=document.add_paragraph("")
        hello.add_run(f"Date of Issue:-{date1}").bold=True

        preby=document.add_paragraph("_______________\t\t\t\t\t\t\t\t\t_______________")

        preby1 = document.add_paragraph("Prepared by \t\t\t\t\t\t\t\t\t     Principal")
        document.add_page_break()
    document.save(f"E:\\{classname}_marksheet.docx")
    wsw=tmsg.askyesno("Warning","Make Sure that not even a single WORD file is opened \nPress YES to preview marksheet")

    if wsw==True:
        os.startfile(f"E:\\{classname}_marksheet.docx")
    #os.startfile(f"E:\\{classname}_marksheet.docx","print")


    if os.path.exists("E:\\Nursery_marksheet.docx")==True:
        shutil.copyfile(r'E:\\Nursery_marksheet.docx', r'C:\\MKBS_files\\Nursery\\Nursery_marksheet.docx')
    if os.path.exists("E:\\Nursery_with_rank.xlsx") == True:
        shutil.copyfile(r'E:\\Nursery_with_rank.xlsx', r'C:\\MKBS_files\\Nursery\\Nursery_with_rank.xlsx')
    if os.path.exists("E:\\LKG_marksheet.docx")==True:
        shutil.copyfile(r'E:\\LKG_marksheet.docx', r'C:\\MKBS_files\\LKG\\LKG_marksheet.docx')
    if os.path.exists("E:\\LKG_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\LKG_with_rank.xlsx', r'C:\\MKBS_files\\LKG\\LKG_with_rank.xlsx')
    if os.path.exists("E:\\UKG_marksheet.docx")==True:
        shutil.copyfile(r'E:\\UKG_marksheet.docx', r'C:\\MKBS_files\\UKG\\UKG_marksheet.docx')
    if os.path.exists("E:\\UKG_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\UKG_with_rank.xlsx', r'C:\\MKBS_files\\UKG\\UKG_with_rank.xlsx')
    if os.path.exists("E:\\class1_marksheet.docx")==True:
        shutil.copyfile(r'E:\\class1_marksheet.docx', r'C:\\MKBS_files\\Class_1\\class1_marksheet.docx')
    if os.path.exists("E:\\class1_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\class1_with_rank.xlsx', r'C:\\MKBS_files\\Class_1\\class1_with_rank.xlsx')
    if os.path.exists("E:\\class2_marksheet.docx")==True:
        shutil.copyfile(r'E:\\class2_marksheet.docx', r'C:\\MKBS_files\\Class_2\\class2_marksheet.docx')
    if os.path.exists("E:\\class2_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\class2_with_rank.xlsx', r'C:\\MKBS_files\\Class_2\\class2_with_rank.xlsx')
    if os.path.exists("E:\\class3_marksheet.docx")==True:
        shutil.copyfile(r'E:\\class3_marksheet.docx', r'C:\\MKBS_files\\Class_3\\class3_marksheet.docx')
    if os.path.exists("E:\\class3_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\class3_with_rank.xlsx', r'C:\\MKBS_files\\Class_3\\class3_with_rank.xlsx')
    if os.path.exists("E:\\class4_marksheet.docx")==True:
        shutil.copyfile(r'E:\\class4_marksheet.docx', r'C:\\MKBS_files\\Class_4\\class4_marksheet.docx')
    if os.path.exists("E:\\class4_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\class4_with_rank.xlsx', r'C:\\MKBS_files\\Class_4\\class4_with_rank.xlsx')
    if os.path.exists("E:\\class5_marksheet.docx")==True:
        shutil.copyfile(r'E:\\class5_marksheet.docx', r'C:\\MKBS_files\\Class_5\\class5_marksheet.docx')
    if os.path.exists("E:\\class5_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\class5_with_rank.xlsx', r'C:\\MKBS_files\\Class_5\\class5_with_rank.xlsx')
    if os.path.exists("E:\\class6_marksheet.docx")==True:
        shutil.copyfile(r'E:\\class6_marksheet.docx', r'C:\\MKBS_files\\Class_6\\class6_marksheet.docx')
    if os.path.exists("E:\\class6_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\class6_with_rank.xlsx', r'C:\\MKBS_files\\Class_6\\class6_with_rank.xlsx')
    if os.path.exists("E:\\class7_marksheet.docx")==True:
        shutil.copyfile(r'E:\\class7_marksheet.docx', r'C:\\MKBS_files\\Class_7\\class7_marksheet.docx')
    if os.path.exists("E:\\class7_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\class7_with_rank.xlsx', r'C:\\MKBS_files\\Class_7\\class7_with_rank.xlsx')
    if os.path.exists("E:\\class8_marksheet.docx")==True:
        shutil.copyfile(r'E:\\class8_marksheet.docx', r'C:\\MKBS_files\\Class_8\\class8_marksheet.docx')
    if os.path.exists("E:\\class8_with_rank.xlsx")==True:
        shutil.copyfile(r'E:\\class8_with_rank.xlsx', r'C:\\MKBS_files\\Class_8\\class8_with_rank.xlsx')
def folder():
    os.startfile("C:\\MKBS_files")
def proceed_for_marksheet():
    global f45
    if os.path.exists(f"C:\\MKBS_files")==False:
        os.mkdir(f"C:\\MKBS_files")
    if os.path.exists(f"C:\\MKBS_files") == True:
        if os.path.exists(f"C:\\MKBS_files\\Nursery")==False:
            os.mkdir(f"C:\\MKBS_files\\Nursery")
        if os.path.exists(f"C:\\MKBS_files\\LKG") == False:
            os.mkdir(f"C:\\MKBS_files\\LKG")
        if os.path.exists(f"C:\\MKBS_files\\UKG") == False:
            os.mkdir(f"C:\\MKBS_files\\UKG")
        if os.path.exists(f"C:\\MKBS_files\\Class_1") == False:
            os.mkdir(f"C:\\MKBS_files\\Class_1")
        if os.path.exists(f"C:\\MKBS_files\\Class_2") == False:
            os.mkdir(f"C:\\MKBS_files\\Class_2")
        if os.path.exists(f"C:\\MKBS_files\\Class_3") == False:
            os.mkdir(f"C:\\MKBS_files\\Class_3")
        if os.path.exists(f"C:\\MKBS_files\\Class_4") == False:
            os.mkdir(f"C:\\MKBS_files\\Class_4")
        if os.path.exists(f"C:\\MKBS_files\\Class_5") == False:
            os.mkdir(f"C:\\MKBS_files\\Class_5")
        if os.path.exists(f"C:\\MKBS_files\\Class_6") == False:
            os.mkdir(f"C:\\MKBS_files\\Class_6")
        if os.path.exists(f"C:\\MKBS_files\\Class_7") == False:
            os.mkdir(f"C:\\MKBS_files\\Class_7")
        if os.path.exists(f"C:\\MKBS_files\\Class_8") == False:
            os.mkdir(f"C:\\MKBS_files\\Class_8")



    f45 = Frame(f12,bg="cyan2")
    f45.place(x=0, y=0, width=2500, height=1500)

    df = pd.read_excel(f"E:\\{classname}.xlsx")
    af = df.sort_values("PER", ascending=False)
    af=df.drop(['ADDRESS'],axis=1)
    af.to_excel(f"E:\\{classname}_with_rank.xlsx")

    listp = [ 'D','E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
             'U',
             'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM',
             'AN',
             'AO', 'AP']
    ld890 = op.load_workbook(f"E:\\{classname}_with_rank.xlsx")
    ws = ld890.active
    for i in range(1, len(listp) + 1):
        ws.column_dimensions[listp[i - 1]].width = 7
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 10
    #ws.column_dimensions['D'].width = 20
    s890 = ld890["Sheet1"]
    s890.cell(row=1, column=1).value = "ENTRY NO."
    s890.cell(row=1, column=4 + 3 * int(no_of_subject.get()) + 5-1).value = "RANK"
    for i in range(1, int(total_no_of_students.get()) + 1):
        s890.cell(row=i + 1, column=4 + 3 * int(no_of_subject.get()) + 5-1).value = i
    ld890.save(f"E:\\{classname}_with_rank.xlsx")
    # students_name=[i for i in range(i,total_no_of_students.get()+1) s123.cell(row=i+1,column=1).value]
    students_name = []
    roll_list = []
    for i in range(1, int(total_no_of_students.get()) + 1):
        students_name.append(str(s890.cell(row=i + 1, column=2).value))
    # print(students_name)
    ldld=op.load_workbook(f"E:\\{classname}.xlsx")
    ss=ldld["Sheet"]
    for i in range(1, int(total_no_of_students.get()) + 1):
        roll_list.append(str(ss.cell(row=i + 1, column=2).value))
    both_list = []

    for i in range(0, len(students_name)):
        both_list.append(roll_list[i])
    global variable,exam_type_variable
    variable = StringVar()
    exam_type_variable=StringVar()
    exam_tpe_list=['FIRST TERMINAL','SECOND TERMINAL','MID TERMINAL','THIRD TERMINAL','PRE-BOARD','FINAL']
    dpmenu = OptionMenu(f45, variable, *both_list)
    dpmenu.grid(row=10, column=0)
    exam_type_menu=OptionMenu(f45,exam_type_variable,*exam_tpe_list)
    exam_type_menu.grid(row=10, column=1,padx=5)
    variable.set("SELECT ROLL NO.")
    exam_type_variable.set("SELECT EXAM TYPE")
    pm_button = Button(f45, text="View Marksheet", font="arial 10 bold", bg="bisque",fg="maroon", command=view_marksheet)
    pm_button.grid(row=10, column=2, padx=5)
    excel_opening = Button(f45, text="View Records in Excel file with Rank", font="arial 10 bold", bg="bisque",fg="maroon",
                           command=open_excel)
    excel_opening.grid(row=10, column=7, padx=5)
    back = Button(f45, text="Back to previous page", font="arial 10 bold", bg="bisque",fg="maroon", command=back_back)
    back.grid(row=10, column=8, padx=5)
    back = Button(f45, text="Homepage", font="arial 10 bold", bg="bisque",fg="maroon", command=homepage_once_again)
    back.grid(row=10, column=9, padx=5)
    folder_button = Button(f45, text="View Files in Folder", font="arial 10 bold", bg="bisque", fg="maroon", command=folder)
    folder_button.grid(row=10, column=10, padx=5)


def view_marksheet():

    """print(total_grade_th1)
    print(len(total_grade_th1))
    print(total_grade_th)
    print(len(total_grade_th))"""
    global f46
    ld0 = op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
    s0 = ld0["Sheet"]
    total_marks_list35 = []
    for i in range(1, int(no_of_subject.get()) + 1):
        total_marks_list35.append(int(s0.cell(row=i + 1, column=2).value) + int(s0.cell(row=i + 1, column=3).value))
    # print(total_marks_list35)
    ld01 = op.load_workbook(f"E:\\{classname}.xlsx")
    s01 = ld01["Sheet"]
    obtained_total_marks1 = []
    obtained_total_marks = []
    for i in range(1, int(total_no_of_students.get()) + 1):
        for j in range(1, int(no_of_subject.get()) + 1):
            obtained_total_marks1.append(int(s01.cell(row=i + 1, column=3 + 3 * j).value))

        obtained_total_marks.append(obtained_total_marks1)
        obtained_total_marks1 = []
    # print(obtained_total_marks)
    # print(len(obtained_total_marks))
    # print(obtained_total_marks[3])
    percentage_list = []
    percentage_list_chahiyeko = []
    for i in range(1, int(total_no_of_students.get()) + 1):
        x1 = obtained_total_marks[i - 1]
        for j in range(1, int(no_of_subject.get()) + 1):
            percentage_list.append(int(x1[j - 1]) / int(total_marks_list35[j - 1]) * 100)
        percentage_list_chahiyeko.append(percentage_list)
        percentage_list = []



    if int(total_no_of_students.get())>25:

        theory_total_list=total_grade_th+total_grade_th1
        practical_total_list=total_grade_pr+total_grade_pr1
        #print(practical_total_list)
        final_total_grade_list=total_grade+total_grade1
        final_total_gpa_list=total_gpa+total_gpa1


        if variable.get()!="SELECT ROLL NO." and exam_type_variable.get()!="SELECT EXAM TYPE" :

            f46=Frame(f45,bg="cyan2")
            f46.place(x=0,y=70)
            space_label1=Label(f46,text="\t\t\t\t\t\t\t\t\t",bg="cyan2")
            space_label1.grid(row=0,column=0)
            school_name2=Label(f46,text="SHREE MALIKA BASIC SCHOOL",font="arial 30 bold",bg="cyan2")
            school_name2.grid(row=0,column=1)

            f47 = Frame(f45,bg="cyan2")
            f47.place(x=0, y=120)
            space_label3 = Label(f47, text="\t\t\t\t\t\t\t\t\t\t\t\t\t",bg="cyan2")
            space_label3.grid(row=0, column=0)
            school_name4 = Label(f47, text="SINDHUPALCHOWK,NEPAL", font="arial 15 bold",bg="cyan2")
            school_name4.grid(row=0, column=1)

            f48 = Frame(f45,bg="cyan2")
            f48.place(x=0, y=170)
            space_label5 = Label(f48, text="\t\t\t\t\t\t\t\t\t\t\t\t\t\t",bg="cyan2")
            space_label5.grid(row=0, column=0)
            school_name6 = Label(f48, text="GRADE SHEET", font="arial 15 bold",bg="cyan2")
            school_name6.grid(row=0, column=1)
            ld321=op.load_workbook(f"E:\\{classname}.xlsx")
            s321=ld321["Sheet"]

            for i in range(1,int(total_no_of_students.get())+1):
                if s321.cell(row=i+1,column=2).value==variable.get():
                    m=i
            f49 = Frame(f45,bg="cyan2")
            f49.place(x=0, y=230)
            space_label7 = Label(f49, text="\t\t\t\t\t",bg="cyan2")
            space_label7.grid(row=0, column=0)
            space_label7 = Label(f49, text="\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",bg="cyan2")
            space_label7.grid(row=0, column=1)
            school_name8 = Label(f49, text=f"Grades secured by {s321.cell(row=m+1,column=1).value} roll no.- {variable.get()} in the {exam_type_variable.get()} examination are given below :-", font="arial 14 bold",bg="cyan2")
            school_name8.grid(row=0, column=1)



            f51 = Frame(f45,bg="cyan2")
            f51.place(x=0, y=330)
            space_label11 = Label(f51, text="\t\t\t\t\t\t\t",bg="cyan2")
            space_label11.grid(row=0, column=0)
            school_name12 = Label(f51,
                                text="NAME OF THE SUBJECT",
                                font="arial 12 bold",bg="cyan2")
            school_name12.grid(row=0, column=1)
            school_name_a = Label(f51,
                                text="   GRADES OBTAINED(TH)",
                                font="arial 12 bold",bg="cyan2")
            school_name_a.grid(row=0, column=2)
            school_name_b = Label(f51,
                                  text="  GRADES OBTAINED(PR)",
                                  font="arial 12 bold",bg="cyan2")
            school_name_b.grid(row=0, column=3)
            school_name_b = Label(f51,
                                  text="  MARKS %",
                                  font="arial 12 bold",bg="cyan2")
            school_name_b.grid(row=0, column=4)
            school_name_c = Label(f51,
                                  text="  TOTAL GRADE",
                                  font="arial 12 bold",bg="cyan2")
            school_name_c.grid(row=0, column=5)
            school_name_d = Label(f51,
                                  text="  TOTAL GPA",
                                  font="arial 12 bold",bg="cyan2")
            school_name_d.grid(row=0, column=6)


            ldf=op.load_workbook(f"E:{classname}_full_marks.xlsx")
            sf=ldf["Sheet"]
            ldf1 = op.load_workbook(f"E:{classname}.xlsx")
            sf1 = ldf1["Sheet"]
            for i in range(1,int(no_of_subject.get())+1):
                sub_label_w =Label(f51,text=sf.cell(row=i+1,column=1).value.upper(),font='arial 10 bold',bg="cyan2")
                sub_label_w.grid(row=i,column=1)
            for i in range(1,int(total_no_of_students.get())+1):
                if str(sf1.cell(row=i+1,column=2).value)==str(variable.get()):
                    fg=i
            #print(variable.get())
            global chahiyeko_list,chahiyeko_list1,chahiyeko_list2,chahiyeko_list3
            chahiyeko_list=theory_total_list[fg-1]
            chahiyeko_list1 = practical_total_list[fg-1]
            chahiyeko_list2 = final_total_grade_list[fg-1]
            chahiyeko_list3 = final_total_gpa_list[fg-1]
            chahiyeko_list4=percentage_list_chahiyeko[fg-1]
            #print(chahiyeko_list)
            for i in range(0,len(chahiyeko_list)):
                sub_label_wq = Label(f51, text=chahiyeko_list[i], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i+1, column=2)
            for i in range(0,len(chahiyeko_list1)):
                sub_label_wq = Label(f51, text=chahiyeko_list1[i], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i+1, column=3)
            for i in range(0,len(chahiyeko_list4)):
                sub_label_wq = Label(f51, text=str(chahiyeko_list4[i])[0:5], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i+1, column=4)
            for i in range(0,len(chahiyeko_list2)):
                sub_label_wq = Label(f51, text=chahiyeko_list2[i], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i+1, column=5)
            for i in range(0,len(chahiyeko_list3)):
                sub_label_wq = Label(f51, text=chahiyeko_list3[i], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i+1, column=6)
            f67=Frame(f45,bg="cyan2")
            f67.place(x=270,y=650)
            ldr=op.load_workbook(f"E:\\{classname}.xlsx")
            s90=ldr["Sheet"]
            preview_marksheet_button=Button(f67,text="Preview Marksheet",font="arial 12 bold",bg="bisque",fg="maroon",command=preview_marksheet)
            preview_marksheet_button.grid(row=0,column=1)
            print_marksheet_button = Button(f67, text="Print Marksheet", font="arial 12 bold", bg="bisque",fg="maroon",
                                              command=print_marksheet)
            print_marksheet_button.grid(row=0, column=0,padx=10)
            sp_label=Label(f67,text="\t\t\t\t\t\t\t",bg="cyan2")
            sp_label.grid(row=0, column=2,padx=30)
            grade_label = Label(f67, text=f"Final Grade-{s90.cell(row=fg+1,column=3+3*int(no_of_subject.get())+3).value}     Final GPA-{s90.cell(row=fg+1, column=3 + 3 * int(no_of_subject.get()) + 4).value}", font="arial 12 bold",bg="cyan2")
            grade_label.grid(row=0, column=3,padx=30)
            """global remarks_dict
            remarks_dict={i:StringVar() for i in range(1,int(no_of_subject.get())+1)}
            for i in remarks_dict:
                remarks_entry=Entry(f51,textvariable=remarks_dict[i],border=4,font="arial 10 bold",relief=SUNKEN,width=15)
                remarks_entry.grid(row=i,column=6)"""
        else:
            tmsg.showinfo("Error","Please select the roll no. and the exam type")
            #print(len(total_gpa_th1))
            #print(len(total_gpa_th))
            #print(theory_total_list)
            #print(len(theory_total_list))
    elif int(total_no_of_students.get())<=25:
        theory_total_list = total_grade_th2
        practical_total_list = total_grade_pr2
        # print(practical_total_list)
        final_total_grade_list = total_grade2
        final_total_gpa_list = total_gpa2



        if variable.get() != "SELECT STUDENT'S NAME" and exam_type_variable.get() != "SELECT EXAM TYPE":

            f46 = Frame(f45,bg="cyan2")
            f46.place(x=0, y=70)
            space_label1 = Label(f46, text="\t\t\t\t\t\t\t\t\t",bg="cyan2")
            space_label1.grid(row=0, column=0)
            school_name2 = Label(f46, text="SHREE MALIKA BASIC SCHOOL", font="arial 30 bold",bg="cyan2")
            school_name2.grid(row=0, column=1)

            f47 = Frame(f45,bg="cyan2")
            f47.place(x=0, y=120)
            space_label3 = Label(f47, text="\t\t\t\t\t\t\t\t\t\t\t\t\t",bg="cyan2")
            space_label3.grid(row=0, column=0)
            school_name4 = Label(f47, text="SINDHUPALCHOWK,NEPAL", font="arial 15 bold",bg="cyan2")
            school_name4.grid(row=0, column=1)

            f48 = Frame(f45,bg="cyan2")
            f48.place(x=0, y=170)
            space_label5 = Label(f48, text="\t\t\t\t\t\t\t\t\t\t\t\t\t\t",bg="cyan2")
            space_label5.grid(row=0, column=0)
            school_name6 = Label(f48, text="GRADE SHEET", font="arial 15 bold",bg="cyan2")
            school_name6.grid(row=0, column=1)
            ld453=op.load_workbook(f"E:\\{classname}.xlsx")
            s453=ld453["Sheet"]
            for i in range(1,int(total_no_of_students.get())+1):
                if s453.cell(row=i+1,column=2).value==variable.get():
                    m=i
            f49 = Frame(f45,bg="cyan2")
            f49.place(x=0, y=230)
            space_label7 = Label(f49, text="\t\t\t\t\t",bg="cyan2")
            space_label7.grid(row=0, column=0)
            space_label7 = Label(f49, text="\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",bg="cyan2")
            space_label7.grid(row=0, column=1)
            school_name8 = Label(f49,
                                 text=f"Grades secured by {s453.cell(row=m + 1, column=1).value} roll no.- {variable.get()} in the {exam_type_variable.get()} examination are given below :-",
                                 font="arial 14 bold",bg="cyan2")
            school_name8.grid(row=0, column=1)

            f51 = Frame(f45,bg="cyan2")
            f51.place(x=0, y=330)
            space_label11 = Label(f51, text="\t\t\t\t\t\t\t",bg="cyan2")
            space_label11.grid(row=0, column=0)
            school_name12 = Label(f51,
                                  text="NAME OF THE SUBJECT",
                                  font="arial 12 bold",bg="cyan2")
            school_name12.grid(row=0, column=1)
            school_name_a = Label(f51,
                                  text="   GRADES OBTAINED(TH)",
                                  font="arial 12 bold",bg="cyan2")
            school_name_a.grid(row=0, column=2)
            school_name_b = Label(f51,
                                  text="  GRADES OBTAINED(PR)",
                                  font="arial 12 bold",bg="cyan2")
            school_name_b.grid(row=0, column=3)
            school_name_b = Label(f51,
                                  text="  MARKS %",
                                  font="arial 12 bold",bg="cyan2")
            school_name_b.grid(row=0, column=4)
            school_name_c = Label(f51,
                                  text="  TOTAL GRADE",
                                  font="arial 12 bold",bg="cyan2")
            school_name_c.grid(row=0, column=5)
            school_name_d = Label(f51,
                                  text="  TOTAL GPA",
                                  font="arial 12 bold",bg="cyan2")
            school_name_d.grid(row=0, column=6)


            ldf = op.load_workbook(f"E:{classname}_full_marks.xlsx")
            sf = ldf["Sheet"]
            ldf1 = op.load_workbook(f"E:{classname}.xlsx")
            sf1 = ldf1["Sheet"]
            for i in range(1, int(no_of_subject.get()) + 1):
                sub_label_w = Label(f51, text=sf.cell(row=i + 1, column=1).value.upper(), font='arial 10 bold',bg="cyan2")
                sub_label_w.grid(row=i, column=1)
            for i in range(1, int(total_no_of_students.get()) + 1):
                if str(sf1.cell(row=i + 1, column=2).value) == str(variable.get()):
                    fg = i
            # print(variable.get())
            chahiyeko_list = theory_total_list[fg - 1]
            chahiyeko_list1 = practical_total_list[fg - 1]
            chahiyeko_list2 = final_total_grade_list[fg - 1]
            chahiyeko_list3 = final_total_gpa_list[fg - 1]

            chahiyeko_list4 = percentage_list_chahiyeko[fg - 1]
            # print(chahiyeko_list)
            for i in range(0, len(chahiyeko_list)):
                sub_label_wq = Label(f51, text=chahiyeko_list[i], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i + 1, column=2)
            for i in range(0, len(chahiyeko_list1)):
                sub_label_wq = Label(f51, text=chahiyeko_list1[i], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i + 1, column=3)
            for i in range(0, len(chahiyeko_list4)):
                sub_label_wq = Label(f51, text=str(chahiyeko_list4[i])[0:5], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i + 1, column=4)
            for i in range(0, len(chahiyeko_list2)):
                sub_label_wq = Label(f51, text=chahiyeko_list2[i], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i + 1, column=5)
            for i in range(0, len(chahiyeko_list3)):
                sub_label_wq = Label(f51, text=chahiyeko_list3[i], font='arial 10 bold',bg="cyan2")
                sub_label_wq.grid(row=i + 1, column=6)


            f67 = Frame(f45,bg="cyan2")
            f67.place(x=270, y=650)
            ldr = op.load_workbook(f"E:\\{classname}.xlsx")
            s90 = ldr["Sheet"]
            preview_marksheet_button = Button(f67, text="Preview Marksheet", font="arial 12 bold", bg="bisque",fg="maroon",command=preview_marksheet)
            preview_marksheet_button.grid(row=0, column=1)
            print_marksheet_button = Button(f67, text="Print Marksheet", font="arial 12 bold", bg="bisque",fg="maroon",
                                              command=print_marksheet)
            print_marksheet_button.grid(row=0, column=0,padx=10)
            sp_label = Label(f67, text="\t\t\t\t\t\t\t",bg="cyan2")
            sp_label.grid(row=0, column=2, padx=30)
            grade_label = Label(f67,
                                text=f"Final Grade-{s90.cell(row=fg + 1, column=3 + 3 * int(no_of_subject.get()) + 3).value}     Final GPA-{s90.cell(row=fg + 1, column=3 + 3 * int(no_of_subject.get()) + 4).value}",
                                font="arial 12 bold",bg="cyan2")
            grade_label.grid(row=0, column=3, padx=30)

            """global remarks_dict1
            remarks_dict1 = {i: StringVar() for i in range(1, int(no_of_subject.get()) + 1)}
            for i in remarks_dict1:
                remarks_entry = Entry(f51, textvariable=remarks_dict1[i], border=4,font="arial 10 bold",relief=SUNKEN, width=15)
                remarks_entry.grid(row=i, column=6)"""



        else:
            tmsg.showinfo("Error", "Please select the roll no. of the student and the exam type")


def full_marks():

    global windows
    ld111=op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
    s111=ld111["Sheet"]
    windows=Tk()
    windows.title("Full Marks of Subjects")
    windows.geometry("1300x130+0+0")

    lab1=Label(windows,text="Subject Name",font="arial 12 bold")
    lab1.grid(row=0,column=0)
    lab2 = Label(windows, text="Full Marks(Theory) ", font="arial 12 bold")
    lab2.grid(row=1, column=0)
    lab3 = Label(windows, text="Full Marks(Practical) ", font="arial 12 bold")
    lab3.grid(row=2, column=0)
    for i in range(1,int(no_of_subject.get())+1):
        lab11=Label(windows,text=s111.cell(row=i+1,column=1).value, font="arial 12 bold")
        lab11.grid(row=0,column=i)
        lab12 = Label(windows, text=s111.cell(row=i + 1, column=2).value, font="arial 12 bold")
        lab12.grid(row=1, column=i)
        lab13 = Label(windows, text=s111.cell(row=i + 1, column=3).value, font="arial 12 bold")
        lab13.grid(row=2, column=i)


def print_marks1():
    ld345 = op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
    s345 = ld345["Sheet"]
    #global total_marks_list
    total_marks_list = []
    for i in range(1, int(no_of_subject.get()) + 1):
        total_marks_list.append(int(s345.cell(row=i + 1, column=2).value) + int(s345.cell(row=i + 1, column=3).value))
    #print(total_marks_list)
    list86 = []
    list87 = []
    ld123 = op.load_workbook(f"E:\\{classname}.xlsx")
    s123 = ld123["Sheet"]
    th_marks_third=[]
    pr_marks_third = []
    for i in range(1, (int(total_no_of_students.get()))+1):
        v = th_memory1_third[i - 1]
        w = pr_memory1_third[i - 1]
        for j in range(0, int(no_of_subject.get())):
            list86.append(v[j].get())
            list87.append(w[j].get())
    y = 2
    for i in range(0, (int(no_of_subject.get()) * (int(total_no_of_students.get()))), (int(no_of_subject.get()))):
        if i == 0:
            th_marks_third.append(list86[0:int(no_of_subject.get())])
            pr_marks_third.append(list87[0:int(no_of_subject.get())])
        else:
            th_marks_third.append(list86[i:int(no_of_subject.get()) * y])
            pr_marks_third.append(list87[i:int(no_of_subject.get()) * y])
            y = y + 1
    done = False
    b1 = 0

    for i in range(1, (int(total_no_of_students.get())) + 1):
        u1 = th_marks_third[i - 1]
        fo1 = pr_marks_third[i - 1]
        for j in range(1, int(no_of_subject.get()) + 1):
            if str(u1[j - 1]).isdigit()==True and str(fo1[j - 1]).isdigit()==True:
                if int(u1[j - 1]) <= int(th_fm_list[j - 1]) and int(fo1[j - 1]) <= int(pr_fm_list[j - 1]):
                    s123.cell(row=i + 1, column=3 * j + 1).value = u1[j - 1]
                    s123.cell(row=i + 1, column=3 * j + 2).value = fo1[j - 1]
                    as2 = int(s123.cell(row=i + 1, column=3 * j + 1).value) + int(s123.cell(row=i + 1, column=3 * j + 2).value)
                    s123.cell(row=i + 1, column=3 * j + 3).value = as2


                    ld123.save(f"E:\\{classname}.xlsx")
                else:
                    b1 = b1 + 1
                    tmsg.showinfo(f"Error",f"obtained marks  exceeded the full marks in {i}th row {j}th subject ")
                    break
                    done = True
            else:
                b1 = b1 + 1
                tmsg.showinfo("Error", f"Marks shoud be postive integer\n It's input field can not be left empty\ncheck {i}th row {j}th subject ")
                break
                done = True
        if done:
            break
    if b1==0:
        s123.cell(row=1,column=(3+int(no_of_subject.get())*3+1)).value="G.T"
        s123.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 2)).value = "PER"
        s123.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 3)).value = "F.GR"
        s123.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 4)).value = "F.GPA"
        ld123.save(f"E:\\{classname}.xlsx")
        th_sum=0
        pr_sum=0
        for i in range(0,len(th_fm_list)):
            th_sum+=int(th_fm_list[i])
        for i in range(0,len(pr_fm_list)):
            pr_sum+=int(pr_fm_list[i])
        full_marks=th_sum+pr_sum
        th_om_sum=0
        pr_om_sum=0
        for i in range(1, (int(total_no_of_students.get())) + 1):
            u11 = th_marks_third[i - 1]
            fo11 = pr_marks_third[i - 1]
            for j in range(1, int(no_of_subject.get()) + 1):
                th_om_sum+=int(u11[j-1])
                pr_om_sum+=int(fo11[j-1])
            gt_sum=th_om_sum+pr_om_sum
            s123.cell(row=i+1,column=3+int(no_of_subject.get())*3+1).value=gt_sum
            th_om_sum=0
            pr_om_sum=0
            percentage=round(gt_sum/full_marks*100,3)
            s123.cell(row=i + 1, column=3 + int(no_of_subject.get()) * 3 + 2).value = percentage

            ld123.save(f"E:\\{classname}.xlsx")
        global total_grade_th2,total_gpa_th2
        indi_gpa_th2=[]
        total_gpa_th2=[]
        indi_grade_th2=[]
        total_grade_th2=[]
        for i in range(1,int(total_no_of_students.get())+1):
            u12 = th_marks_third[i - 1]
            #f12 = pr_marks_third[i - 1]
            for j in range(1,int(no_of_subject.get())+1):
                if int(u12[j-1])==0:
                    gpa="-"
                    grade="-"
                    indi_gpa_th2.append(gpa)
                    indi_grade_th2.append(grade)
                else:
                    if float(u12[j-1])>=0.9*float(th_fm_list[j-1]) and float(u12[j-1])<=1.0*float(th_fm_list[j-1]):
                        gpa=4.0
                        grade="A+"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j-1])>=0.8*float(th_fm_list[j-1]) and float(u12[j-1])<0.9*float(th_fm_list[j-1]):
                        gpa=3.6
                        grade="A"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j - 1]) >= 0.7 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.8 * float(
                            th_fm_list[j - 1]):
                        gpa = 3.2
                        grade = "B+"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j - 1]) >= 0.6 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.7 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.8
                        grade = "B"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j - 1]) >= 0.5 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.6 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.4
                        grade = "C+"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j - 1]) >= 0.4 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.5 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.0
                        grade = "C"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j - 1]) >= 0.3 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.4 * float(
                            th_fm_list[j - 1]):
                        gpa = 1.6
                        grade = "D+"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j - 1]) >= 0.2 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.3 * float(
                            th_fm_list[j - 1]):
                        gpa = 1.2
                        grade = "D"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j - 1]) >= 0.1 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.2 * float(
                            th_fm_list[j - 1]):
                        gpa = 0.8
                        grade = "E+"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)
                    elif float(u12[j - 1]) >= 0.0 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.1 * float(
                            th_fm_list[j - 1]):
                        gpa = 0.4
                        grade = "E"
                        indi_gpa_th2.append(gpa)
                        indi_grade_th2.append(grade)

            #total_gpa_th2.append(indi_gpa_th2)
            #total_grade_th2.append(indi_grade_th2)
        aq=2
        for i in range(0,int(no_of_subject.get())*int(total_no_of_students.get()),int(no_of_subject.get())):
            if i==0:
                total_gpa_th2.append(indi_gpa_th2[0:int(no_of_subject.get())])
                total_grade_th2.append(indi_grade_th2[0:int(no_of_subject.get())])
            else:
                total_gpa_th2.append(indi_gpa_th2[i:int(no_of_subject.get())*aq])
                total_grade_th2.append(indi_grade_th2[i:int(no_of_subject.get())*aq])
                aq=aq+1
        global total_gpa_pr2,total_grade_pr2
        indi_gpa_pr2=[]
        indi_grade_pr2=[]
        total_gpa_pr2=[]
        total_grade_pr2=[]
        for i in range(1,int(total_no_of_students.get())+1):
            fo12 = pr_marks_third[i - 1]
            for j in range(1,int(no_of_subject.get())+1):
                if int(fo12[j-1])==0:
                    gpa_pr="-"
                    grade_pr="-"
                    indi_gpa_pr2.append(gpa_pr)
                    indi_grade_pr2.append(grade_pr)
                else:
                    if float(fo12[j-1])>=0.9*float(pr_fm_list[j-1]) and float(fo12[j-1])<=1.0*float(pr_fm_list[j-1]):
                        gpa_pr = 4.0
                        grade_pr = "A+"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j-1])>=0.8*float(pr_fm_list[j-1]) and float(fo12[j-1])<0.9*float(pr_fm_list[j-1]):
                        gpa_pr = 3.6
                        grade_pr = "A"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.7 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.8 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 3.2
                        grade_pr = "B+"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.6 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.7 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.8
                        grade_pr = "B"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.5 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.6 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.4
                        grade_pr = "C+"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.4 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.5 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.0
                        grade_pr = "C"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.3 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.4 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 1.6
                        grade_pr = "D+"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.2 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.3 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 1.2
                        grade_pr = "D"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.1 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.2 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 0.8
                        grade_pr = "E+"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.0 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.1 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 0.4
                        grade_pr = "E"
                        indi_gpa_pr2.append(gpa_pr)
                        indi_grade_pr2.append(grade_pr)

            #total_gpa_th2.append(indi_gpa_th2)
            #total_grade_th2.append(indi_grade_th2)
        aq1=2
        for i in range(0,int(no_of_subject.get())*int(total_no_of_students.get()),int(no_of_subject.get())):
            if i==0:
                total_gpa_pr2.append(indi_gpa_pr2[0:int(no_of_subject.get())])
                total_grade_pr2.append(indi_grade_pr2[0:int(no_of_subject.get())])
            else:
                total_gpa_pr2.append(indi_gpa_pr2[i:int(no_of_subject.get())*aq1])
                total_grade_pr2.append(indi_grade_pr2[i:int(no_of_subject.get())*aq1])
                aq1=aq1+1
        global total_gpa2,total_grade2
        indi_gpa2 = []
        indi_grade2 = []
        total_gpa2 = []
        total_grade2 = []
        for i in range(1,int(total_no_of_students.get())+1):
            fo12 = pr_marks_third[i - 1]
            u12 = th_marks_third[i - 1]
            for j in range(1, int(no_of_subject.get()) + 1):
                if int(fo12[j - 1])+int(u12[j-1]) == 0:
                    gpa = "-"
                    grade = "-"
                    indi_gpa2.append(gpa)
                    indi_grade2.append(grade)
                else:
                    if float(fo12[j-1])+float(u12[j-1])>=0.9*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<=1.0*float(total_marks_list[j-1]):
                        gpa = 4.0
                        grade = "A+"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.8*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.9*float(total_marks_list[j-1]):
                        gpa = 3.6
                        grade = "A"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.7*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.8*float(total_marks_list[j-1]):
                        gpa = 3.2
                        grade = "B+"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.6*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.7*float(total_marks_list[j-1]):
                        gpa = 2.8
                        grade = "B"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.5*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.6*float(total_marks_list[j-1]):
                        gpa = 2.4
                        grade = "C+"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.4*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.5*float(total_marks_list[j-1]):
                        gpa = 2.0
                        grade = "C"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.3*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.4*float(total_marks_list[j-1]):
                        gpa = 1.6
                        grade = "D+"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.2*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.3*float(total_marks_list[j-1]):
                        gpa = 1.2
                        grade = "D"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.1*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.2*float(total_marks_list[j-1]):
                        gpa = 0.8
                        grade = "E+"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.0*float(total_marks_list[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.1*float(total_marks_list[j-1]):
                        gpa = 0.4
                        grade = "E"
                        indi_gpa2.append(gpa)
                        indi_grade2.append(grade)
        aq2 = 2
        for i in range(0, int(no_of_subject.get()) * int(total_no_of_students.get()), int(no_of_subject.get())):
            if i == 0:
                total_gpa2.append(indi_gpa2[0:int(no_of_subject.get())])
                total_grade2.append(indi_grade2[0:int(no_of_subject.get())])
            else:
                total_gpa2.append(indi_gpa2[i:int(no_of_subject.get()) * aq2])
                total_grade2.append(indi_grade2[i:int(no_of_subject.get()) * aq2])
                aq2 = aq2 + 1
        #print(total_gpa)
        credit_sum = 0
        gpa_list = []
        for i in range(1, len(credit_list) + 1):
            credit_sum += int(credit_list[i - 1])

        for i in range(1, int(total_no_of_students.get()) + 1):
            xd = total_gpa2[i - 1]
            partial_gpa = 0
            for j in range(1, int(no_of_subject.get()) + 1):
                if type(xd[j - 1]) == float:
                    partial_gpa = partial_gpa + (float(xd[j - 1]) * int(credit_list[j - 1]))
            final_gpa = round(partial_gpa / credit_sum, 2)
            gpa_list.append(final_gpa)
            # partial_gpa = 0

        # print(gpa_list)
        # print(len(gpa_list))
        grade_list = []
        for i in range(1, len(gpa_list) + 1):
            if float(gpa_list[i - 1]) >= 3.6 and float(gpa_list[i - 1]) <= 4.0:
                final_grade = 'A+'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 3.2 and float(gpa_list[i - 1]) < 3.6:
                final_grade = 'A'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 2.8 and float(gpa_list[i - 1]) < 3.2:
                final_grade = 'B+'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 2.4 and float(gpa_list[i - 1]) < 2.8:
                final_grade = 'B'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 2.0 and float(gpa_list[i - 1]) < 2.4:
                final_grade = 'C+'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 1.6 and float(gpa_list[i - 1]) < 2.0:
                final_grade = 'C'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 1.2 and float(gpa_list[i - 1]) < 1.6:
                final_grade = 'D+'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 0.8 and float(gpa_list[i - 1]) < 1.2:
                final_grade = 'D'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 0.4 and float(gpa_list[i - 1]) < 0.8:
                final_grade = 'E+'
                grade_list.append(final_grade)
            elif float(gpa_list[i - 1]) >= 0.0 and float(gpa_list[i - 1]) < 0.4:
                final_grade = 'E'
                grade_list.append(final_grade)

        for i in range(1, len(gpa_list) + 1):
            s123.cell(row=i + 1, column=3 + 3 * int(no_of_subject.get()) + 4).value = gpa_list[i - 1]
            s123.cell(row=i + 1, column=3 + 3 * int(no_of_subject.get()) + 3).value = grade_list[i - 1]
            ld123.save(f"E:\\{classname}.xlsx")
        global bu
        f34=Frame(f12,bg="cyan2")
        f34.place(x=665 , y=638)
        bu = Button(f34, text="Proceed for marksheet", bg="bisque",fg="maroon", font="arial 12 bold",command=proceed_for_marksheet)
        bu.grid(row=0,column=0)
        bu = Button(f34, text="Exit", bg="bisque",fg="maroon", font="arial 12 bold", command=homepage_once_again1)
        bu.grid(row=0, column=1,padx=10)



def print_marks():
    ld3456 = op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
    s3456 = ld3456["Sheet"]
    #global total_marks_list
    total_marks_list1 = []
    for i in range(1, int(no_of_subject.get()) + 1):
        total_marks_list1.append(int(s3456.cell(row=i + 1, column=2).value) + int(s3456.cell(row=i + 1, column=3).value))
    # print(total_marks_list)
    list76=[]
    list77=[]
    ld678=op.load_workbook(f"E:\\{classname}.xlsx")
    s678=ld678["Sheet"]
    th_marks_second=[]
    pr_marks_second=[]
    for i in range(1, (int(total_no_of_students.get())-25)+1):
        s = th_memory1_second[i - 1]
        p = pr_memory1_second[i - 1]
        for j in range(0, int(no_of_subject.get())):
            list76.append(s[j].get())
            list77.append(p[j].get())
    q = 2
    for i in range(0, (int(no_of_subject.get()) * (int(total_no_of_students.get())-25)), (int(no_of_subject.get()))):
        if i == 0:
            th_marks_second.append(list76[0:int(no_of_subject.get())])
            pr_marks_second.append(list77[0:int(no_of_subject.get())])
        else:
            th_marks_second.append(list76[i:int(no_of_subject.get()) * q])
            pr_marks_second.append(list77[i:int(no_of_subject.get()) * q])
            q = q + 1
    done = False
    b = 0
    for i in range(1,(int(total_no_of_students.get())-25)+1):
        u=th_marks_second[i-1]
        fo=pr_marks_second[i-1]
        for j in range(1,int(no_of_subject.get())+1):
            if str(u[j-1]).isdigit()==True and str(fo[j-1]).isdigit()==True:
                if int(u[j-1])<=int(th_fm_list[j-1]) and int(fo[j-1])<=int(pr_fm_list[j-1]):
                    s678.cell(row=i+26,column=3*j+1).value=u[j-1]
                    s678.cell(row=i + 26, column=3 * j + 2).value = fo[j - 1]
                    as2 = int(s678.cell(row=i+26,column=3*j+1).value)+int(s678.cell(row=i + 26, column=3 * j + 2).value)
                    s678.cell(row=i + 26, column=3 * j + 3).value = as2
                    ld678.save(f"E:\\{classname}.xlsx")
                else:
                    b = b + 1
                    tmsg.showinfo("Error",f"obtained marks  exceeded full marks in {i+25}th row and {j}th subject ")
                    break
                    done=True
            else:
                b = b + 1
                tmsg.showinfo("Error", f"Marks should be a positive integer\n It's input field can not be left empty \nCheck {i+25}th row and {j}th subject")
                break
                done = True
        if done:
            break
    ld777 = op.load_workbook(f"E:\\{classname}.xlsx")
    s777 = ld777["Sheet"]
    if b==0:
        s777.cell(row=1,column=(3+int(no_of_subject.get())*3+1)).value="G.T"
        s777.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 2)).value = "PER"
        s777.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 3)).value = "F.GR"
        s777.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 4)).value = "F.GPA"
        ld777.save(f"E:\\{classname}.xlsx")
        th_sum=0
        pr_sum=0
        for i in range(0,len(th_fm_list)):
            th_sum+=int(th_fm_list[i])
        for i in range(0,len(pr_fm_list)):
            pr_sum+=int(pr_fm_list[i])
        full_marks=th_sum+pr_sum
        th_om_sum=0
        pr_om_sum=0
        for i in range(1, (int(total_no_of_students.get())-25) + 1):
            u11 = th_marks_second[i - 1]
            fo11 = pr_marks_second[i - 1]
            for j in range(1, int(no_of_subject.get()) + 1):
                th_om_sum+=int(u11[j-1])
                pr_om_sum+=int(fo11[j-1])
            gt_sum=th_om_sum+pr_om_sum
            s777.cell(row=i+26,column=3+int(no_of_subject.get())*3+1).value=gt_sum
            th_om_sum=0
            pr_om_sum=0
            percentage=round(gt_sum/full_marks*100,3)
            s777.cell(row=i + 26, column=3 + int(no_of_subject.get()) * 3 + 2).value = percentage

            ld777.save(f"E:\\{classname}.xlsx")
        global total_gpa_th1,total_grade_th1
        indi_gpa_th=[]
        total_gpa_th1=[]
        indi_grade_th=[]
        total_grade_th1=[]
        for i in range(1,(int(total_no_of_students.get())-25)+1):
            u12 = th_marks_second[i - 1]
            #f12 = pr_marks_third[i - 1]
            for j in range(1,int(no_of_subject.get())+1):
                if int(u12[j-1])==0:
                    gpa="-"
                    grade="-"
                    indi_gpa_th.append(gpa)
                    indi_grade_th.append(grade)
                else:
                    if float(u12[j-1])>=0.9*float(th_fm_list[j-1]) and float(u12[j-1])<=1.0*float(th_fm_list[j-1]):
                        gpa=4.0
                        grade="A+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j-1])>=0.8*float(th_fm_list[j-1]) and float(u12[j-1])<0.9*float(th_fm_list[j-1]):
                        gpa=3.6
                        grade="A"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j - 1]) >= 0.7 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.8 * float(
                            th_fm_list[j - 1]):
                        gpa = 3.2
                        grade = "B+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j - 1]) >= 0.6 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.7 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.8
                        grade = "B"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j - 1]) >= 0.5 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.6 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.4
                        grade = "C+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j - 1]) >= 0.4 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.5 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.0
                        grade = "C"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j - 1]) >= 0.3 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.4 * float(
                            th_fm_list[j - 1]):
                        gpa = 1.6
                        grade = "D+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j - 1]) >= 0.2 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.3 * float(
                            th_fm_list[j - 1]):
                        gpa = 1.2
                        grade = "D"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j - 1]) >= 0.1 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.2 * float(
                            th_fm_list[j - 1]):
                        gpa = 0.8
                        grade = "E+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u12[j - 1]) >= 0.0 * float(th_fm_list[j - 1]) and float(u12[j - 1]) < 0.1 * float(
                            th_fm_list[j - 1]):
                        gpa = 0.4
                        grade = "E"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)

            #total_gpa_th.append(indi_gpa_th)
            #total_grade_th.append(indi_grade_th)
        aq23=2
        for i in range(0,int(no_of_subject.get())*(int(total_no_of_students.get())-25),int(no_of_subject.get())):
            if i==0:
                total_gpa_th1.append(indi_gpa_th[0:int(no_of_subject.get())])
                total_grade_th1.append(indi_grade_th[0:int(no_of_subject.get())])
            else:
                total_gpa_th1.append(indi_gpa_th[i:int(no_of_subject.get())*aq23])
                total_grade_th1.append(indi_grade_th[i:int(no_of_subject.get())*aq23])
                aq23=aq23+1
        global total_gpa_pr1,total_grade_pr1
        indi_gpa_pr=[]
        indi_grade_pr=[]
        total_gpa_pr1=[]
        total_grade_pr1=[]
        for i in range(1,(int(total_no_of_students.get())-25)+1):
            fo12 = pr_marks_second[i - 1]
            for j in range(1,int(no_of_subject.get())+1):
                if int(fo12[j-1])==0:
                    gpa_pr="-"
                    grade_pr="-"
                    indi_gpa_pr.append(gpa_pr)
                    indi_grade_pr.append(grade_pr)
                else:
                    if float(fo12[j-1])>=0.9*float(pr_fm_list[j-1]) and float(fo12[j-1])<=1.0*float(pr_fm_list[j-1]):
                        gpa_pr = 4.0
                        grade_pr = "A+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j-1])>=0.8*float(pr_fm_list[j-1]) and float(fo12[j-1])<0.9*float(pr_fm_list[j-1]):
                        gpa_pr = 3.6
                        grade_pr = "A"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.7 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.8 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 3.2
                        grade_pr = "B+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.6 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.7 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.8
                        grade_pr = "B"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.5 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.6 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.4
                        grade_pr = "C+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.4 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.5 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.0
                        grade_pr = "C"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.3 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.4 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 1.6
                        grade_pr = "D+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.2 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.3 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 1.2
                        grade_pr = "D"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.1 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.2 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 0.8
                        grade_pr = "E+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo12[j - 1]) >= 0.0 * float(pr_fm_list[j - 1]) and float(fo12[j - 1]) < 0.1 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 0.4
                        grade_pr = "E"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)

            #total_gpa_th.append(indi_gpa_th)
            #total_grade_th.append(indi_grade_th)
        aq13=2
        for i in range(0,int(no_of_subject.get())*(int(total_no_of_students.get())-25),int(no_of_subject.get())):
            if i==0:
                total_gpa_pr1.append(indi_gpa_pr[0:int(no_of_subject.get())])
                total_grade_pr1.append(indi_grade_pr[0:int(no_of_subject.get())])
            else:
                total_gpa_pr1.append(indi_gpa_pr[i:int(no_of_subject.get())*aq13])
                total_grade_pr1.append(indi_grade_pr[i:int(no_of_subject.get())*aq13])
                aq13=aq13+1
        global total_gpa1,total_grade1
        indi_gpa = []
        indi_grade = []
        total_gpa1 = []
        total_grade1 = []
        for i in range(1,(int(total_no_of_students.get())-25)+1):
            fo12 = pr_marks_second[i - 1]
            u12 = th_marks_second[i - 1]
            for j in range(1, int(no_of_subject.get()) + 1):
                if int(fo12[j - 1])+int(u12[j-1]) == 0:
                    gpa = "-"
                    grade = "-"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                else:
                    if float(fo12[j-1])+float(u12[j-1])>=0.9*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<=1.0*float(total_marks_list1[j-1]):
                        gpa = 4.0
                        grade = "A+"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.8*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.9*float(total_marks_list1[j-1]):
                        gpa = 3.6
                        grade = "A"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.7*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.8*float(total_marks_list1[j-1]):
                        gpa = 3.2
                        grade = "B+"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.6*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.7*float(total_marks_list1[j-1]):
                        gpa = 2.8
                        grade = "B"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.5*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.6*float(total_marks_list1[j-1]):
                        gpa = 2.4
                        grade = "C+"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.4*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.5*float(total_marks_list1[j-1]):
                        gpa = 2.0
                        grade = "C"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.3*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.4*float(total_marks_list1[j-1]):
                        gpa = 1.6
                        grade = "D+"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.2*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.3*float(total_marks_list1[j-1]):
                        gpa = 1.2
                        grade = "D"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.1*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.2*float(total_marks_list1[j-1]):
                        gpa = 0.8
                        grade = "E+"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                    elif float(fo12[j-1])+float(u12[j-1])>=0.0*float(total_marks_list1[j-1]) and float(fo12[j-1])+float(u12[j-1])<0.1*float(total_marks_list1[j-1]):
                        gpa = 0.4
                        grade = "E"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
        aq233 = 2
        for i in range(0, int(no_of_subject.get()) * (int(total_no_of_students.get())-25), int(no_of_subject.get())):
            if i == 0:
                total_gpa1.append(indi_gpa[0:int(no_of_subject.get())])
                total_grade1.append(indi_grade[0:int(no_of_subject.get())])
            else:
                total_gpa1.append(indi_gpa[i:int(no_of_subject.get()) * aq233])
                total_grade1.append(indi_grade[i:int(no_of_subject.get()) * aq233])
                aq233 = aq233 + 1
        #print(total_gpa1)
        credit_sum1 = 0
        global gpa_list1
        gpa_list1 = []
        for i in range(1, len(credit_list) + 1):
            credit_sum1 += int(credit_list[i - 1])

        for i in range(1, (int(total_no_of_students.get())-25) + 1):
            xd1 = total_gpa1[i - 1]
            partial_gpa1 = 0
            for j in range(1, int(no_of_subject.get()) + 1):
                if type(xd1[j - 1]) == float:
                    partial_gpa1 = partial_gpa1 + (float(xd1[j - 1]) * float(credit_list[j - 1]))
            final_gpa = round(partial_gpa1 / credit_sum1, 2)
            gpa_list1.append(final_gpa)
            # partial_gpa = 0

        #print(gpa_list1)

        # print(len(gpa_list))
        global grade_list1
        grade_list1 = []
        for i in range(1, len(gpa_list1) + 1):
            if float(gpa_list1[i - 1]) >= 3.6 and float(gpa_list1[i - 1]) <= 4.0:
                final_grade = 'A+'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 3.2 and float(gpa_list1[i - 1]) < 3.6:
                final_grade = 'A'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 2.8 and float(gpa_list1[i - 1]) < 3.2:
                final_grade = 'B+'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 2.4 and float(gpa_list1[i - 1]) < 2.8:
                final_grade = 'B'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 2.0 and float(gpa_list1[i - 1]) < 2.4:
                final_grade = 'C+'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 1.6 and float(gpa_list1[i - 1]) < 2.0:
                final_grade = 'C'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 1.2 and float(gpa_list1[i - 1]) < 1.6:
                final_grade = 'D+'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 0.8 and float(gpa_list1[i - 1]) < 1.2:
                final_grade = 'D'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 0.4 and float(gpa_list1[i - 1]) < 0.8:
                final_grade = 'E+'
                grade_list1.append(final_grade)
            elif float(gpa_list1[i - 1]) >= 0.0 and float(gpa_list1[i - 1]) < 0.4:
                final_grade = 'E'
                grade_list1.append(final_grade)

        for i in range(1, len(gpa_list1) + 1):
            s777.cell(row=i + 26, column=3 + 3 * int(no_of_subject.get()) + 4).value = gpa_list1[i - 1]
            s777.cell(row=i + 26, column=3 + 3 * int(no_of_subject.get()) + 3).value = grade_list1[i - 1]
            ld777.save(f"E:\\{classname}.xlsx")
        global bu1
        f34 = Frame(f12,bg="cyan2")
        f34.place(x=665, y=600)
        bu1 = Button(f34, text="Proceed for marksheet", bg="bisque",fg="maroon", font="arial 12 bold", command=proceed_for_marksheet)
        bu1.grid(row=0, column=0)
        bu = Button(f34, text="Exit", bg="bisque",fg="maroon", font="arial 12 bold", command=homepage_once_again1)
        bu.grid(row=0, column=1,padx=10)



def add_remaining_data():
    global f13,th_memory1_second,pr_memory1_second

    #print(th_memory1)
    #print(th_memory1[1])
    list34=[]
    list35=[]
    ld458=op.load_workbook(f"E:\\{classname}.xlsx")
    s458=ld458["Sheet"]
    total_marks_list = []
    ld459=op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
    s459=ld459["Sheet"]
    for i in range(1, int(no_of_subject.get()) + 1):
        total_marks_list.append(int(s459.cell(row=i + 1, column=2).value) + int(s459.cell(row=i + 1, column=3).value))
    # print(total_marks_list)
    th_marks=[]
    pr_marks=[]
    for i in range(1,26):
        e = th_memory1[i-1]
        fo = pr_memory1[i-1]
        for j in range(0, int(no_of_subject.get())):
            list34.append(e[j].get())
            list35.append(fo[j].get())
            #print(e[j].get()


    a=2
    for i in range(0,(int(no_of_subject.get())*25),(int(no_of_subject.get()))):
        if i == 0:
            th_marks.append(list34[0:int(no_of_subject.get())])
            pr_marks.append(list35[0:int(no_of_subject.get())])
        else:
            th_marks.append(list34[i:int(no_of_subject.get()) * a])
            pr_marks.append(list35[i:int(no_of_subject.get()) * a])
            a = a + 1
    done=False
    z=0


    for i in range(1,26):
        g=th_marks[i-1]
        h=pr_marks[i-1]
        for j in range(1,int(no_of_subject.get())+1):
            if str(g[j-1]).isdigit()==True and str(h[j-1]).isdigit()==True:
                if int(g[j-1])<=int(th_fm_list[j-1]) and int(h[j-1])<=int(pr_fm_list[j-1]):
                    s458.cell(row=i+1,column=3*j+1).value=g[j-1]
                    s458.cell(row=i + 1, column=3 * j + 2).value = h[j - 1]
                    as1=int(s458.cell(row=i+1,column=3*j+1).value)+int(s458.cell(row=i + 1, column=3 * j + 2).value)
                    s458.cell(row=i+1,column=3*j+3).value=as1
                    ld458.save(f"E:\\{classname}.xlsx")


                else:
                    z = z + 1
                    tmsg.showinfo(f"Error",f"obtained marks  exceeded the full marks in {i}th row {j}th subject ")
                    break
                    done=True
            else:
                z = z + 1
                tmsg.showinfo("Error", f"Marks shoud be postive integer\n It's input field can not be left empty\ncheck {i}th row {j}th subject ")
                break
                done = True



        if done:
            break
    th_memory_second = []
    pr_memory_second = []
    th_memory1_second = []
    pr_memory1_second = []

    for i in range(1, ((int(total_no_of_students.get())-25) * int(no_of_subject.get())) + 1):
        th_memory_second.append(StringVar())
    for i in range(1, ((int(total_no_of_students.get())-25) * int(no_of_subject.get())) + 1):
        pr_memory_second.append(StringVar())
    p = 2
    for i in range(0, (int(no_of_subject.get()) *(int(total_no_of_students.get())-25)), (int(no_of_subject.get()))):
        if i == 0:
            th_memory1_second.append(th_memory_second[0:int(no_of_subject.get())])
            pr_memory1_second.append(pr_memory_second[0:int(no_of_subject.get())])
        else:
            th_memory1_second.append(th_memory_second[i:int(no_of_subject.get()) * p])
            pr_memory1_second.append(pr_memory_second[i:int(no_of_subject.get()) * p])
            p = p + 1

    if z==0:
        f13=Frame(f12,bg="cyan2")
        f13.place(x=0,y=0,width=2500,height=1500)

        labu = Label(f13, text='S\nNo.', font='arial 10 bold',bg="cyan2")
        labu.grid(row=1, column=40)
        for i in range(1,(int(total_no_of_students.get())-25)+1):
            labu1 = Label(f13, text=25+i, font='arial 10 bold',bg="cyan2")
            labu1.grid(row=i + 1, column=40)

        for i in range(1,(int(total_no_of_students.get())-25)+1):
            lab=Label(f13,text=s458.cell(row=26+i,column=1).value,font="arial 10 bold",bg="cyan2")
            lab.grid(row=i+1,column=0)
        for i in range(0, int(no_of_subject.get()) ):
            sub_label = Label(f13, text=f'{subject_name_list[i]}\n(TH)'.upper(), font='arial 9 bold',bg="cyan2")
            sub_label.grid(row=1, column=2 * (i+1))
            sub_label = Label(f13, text=f'{subject_name_list[i]}\n(PR)'.upper(), font='arial 9 bold',bg="cyan2")
            sub_label.grid(row=1, column=(2 * (i + 1))+1)


        for i in range(0, (int(total_no_of_students.get())-25)):
            c = th_memory1_second[i]
            d = pr_memory1_second[i]
            for j in range(0, int(no_of_subject.get())):
                theory_marks_entry = Entry(f13, textvariable=c[j], width=3, border=2, relief=SUNKEN, font="arial 10 bold")
                theory_marks_entry.grid(row= i+2, column=2 * (j + 1))
                practical_marks_entry = Entry(f13, textvariable=d[j], width=3, border=2, relief=SUNKEN,
                                              font="arial 10 bold")
                practical_marks_entry.grid(row= i+2, column=(2 * (j + 1)) + 1)
                c[j].set(0)
                d[j].set(0)
        f14=Frame(f13,bg="cyan2")
        f14.place(x=0,y=600)
        print_marks_button=Button(f14,text="Save and Continue",font="arial 12 bold",bg="bisque",fg="maroon",command=print_marks)
        print_marks_button.grid(row=0,column=1)
        back_again_button = Button(f14, text="Back to Previous Page", font="arial 12 bold", bg="bisque",fg="maroon",command=back_once_again)
        back_again_button .grid(row=0, column=0,padx=5)
        full_marks_button2 = Button(f14, text="View Full Marks of Subjects", font="arial 12 bold", bg="bisque",fg="maroon",
                                   command=full_marks)
        full_marks_button2.grid(row=0, column=2, padx=10)
        if s458.cell(row=27, column=4).value is not None:
            for i in range(0, int(total_no_of_students.get())-25):
                c9 = th_memory1_second[i]
                d9 = pr_memory1_second[i]
                for j in range(1, int(no_of_subject.get()) + 1):
                    c9[j - 1].set(s458.cell(row=i + 27, column=3 * j + 1).value)
                    d9[j - 1].set(s458.cell(row=i + 27, column=3 * j + 2).value)
        tmsg.showinfo("message", "Simply write '0' in place of None for the added no. of students and added no. of subjects")


        th_sum=0
        pr_sum=0
        ld321=op.load_workbook(f"E:\\{classname}.xlsx")
        s321=ld321["Sheet"]
        s321.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 1)).value = "GRAND TOTAL"
        s321.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 2)).value = "PERCENTAGE"
        s321.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 3)).value = "FINAL GRADE"
        s321.cell(row=1, column=(3 + int(no_of_subject.get()) * 3 + 4)).value = "FINAL GPA"
        ld321.save(f"E:\\{classname}.xlsx")
        th_om_sum = 0
        pr_om_sum = 0
        for i in range(0,len(th_fm_list)):
            th_sum+=int(th_fm_list[i])
        for i in range(0,len(pr_fm_list)):
            pr_sum+=int(pr_fm_list[i])
        full_marks1=th_sum+pr_sum

        for i in range(1, 25 + 1):
            u11 = th_marks[i - 1]
            fo11 = pr_marks[i - 1]
            for j in range(1, int(no_of_subject.get()) + 1):
                th_om_sum+=int(u11[j-1])
                pr_om_sum+=int(fo11[j-1])
            gt_sum=th_om_sum+pr_om_sum
            s321.cell(row=i+1,column=3+int(no_of_subject.get())*3+1).value=gt_sum
            th_om_sum=0
            pr_om_sum=0
            percentage=round(gt_sum/full_marks1*100,3)
            s321.cell(row=i + 1, column=3 + int(no_of_subject.get()) * 3 + 2).value = percentage

            ld321.save(f"E:\\{classname}.xlsx")
        global total_gpa_th,total_grade_th
        indi_gpa_th = []
        total_gpa_th = []
        indi_grade_th = []
        total_grade_th = []
        for i in range(1, 25 + 1):
            u13 = th_marks[i - 1]
            # f12 = pr_marks_third[i - 1]
            for j in range(1, int(no_of_subject.get()) + 1):
                if int(u13[j - 1]) == 0:
                    gpa = "-"
                    grade = "-"
                    indi_gpa_th.append(gpa)
                    indi_grade_th.append(grade)
                else:
                    if float(u13[j - 1]) >= 0.9 * float(th_fm_list[j - 1]) and float(u13[j - 1]) <= 1.0 * float(
                            th_fm_list[j - 1]):
                        gpa = 4.0
                        grade = "A+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.8 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.9 * float(
                            th_fm_list[j - 1]):
                        gpa = 3.6
                        grade = "A"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.7 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.8 * float(
                            th_fm_list[j - 1]):
                        gpa = 3.2
                        grade = "B+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.6 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.7 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.8
                        grade = "B"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.5 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.6 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.4
                        grade = "C+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.4 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.5 * float(
                            th_fm_list[j - 1]):
                        gpa = 2.0
                        grade = "C"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.3 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.4 * float(
                            th_fm_list[j - 1]):
                        gpa = 1.6
                        grade = "D+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.2 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.3 * float(
                            th_fm_list[j - 1]):
                        gpa = 1.2
                        grade = "D"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.1 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.2 * float(
                            th_fm_list[j - 1]):
                        gpa = 0.8
                        grade = "E+"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
                    elif float(u13[j - 1]) >= 0.0 * float(th_fm_list[j - 1]) and float(u13[j - 1]) < 0.1 * float(
                            th_fm_list[j - 1]):
                        gpa = 0.4
                        grade = "E"
                        indi_gpa_th.append(gpa)
                        indi_grade_th.append(grade)
        aq = 2
        for i in range(0, int(no_of_subject.get()) * 25, int(no_of_subject.get())):
            if i == 0:
                total_gpa_th.append(indi_gpa_th[0:int(no_of_subject.get())])
                total_grade_th.append(indi_grade_th[0:int(no_of_subject.get())])
            else:
                total_gpa_th.append(indi_gpa_th[i:int(no_of_subject.get()) * aq])
                total_grade_th.append(indi_grade_th[i:int(no_of_subject.get()) * aq])
                aq = aq + 1
        global total_gpa_pr,total_grade_pr
        indi_gpa_pr = []
        indi_grade_pr = []
        total_gpa_pr = []
        total_grade_pr = []
        for i in range(1, 25 + 1):
            fo13 = pr_marks[i - 1]
            for j in range(1, int(no_of_subject.get()) + 1):
                if int(fo13[j - 1]) == 0:
                    gpa_pr = "-"
                    grade_pr = "-"
                    indi_gpa_pr.append(gpa_pr)
                    indi_grade_pr.append(grade_pr)
                else:
                    if float(fo13[j - 1]) >= 0.9 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) <= 1.0 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 4.0
                        grade_pr = "A+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.8 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.9 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 3.6
                        grade_pr = "A"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.7 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.8 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 3.2
                        grade_pr = "B+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.6 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.7 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.8
                        grade_pr = "B"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.5 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.6 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.4
                        grade_pr = "C+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.4 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.5 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 2.0
                        grade_pr = "C"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.3 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.4 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 1.6
                        grade_pr = "D+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.2 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.3 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 1.2
                        grade_pr = "D"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.1 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.2 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 0.8
                        grade_pr = "E+"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)
                    elif float(fo13[j - 1]) >= 0.0 * float(pr_fm_list[j - 1]) and float(fo13[j - 1]) < 0.1 * float(
                            pr_fm_list[j - 1]):
                        gpa_pr = 0.4
                        grade_pr = "E"
                        indi_gpa_pr.append(gpa_pr)
                        indi_grade_pr.append(grade_pr)

            # total_gpa_th.append(indi_gpa_th)
            # total_grade_th.append(indi_grade_th)
        aq12 = 2
        for i in range(0, int(no_of_subject.get()) * 25, int(no_of_subject.get())):
            if i == 0:
                total_gpa_pr.append(indi_gpa_pr[0:int(no_of_subject.get())])
                total_grade_pr.append(indi_grade_pr[0:int(no_of_subject.get())])
            else:
                total_gpa_pr.append(indi_gpa_pr[i:int(no_of_subject.get()) * aq12])
                total_grade_pr.append(indi_grade_pr[i:int(no_of_subject.get()) * aq12])
                aq12 = aq12 + 1

        global total_gpa,total_grade
        indi_gpa = []
        indi_grade = []
        total_gpa = []
        total_grade = []

        #print(total_marks_list)
        for i in range(1, 25 + 1):
            u14 = th_marks[i - 1]
            fo14 = pr_marks[i - 1]

            for j in range(1, int(no_of_subject.get()) + 1):
                if float(fo14[j - 1]) + float(u14[j - 1]) >= 0.9 * float(total_marks_list[j - 1]) and float(
                            fo14[j - 1]) + float(u14[j - 1]) <= 1.0 * float(total_marks_list[j - 1]):
                        gpa = 4.0
                        grade = "A+"
                        indi_gpa.append(gpa)
                        indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.8 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.9 * float(total_marks_list[j - 1]):
                    gpa = 3.6
                    grade = "A"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.7 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.8 * float(total_marks_list[j - 1]):
                    gpa = 3.2
                    grade = "B+"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.6 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.7 * float(total_marks_list[j - 1]):
                    gpa = 2.8
                    grade = "B"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.5 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.6 * float(total_marks_list[j - 1]):
                    gpa = 2.4
                    grade = "C+"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.4 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.5 * float(total_marks_list[j - 1]):
                    gpa = 2.0
                    grade = "C"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.3 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.4 * float(total_marks_list[j - 1]):
                    gpa = 1.6
                    grade = "D+"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.2 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.3 * float(total_marks_list[j - 1]):
                    gpa = 1.2
                    grade = "D"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.1 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.2 * float(total_marks_list[j - 1]):
                    gpa = 0.8
                    grade = "E+"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
                elif float(fo14[j - 1]) + float(u14[j - 1]) >= 0.01 * float(total_marks_list[j - 1]) and float(
                        fo14[j - 1]) + float(u14[j - 1]) < 0.1 * float(total_marks_list[j - 1]):
                    gpa = 0.4
                    grade = "E"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)

                else:
                    gpa = "-"
                    grade = "-"
                    indi_gpa.append(gpa)
                    indi_grade.append(grade)
        aq22 = 2

        for i in range(0, int(no_of_subject.get()) * 25, int(no_of_subject.get())):
            if i == 0:
                total_gpa.append(indi_gpa[0:int(no_of_subject.get())])
                total_grade.append(indi_grade[0:int(no_of_subject.get())])
            else:
                total_gpa.append(indi_gpa[i:int(no_of_subject.get()) * aq22])
                total_grade.append(indi_grade[i:int(no_of_subject.get()) * aq22])
                aq22 = aq22 + 1
        #print(total_gpa)
        #print(len(total_gpa))

        credit_sum=0
        gpa_list=[]
        for i in range(1,len(credit_list)+1):
            credit_sum+=int(credit_list[i-1])

        for i in range(1,25+1):
            xd=total_gpa[i-1]
            partial_gpa = 0
            for j in range(1,int(no_of_subject.get())+1):
                if type(xd[j-1])==float:
                    partial_gpa=partial_gpa+(float(xd[j-1])*int(credit_list[j-1]))
            final_gpa=round(partial_gpa/credit_sum,2)
            gpa_list.append(final_gpa)
            #partial_gpa = 0


        #print(gpa_list)
        #print(len(gpa_list))
        grade_list=[]
        for i in range(1,len(gpa_list)+1):
            if float(gpa_list[i-1])>=3.6 and float(gpa_list[i-1])<=4.0:
                final_grade='A+'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=3.2 and float(gpa_list[i-1])<3.6:
                final_grade='A'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=2.8 and float(gpa_list[i-1])<3.2:
                final_grade='B+'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=2.4 and float(gpa_list[i-1])<2.8:
                final_grade='B'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=2.0 and float(gpa_list[i-1])<2.4:
                final_grade='C+'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=1.6 and float(gpa_list[i-1])<2.0:
                final_grade='C'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=1.2 and float(gpa_list[i-1])<1.6:
                final_grade='D+'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=0.8 and float(gpa_list[i-1])<1.2:
                final_grade='D'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=0.4 and float(gpa_list[i-1])<0.8:
                final_grade='E+'
                grade_list.append(final_grade)
            elif float(gpa_list[i-1])>=0.0 and float(gpa_list[i-1])<0.4:
                final_grade='E'
                grade_list.append(final_grade)

        for i in range(1,len(gpa_list)+1):
            s321.cell(row=i+1,column=3+3*int(no_of_subject.get())+4).value=gpa_list[i-1]
            s321.cell(row=i+1,column=3+3*int(no_of_subject.get())+3).value=grade_list[i-1]
            ld321.save(f"E:\\{classname}.xlsx")




def enter_marks():
    global f12, th_om_dict,pr_om_dict,memory_list_th,c,d,th_memory1,pr_memory1,th_memory1_third,pr_memory1_third

    ld76 = op.load_workbook(f"E:\\{classname}.xlsx")
    s76 = ld76["Sheet"]


    #row_no=s76.max_row
    d=0
    for i in range(1,int(total_no_of_students.get())+1):
        if len(student_name_dict[i].get())!=0 and len(student_name_dict[i].get())<=25 and len(roll_dict[i].get())!=0 and len(address_dict[i].get())!=0:
            s76.cell(row=i + 1, column=1).value = student_name_dict[i].get().upper()
            s76.cell(row=i + 1, column=2).value = roll_dict[i].get().upper()
            s76.cell(row=i + 1, column=3).value = address_dict[i].get().upper()
            ld76.save(f"E:\\{classname}.xlsx")

        else:
            d=d+1
            tmsg.showinfo("Error","Input field can not be empty\nname of the student can not exceed 25 letters")
            break
    if d==0:
        memory_list_th=[]

        def back_again():
            f12.destroy()
            #bu.destroy()
        if int(total_no_of_students.get())<=25:
            f12=Frame(f9,bg="cyan2")
            f12.place(x=0,y=0,width=2500,height=1500)
            ld4567=op.load_workbook(f'E:\\{classname}.xlsx')
            s4567=ld4567["Sheet"]
            labu=Label(f12,text='S\nNo.',font='arial 10 bold',bg="cyan2")
            labu.grid(row=1,column=40)
            for i in range(1,int(total_no_of_students.get())+1):
                labu1 = Label(f12, text=i, font='arial 10 bold',bg="cyan2")
                labu1.grid(row=i + 1, column=40)
            for i in range(1,int(total_no_of_students.get())+1):
                student_name_label=Label(f12,text=s4567.cell(row=i+1,column=1).value,font='arial 10 bold',bg="cyan2")
                student_name_label.grid(row=i+1,column=0)
            for i in range(1,int(no_of_subject.get())+1):
                sub_label=Label(f12,text=f'{subject_name_list[i-1]}\n(TH)'.upper(),font='arial 9 bold',bg="cyan2")
                sub_label.grid(row=1,column=2*i)
                sub_label = Label(f12, text=f'{subject_name_list[i - 1]}\n(PR)'.upper(),font='arial 9 bold',bg="cyan2")
                sub_label.grid(row=1, column=2*i+1)


            th_memory_third=[]
            pr_memory_third = []
            th_memory1_third=[]
            pr_memory1_third = []

            for i in range(1,(int(total_no_of_students.get())*int(no_of_subject.get()))+1):
                th_memory_third.append(StringVar())
            for i in range(1, (int(total_no_of_students.get()) * int(no_of_subject.get())) + 1):
                pr_memory_third.append(StringVar())
            r=2
            for i in range(0,(int(no_of_subject.get())*int(total_no_of_students.get())),(int(no_of_subject.get()))):
                if i==0:
                    th_memory1_third.append(th_memory_third[0:int(no_of_subject.get())])
                    pr_memory1_third.append(pr_memory_third[0:int(no_of_subject.get())])
                else:
                    th_memory1_third.append(th_memory_third[i:int(no_of_subject.get())*r])
                    pr_memory1_third.append(pr_memory_third[i:int(no_of_subject.get())*r])
                    r=r+1






            for i in range(0,int(total_no_of_students.get())):
                m=th_memory1_third[i]
                n=pr_memory1_third[i]
                for j in range(0,int(no_of_subject.get())):
                    theory_marks_entry=Entry(f12,textvariable=m[j],width=3,border=2,relief=SUNKEN,font="arial 10 bold")
                    theory_marks_entry.grid(row=2+i,column=2*(j+1))
                    practical_marks_entry = Entry(f12, textvariable=n[j], width=3, border=2, relief=SUNKEN,font="arial 10 bold")
                    practical_marks_entry.grid(row=2 + i, column=(2 * (j + 1))+1)
                    m[j].set(0)
                    n[j].set(0)

            global f13
            f13 = Frame(f12,bg="cyan2")
            f13.place(x=0, y=640)
            back_again1 = Button(f13, text="Back to previous page", font="arial 12 bold", bg="bisque",fg="maroon",
                                 command=back_again)
            back_again1.grid(row=0, column=0, padx=10)
            print_marks_button1 = Button(f13, text="Save and Continue", font="arial 12 bold", bg="bisque",fg="maroon",command=print_marks1)
            print_marks_button1.grid(row=0, column=1)
            full_marks_button1 = Button(f13, text="View Full Marks of Subjects", font="arial 12 bold", bg="bisque",fg="maroon",
                                       command=full_marks)
            full_marks_button1.grid(row=0, column=2, padx=10)
            if s76.cell(row=2, column=4).value is not None:
                for i in range(0,int(total_no_of_students.get())):
                    c0 = th_memory1_third[i]
                    d0 = pr_memory1_third[i]
                    for j in range(1, int(no_of_subject.get())+1):
                        c0[j-1].set(s76.cell(row=i+2,column=3*j+1).value)
                        d0[j-1].set(s76.cell(row=i + 2, column=3 * j + 2).value)
            tmsg.showinfo("message", "Simply write '0' in place of None for the added no. of students and added no. of subjects")





        elif int(total_no_of_students.get()) > 25 and int(total_no_of_students.get())<=45:
            f12 = Frame(f9,bg="cyan2")
            f12.place(x=0, y=0, width=2500, height=1500)
            ld4567 = op.load_workbook(f'E:\\{classname}.xlsx')
            s4567 = ld4567["Sheet"]
            labu = Label(f12, text='S\nNo.', font='arial 10 bold',bg="cyan2")
            labu.grid(row=1, column=40)
            for i in range(1,26):
                labu1=Label(f12,text=i, font='arial 10 bold',bg="cyan2")
                labu1.grid(row=i+1,column=40)
            for i in range(1,26):
                student_name_label=Label(f12,text=s4567.cell(row=i+1,column=1).value,font='arial 10 bold',bg="cyan2")
                student_name_label.grid(row=i+1,column=0)
            for i in range(1,int(no_of_subject.get())+1):
                sub_label=Label(f12,text=f'{subject_name_list[i-1]}\n(TH)'.upper(),font='arial 9 bold',bg="cyan2")
                sub_label.grid(row=1,column=2*i)
                sub_label = Label(f12, text=f'{subject_name_list[i - 1]}\n(PR)'.upper(),font='arial 9 bold',bg="cyan2")
                sub_label.grid(row=1, column=2*i+1)


            th_memory=[]
            pr_memory = []
            th_memory1=[]
            pr_memory1 = []

            for i in range(1,(25*int(no_of_subject.get()))+1):
                th_memory.append(StringVar())
            for i in range(1, (25 * int(no_of_subject.get())) + 1):
                pr_memory.append(StringVar())
            a=2
            for i in range(0,(int(no_of_subject.get())*25),(int(no_of_subject.get()))):
                if i==0:
                    th_memory1.append(th_memory[0:int(no_of_subject.get())])
                    pr_memory1.append(pr_memory[0:int(no_of_subject.get())])
                else:
                    th_memory1.append(th_memory[i:int(no_of_subject.get())*a])
                    pr_memory1.append(pr_memory[i:int(no_of_subject.get())*a])
                    a=a+1






            for i in range(0,25):
                c=th_memory1[i]
                d=pr_memory1[i]
                for j in range(0,int(no_of_subject.get())):
                    theory_marks_entry=Entry(f12,textvariable=c[j],width=3,border=2,relief=SUNKEN,font="arial 10 bold")
                    theory_marks_entry.grid(row=2+i,column=2*(j+1))
                    practical_marks_entry = Entry(f12, textvariable=d[j], width=3, border=2, relief=SUNKEN,font="arial 10 bold")
                    practical_marks_entry.grid(row=2 + i, column=(2 * (j + 1))+1)
                    c[j].set(0)
                    d[j].set(0)
            #print(len(th_memory1[1]))





            f13 = Frame(f12,bg="cyan2")
            f13.place(x=0, y=640)
            remaining_button=Button(f13,text="Save and Add Remaining data",font="arial 12 bold",bg="bisque",fg="maroon",command=add_remaining_data)
            remaining_button.grid(row=0,column=1)

            #print(memory_list_th)
            #print(len(memory_list_th))



            back_again1=Button(f13,text="Back to previous page",font="arial 12 bold",bg="bisque",fg="maroon",command=back_again)
            back_again1.grid(row=0, column=0,padx=10)
            full_marks_button=Button(f13,text="View Full Marks of Subjects",font="arial 12 bold",bg="bisque",fg="maroon",command=full_marks)
            full_marks_button.grid(row=0, column=2,padx=10)
            if s4567.cell(row=2, column=4).value is not None:
                for i in range(0,25):
                    c10 = th_memory1[i]
                    d10 = pr_memory1[i]
                    for j in range(1, int(no_of_subject.get())+1):
                        c10[j-1].set(s4567.cell(row=i+2,column=3*j+1).value)
                        d10[j-1].set(s4567.cell(row=i + 2, column=3 * j + 2).value)
            tmsg.showinfo("message", "Simply write '0' in place of None for the added no. of students and added no. of subjects")







    """if row_no==int(total_no_of_students.get())+1:
        f11=Frame(f4)
        f11.place(x=0,y=0,width=2500,height=1500)
        l11=Label(f11,text="hello")
        l11.pack()"""





def back():
    f4.destroy()


def previous():
    # f7.destroy()
   f9.destroy()
def Homepage():
    f9.destroy()
    f4.destroy()
def clear():
    yesno23=tmsg.askyesno("Confirmation","Do you want to clear all the entered data?\nTHINK TWICE\nPress YES to clear all the data")
    if yesno23==True:
        yesno24=tmsg.askyesno("Confirmation","Do you really want to clear all the data\n press YES to clear all the data" )
        if yesno24==True:
            os.remove(f"E:\\{classname}.xlsx")
            if os.path.exists(f"E:\\{classname}_full_marks.xlsx")==True:
                os.remove(f"E:\\{classname}_full_marks.xlsx")
            #os.remove(f"E:\\{classname}_full_marks.xlsx")
            wb2=op.Workbook(f"E:\\{classname}.xlsx")
            wb2.save(f"E:\\{classname}.xlsx")
            ld11=op.load_workbook(f"E:\\{classname}.xlsx")
            z = ld11['Sheet']
            z.cell(row=1, column=1).value = "NAME"
            z.cell(row=1, column=2).value = "ROLL NO."
            z.cell(row=1, column=3).value = "ADDRESS"
            ld11.save(f"E:\\{classname}.xlsx")

            ld12 = op.load_workbook(f"E:\\{classname}.xlsx")
            ws = ld12.active
            ws.column_dimensions['A'].width=25
            list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
                    'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM',
                    'AN', 'AO', 'AP']
            for i in range(2, len(list) + 1):
                ws.column_dimensions[list[i - 1]].width = 18
            ld12.save(f"E:\\{classname}.xlsx")
            if os.path.exists(f"E:{classname}_with_rank.xlsx")==True:
                os.remove(f"E:{classname}_with_rank.xlsx")

            f4.destroy()


def student_details():

    global f9,first_half,rem_half,student_name_dict,roll_dict,address_dict

    ld32 = op.load_workbook(f"E:\\{classname}.xlsx")
    s32 = ld32["Sheet"]
    row_no = s32.max_row
    f9 = Frame(f4,bg="cyan2")
    f9.place(x=0, y=0,width=2500,height=1500)
    sn_label = Label(f9, text="S No.", font="arial 14 bold",bg="cyan2")
    sn_label.grid(row=0, column=0)
    name_label=Label(f9,text="Student's Name",font="arial 14 bold",bg="cyan2")
    name_label.grid(row=0,column=1)
    roll_label = Label(f9, text="Roll no.", font="arial 14 bold",bg="cyan2")
    roll_label.grid(row=0, column=2)
    address_label = Label(f9, text="Address", font="arial 14 bold",bg="cyan2")
    address_label.grid(row=0, column=3)
    sn_label1 = Label(f9, text="S No.", font="arial 14 bold",bg="cyan2")
    sn_label1.grid(row=0, column=4)
    name_label1 = Label(f9, text="Student's Name", font="arial 14 bold",bg="cyan2")
    name_label1.grid(row=0, column=5)
    roll_label1 = Label(f9, text="Roll no.", font="arial 14 bold",bg="cyan2")
    roll_label1.grid(row=0, column=6)
    address_label1 = Label(f9, text="Address", font="arial 14 bold",bg="cyan2")
    address_label1.grid(row=0, column=7)
    student_name_dict={i:StringVar() for i in range(1,int(total_no_of_students.get())+1)}
    roll_dict = {i: StringVar() for i in range(1, int(total_no_of_students.get()) + 1)}
    address_dict = {i: StringVar() for i in range(1, int(total_no_of_students.get()) + 1)}

    if int(total_no_of_students.get())>22:
        """if int(total_no_of_students.get())==1:
            first_half=1
            rem_half=0
            sn_label1.destroy()
            name_label1.destroy()
            roll_label1.destroy()
            address_label1.destroy()
        else:
            first_half=int(total_no_of_students.get())//2
            rem_half=int(total_no_of_students.get())-first_half"""

        for i in range(1,23):
            sn_value=Label(f9,text=i,font="arial 12 bold",bg="cyan2")
            sn_value.grid(row=i, column=0)
            student_name_entry=Entry(f9,textvariable=student_name_dict[i],font="arial 12 bold",border=4,relief=SUNKEN)
            student_name_entry.grid(row=i,column=1)
            roll_entry = Entry(f9, textvariable=roll_dict[i], font="arial 12 bold",
                                       border=4, relief=SUNKEN)
            roll_entry.grid(row=i, column=2)
            address_entry = Entry(f9, textvariable=address_dict[i], font="arial 12 bold",
                               border=4, relief=SUNKEN)
            address_entry.grid(row=i, column=3)
            address_dict[i].set("NA")

            student_name_dict[i].set(s32.cell(row=i+1,column=1).value)
            roll_dict[i].set(s32.cell(row=i + 1, column=2).value)
            if s32.cell(row=i + 1, column=3).value == None or s32.cell(row=i + 1, column=3).value == "None" or s32.cell(
                    row=i + 1, column=3).value == "NONE":
                address_dict[i].set("SINDHUPALCHOWK,NEPAL")
            else:
                address_dict[i].set(s32.cell(row=i + 1, column=3).value)





        for i in range(1,(int(total_no_of_students.get())-22)+1):
            sn_value1 = Label(f9, text=22+i, font="arial 12 bold",bg="cyan2")
            sn_value1.grid(row=i, column=4)
            student_name1_entry = Entry(f9, textvariable=student_name_dict[22+i], font="arial 12 bold",
                                       border=4, relief=SUNKEN)
            student_name1_entry.grid(row=i, column=5)
            roll_entry = Entry(f9, textvariable=roll_dict[22+i], font="arial 12 bold",
                               border=4, relief=SUNKEN)
            roll_entry.grid(row=i, column=6)
            address_entry = Entry(f9, textvariable=address_dict[22+i], font="arial 12 bold",
                                  border=4, relief=SUNKEN)
            address_entry.grid(row=i, column=7)
            address_dict[22+i].set("NA")

            student_name_dict[22+i].set(s32.cell(row=i+23,column=1).value)
            roll_dict[22+i].set(s32.cell(row=i + 23, column=2).value)
            if s32.cell(row=i + 23, column=3).value=="None" or s32.cell(row=i + 23, column=3).value==None or s32.cell(row=i + 23, column=3).value=="NONE":
                address_dict[22+i].set("SINDHUPALCHOWK,NEPAL")
            else:
                address_dict[22 + i].set(s32.cell(row=i + 23, column=3).value)
    elif int(total_no_of_students.get()) <= 22:
        if int(total_no_of_students.get())==1:
                    first_half=1

                    sn_label1.destroy()
                    name_label1.destroy()
                    roll_label1.destroy()
                    address_label1.destroy()
        else:
            first_half=int(total_no_of_students.get())
        sn_label1.destroy()
        name_label1.destroy()
        roll_label1.destroy()
        address_label1.destroy()


        for i in range(1, first_half+1):
            sn_value = Label(f9, text=i, font="arial 12 bold",bg="cyan2")
            sn_value.grid(row=i, column=0)
            student_name_entry = Entry(f9, textvariable=student_name_dict[i], font="arial 12 bold", border=4,
                                       relief=SUNKEN)
            student_name_entry.grid(row=i, column=1)
            roll_entry = Entry(f9, textvariable=roll_dict[i], font="arial 12 bold",
                               border=4, relief=SUNKEN)
            roll_entry.grid(row=i, column=2)
            address_entry = Entry(f9, textvariable=address_dict[i], font="arial 12 bold",
                                  border=4, relief=SUNKEN)
            address_entry.grid(row=i, column=3)
            address_dict[i].set("NA")

            student_name_dict[i].set(s32.cell(row=i+1,column=1).value)
            roll_dict[i].set(s32.cell(row=i+1 , column=2).value)
            if s32.cell(row=i + 1, column=3).value == None or s32.cell(row=i + 1, column=3).value == "None" or s32.cell(row=i + 1, column=3).value == "NONE":
                address_dict[i].set("SINDHUPALCHOWK,NEPAL")
            else:
                address_dict[i].set(s32.cell(row=i + 1, column=3).value)



    previous_button = Button(f9, text="Back to previous page", font="arial 12 bold",
                             bg="bisque",fg="maroon", command=previous)
    previous_button.grid(row=23, column=1)

    marks_button1=Button(f9,text="Save and Enter Marks", font="arial 12 bold",
                             bg="bisque",fg="maroon", command=enter_marks)
    marks_button1.grid(row=23,column=2)
    homepage_button = Button(f9, text="Homepage", font="arial 12 bold",
                           bg="bisque",fg="maroon", command=Homepage)
    homepage_button.grid(row=23, column=3)













def save():

    global subject_name_list,th_fm_list,pr_fm_list,credit_list,credit_list,student_details_button,pr_fm_list,subject_name_list,th_fm_list
    subject_name_list = []
    th_fm_list = []
    pr_fm_list = []
    credit_list=[]
    for i in range(1,int(no_of_subject.get())+1):
        if len(subject_dict[i].get())!=0 and len(th_fm_dict[i].get())!=0 and len(pr_fm_dict[i].get())!=0 and len(credit_dict[i].get())!=0:
            if len(subject_dict[i].get())<=9:
                if th_fm_dict[i].get().isdigit()==True and int(th_fm_dict[i].get())>0 and int(th_fm_dict[i].get())<=100:
                    if pr_fm_dict[i].get().isdigit() == True and int(pr_fm_dict[i].get()) >= 0 and int(pr_fm_dict[i].get()) <= 50:
                        if int(th_fm_dict[i].get())+int(pr_fm_dict[i].get())<=100:
                            if int(credit_dict[i].get())<=4 and credit_dict[i].get().isdigit()==True and int(credit_dict[i].get())>=1:
                                subject_name_list.append(subject_dict[i].get())
                                th_fm_list.append(th_fm_dict[i].get())
                                pr_fm_list.append(pr_fm_dict[i].get())
                                credit_list.append(credit_dict[i].get())
                                ld1=op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
                                s1=ld1['Sheet']
                                s1.cell(row=1,column=1).value="Subject Name"
                                s1.cell(row=1, column=2).value = "full marks(th)"
                                s1.cell(row=1, column=3).value = "full marks(pr)"
                                s1.cell(row=i+1, column=1).value=subject_dict[i].get()
                                s1.cell(row=i + 1, column=2).value = th_fm_dict[i].get()
                                s1.cell(row=i + 1, column=3).value = pr_fm_dict[i].get()
                                s1.cell(row=i+1,column=4).value=credit_dict[i].get()
                                ld1.save(f"E:\\{classname}_full_marks.xlsx")
                                ld4=op.load_workbook(f"E:\\{classname}.xlsx")
                                s4=ld4["Sheet"]
                                p=subject_dict[i].get()
                                s4.cell(row=1,column=1+3*i).value=f"{p[0:3].upper()}(TH)"
                                s4.cell(row=1, column=2 + 3 * i).value = f"{p[0:3].upper()}(PR)"
                                s4.cell(row=1, column=3 + 3 * i).value = f"{p[0:3].upper()}(T)"
                                ld4.save(f"E:\\{classname}.xlsx")
                            else:
                                tmsg.showinfo('error',"credit hour should be a positive integer ranging from 1 to 4")



                        else:
                            tmsg.showinfo("Error","The sum of full marks of theory and practical can not ecxeed 100")
                    else:
                        tmsg.showinfo("Error","full marks of practical should be a positive integer less than or equal to 50 ")
                        break
                else:
                    tmsg.showinfo("Error","full marks of theory should be a positive integer greater than 0 and less than or equal to 100 ")
                    break
            else:
                tmsg.showinfo("Error","no. of letters in subject name can not exceed 9 letters")
                break
        else:
            tmsg.showinfo("Error","Input field can not be empty")
            break
    list1=[]
    list2=[]
    list3=[]
    list4=[]
    c=0
    if os.path.exists(f"E:\\{classname}_full_marks.xlsx"):
        ld45 = op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
        s45 = ld45["Sheet"]
        for i in range(1,int(no_of_subject.get())+1):
            list1.append(s45.cell(row=i+1,column=1).value)
            list2.append(s45.cell(row=i + 1, column=2).value)
            list3.append(s45.cell(row=i + 1, column=3).value)
            list4.append(s45.cell(row=i + 1, column=4).value)
    #print(list1)
    for i in range(1,len(list1)+1):
        if list1[i-1] is None or len(list1[i-1])==0 or subject_dict[i].get()!=list1[i-1] or th_fm_dict[i].get()!=list2[i-1] or pr_fm_dict[i].get()!=list3[i-1] or credit_dict[i].get()!=list4[i-1]    :
            c=c+1




    def direct_marks_entry():
        ld8989=op.load_workbook(f"E:{classname}.xlsx")
        s8989=ld8989["Sheet"]
        gh=s8989.cell(row=int(total_no_of_students.get())+1,column=1).value
        gh1=s8989.cell(row=int(total_no_of_students.get())+1,column=2).value
        gh2=s8989.cell(row=int(total_no_of_students.get())+1,column=3).value
        if gh is not None and gh1 is not None and gh2 is not None :
            global f12, th_om_dict, pr_om_dict, memory_list_th, c, d, th_memory1, pr_memory1, th_memory1_third, pr_memory1_third

            ld76 = op.load_workbook(f"E:\\{classname}.xlsx")
            s76 = ld76["Sheet"]
            memory_list_th = []

            def back_again():
                f12.destroy()

            if int(total_no_of_students.get()) <= 25:
                f12 = Frame(f4,bg="cyan2")
                f12.place(x=0, y=0, width=2500, height=1500)
                ld4567 = op.load_workbook(f'E:\\{classname}.xlsx')
                s4567 = ld4567["Sheet"]
                labu = Label(f12, text='S\nNo.', font='arial 10 bold',bg="cyan2")
                labu.grid(row=1, column=40)
                for i in range(1, int(total_no_of_students.get()) + 1):
                    labu1 = Label(f12, text=i, font='arial 10 bold',bg="cyan2")
                    labu1.grid(row=i + 1, column=40)
                for i in range(1, int(total_no_of_students.get()) + 1):
                    student_name_label = Label(f12, text=s4567.cell(row=i + 1, column=1).value, font='arial 10 bold',bg="cyan2")
                    student_name_label.grid(row=i + 1, column=0)
                for i in range(1, int(no_of_subject.get()) + 1):
                    sub_label = Label(f12, text=f'{subject_name_list[i - 1]}\n(TH)'.upper(), font='arial 9 bold',bg="cyan2")
                    sub_label.grid(row=1, column=2 * i)
                    sub_label = Label(f12, text=f'{subject_name_list[i - 1]}\n(PR)'.upper(), font='arial 9 bold',bg="cyan2")
                    sub_label.grid(row=1, column=2 * i + 1)

                th_memory_third = []
                pr_memory_third = []
                th_memory1_third = []
                pr_memory1_third = []

                for i in range(1, (int(total_no_of_students.get()) * int(no_of_subject.get())) + 1):
                    th_memory_third.append(StringVar())
                for i in range(1, (int(total_no_of_students.get()) * int(no_of_subject.get())) + 1):
                    pr_memory_third.append(StringVar())
                r = 2
                for i in range(0, (int(no_of_subject.get()) * int(total_no_of_students.get())), (int(no_of_subject.get()))):
                    if i == 0:
                        th_memory1_third.append(th_memory_third[0:int(no_of_subject.get())])
                        pr_memory1_third.append(pr_memory_third[0:int(no_of_subject.get())])
                    else:
                        th_memory1_third.append(th_memory_third[i:int(no_of_subject.get()) * r])
                        pr_memory1_third.append(pr_memory_third[i:int(no_of_subject.get()) * r])
                        r = r + 1

                for i in range(0, int(total_no_of_students.get())):
                    m = th_memory1_third[i]
                    n = pr_memory1_third[i]
                    for j in range(0, int(no_of_subject.get())):
                        theory_marks_entry = Entry(f12, textvariable=m[j], width=3, border=2, relief=SUNKEN,
                                                   font="arial 10 bold")
                        theory_marks_entry.grid(row=2 + i, column=2 * (j + 1))
                        practical_marks_entry = Entry(f12, textvariable=n[j], width=3, border=2, relief=SUNKEN,
                                                      font="arial 10 bold")
                        practical_marks_entry.grid(row=2 + i, column=(2 * (j + 1)) + 1)
                        m[j].set(0)
                        n[j].set(0)

                f13 = Frame(f12,bg="cyan")
                f13.place(x=0, y=640)
                back_again1 = Button(f13, text="Back to previous page", font="arial 12 bold", bg="bisque",fg="maroon",
                                     command=back_again)
                back_again1.grid(row=0, column=0, padx=10)
                print_marks_button1 = Button(f13, text="Save and Print Marksheet", font="arial 12 bold",bg="bisque",fg="maroon",
                                             command=print_marks1)
                print_marks_button1.grid(row=0, column=1)
                full_marks_button1 = Button(f13, text="View Full Marks of Subjects", font="arial 12 bold", bg="bisque",fg="maroon",
                                            command=full_marks)
                full_marks_button1.grid(row=0, column=2, padx=10)
                if s76.cell(row=2, column=4).value is not None:
                    for i in range(0, int(total_no_of_students.get())):
                        c0 = th_memory1_third[i]
                        d0 = pr_memory1_third[i]
                        for j in range(1, int(no_of_subject.get()) + 1):
                            c0[j - 1].set(s76.cell(row=i + 2, column=3 * j + 1).value)
                            d0[j - 1].set(s76.cell(row=i + 2, column=3 * j + 2).value)
                tmsg.showinfo("message",
                              "Simply write '0' in place of None for the added no. of students and added no. of subjects")





            elif int(total_no_of_students.get()) > 25 and int(total_no_of_students.get()) <= 45:
                f12 = Frame(f4,bg="cyan2")
                f12.place(x=0, y=0, width=2500, height=1500)
                ld4567 = op.load_workbook(f'E:\\{classname}.xlsx')
                s4567 = ld4567["Sheet"]
                labu = Label(f12, text='S\nNo.', font='arial 10 bold',bg="cyan2")
                labu.grid(row=1, column=40)
                for i in range(1, 26):
                    labu1 = Label(f12, text=i, font='arial 10 bold',bg="cyan2")
                    labu1.grid(row=i + 1, column=40)
                for i in range(1, 26):
                    student_name_label = Label(f12, text=s4567.cell(row=i + 1, column=1).value, font='arial 10 bold',bg="cyan2")
                    student_name_label.grid(row=i + 1, column=0)
                for i in range(1, int(no_of_subject.get()) + 1):
                    sub_label = Label(f12, text=f'{subject_name_list[i - 1]}\n(TH)'.upper(), font='arial 9 bold',bg="cyan2")
                    sub_label.grid(row=1, column=2 * i)
                    sub_label = Label(f12, text=f'{subject_name_list[i - 1]}\n(PR)'.upper(), font='arial 9 bold',bg="cyan2")
                    sub_label.grid(row=1, column=2 * i + 1)

                th_memory = []
                pr_memory = []
                th_memory1 = []
                pr_memory1 = []

                for i in range(1, (25 * int(no_of_subject.get())) + 1):
                    th_memory.append(StringVar())
                for i in range(1, (25 * int(no_of_subject.get())) + 1):
                    pr_memory.append(StringVar())
                a = 2
                for i in range(0, (int(no_of_subject.get()) * 25), (int(no_of_subject.get()))):
                    if i == 0:
                        th_memory1.append(th_memory[0:int(no_of_subject.get())])
                        pr_memory1.append(pr_memory[0:int(no_of_subject.get())])
                    else:
                        th_memory1.append(th_memory[i:int(no_of_subject.get()) * a])
                        pr_memory1.append(pr_memory[i:int(no_of_subject.get()) * a])
                        a = a + 1

                for i in range(0, 25):
                    c = th_memory1[i]
                    d = pr_memory1[i]
                    for j in range(0, int(no_of_subject.get())):
                        theory_marks_entry = Entry(f12, textvariable=c[j], width=3, border=2, relief=SUNKEN,
                                                   font="arial 10 bold")
                        theory_marks_entry.grid(row=2 + i, column=2 * (j + 1))
                        practical_marks_entry = Entry(f12, textvariable=d[j], width=3, border=2, relief=SUNKEN,
                                                      font="arial 10 bold")
                        practical_marks_entry.grid(row=2 + i, column=(2 * (j + 1)) + 1)
                        c[j].set(0)
                        d[j].set(0)
                # print(len(th_memory1[1]))

                f13 = Frame(f12,bg="cyan2")
                f13.place(x=0, y=640)
                remaining_button = Button(f13, text="Save and Add Remaining data", font="arial 12 bold", bg="bisque",fg="maroon",
                                          command=add_remaining_data)
                remaining_button.grid(row=0, column=1)

                # print(memory_list_th)
                # print(len(memory_list_th))

                back_again1 = Button(f13, text="Back to previous page", font="arial 12 bold", bg="bisque",fg="maroon",
                                     command=back_again)
                back_again1.grid(row=0, column=0, padx=10)
                full_marks_button = Button(f13, text="View Full Marks of Subjects", font="arial 12 bold", bg="bisque",fg="maroon",
                                           command=full_marks)
                full_marks_button.grid(row=0, column=2, padx=10)
                if s4567.cell(row=2, column=4).value is not None:
                    for i in range(0, 25):
                        c10 = th_memory1[i]
                        d10 = pr_memory1[i]
                        for j in range(1, int(no_of_subject.get()) + 1):
                            c10[j - 1].set(s4567.cell(row=i + 2, column=3 * j + 1).value)
                            d10[j - 1].set(s4567.cell(row=i + 2, column=3 * j + 2).value)
                tmsg.showinfo("message",
                              "Simply write '0' in place of None for the added no. of students and added no. of subjects")
        else:
            tmsg.showinfo("Error","Student details entry not completed")

    if c==0:
        yesno123=tmsg.askyesno("Warning","Check your entered data once\npress YES to continue")
        if yesno123==True:
            save_button.destroy()
            fv=Frame(f4,bg="cyan2")
            fv.place(x=220,y=470)
            student_details_button = Button(fv, text="Enter Students Details and Marks", font="arial 10 bold", bg="bisque",fg="maroon",
                                            command=student_details)
            student_details_button.grid(row=17 + int(no_of_subject.get()), column=1, pady=10)
            direct_marks_entry=Button(fv,text="Enter marks",font="arial 10 bold", bg="bisque",fg="maroon",
                                            command=direct_marks_entry)
            direct_marks_entry.grid(row=17 + int(no_of_subject.get()),column=2,pady=10,padx=5)
            ld89 = op.load_workbook(f"E:{classname}.xlsx")
            s89 = ld89["Sheet"]
            gh10 = s89.cell(row=int(total_no_of_students.get()) + 1, column=1).value
            gh11 = s89.cell(row=int(total_no_of_students.get()) + 1, column=2).value
            gh12 = s89.cell(row=int(total_no_of_students.get()) + 1, column=3).value
            if gh10 is not None and gh11 is not None and gh12 is not None:
                f89=Frame(f4,bg="cyan2")
                f89.place(x=450,y=600)
                label_info=Label(f89,text="You have completed students entry details\nYou can now enter marks\nclick on ENTER MARKS to enter the marks",font="arial 10 bold",bg="cyan2")
                label_info.pack()
    """print(subject_name_list)
    print(th_fm_list)
    print(pr_fm_list)
    print(credit_list)"""




def go_ahead():
    #total_no_of_students_entry.destroy()


    global f6, subject_dict, th_fm_dict, pr_fm_dict, credit_dict,wb1,save_button
    if os.path.exists(f"E:\\{classname}_full_marks.xlsx")==False:
        wb1 = op.Workbook(f"E:\\{classname}_full_marks.xlsx")
        wb1.save(f"E:\\{classname}_full_marks.xlsx")



    f6=Frame(f4,bg="cyan2")
    f6.place(x=0,y=180)
    if len(total_no_of_students.get())!=0 and  len(no_of_subject.get())!=0:
        if total_no_of_students.get().isdigit()==True and no_of_subject.get().isdigit()==True and int(total_no_of_students.get())>0 and int(no_of_subject.get())>0 :
            if int(no_of_subject.get())<=12:
                if int(total_no_of_students.get())<=45:
                    yesno = tmsg.askyesno("Warning","Check your entered data once\nYou can not submit these data again\npress YES to continue")
                    if yesno == True:
                        subject_order=['Subject 1','Subject 2','Subject 3','Subject 4','Subject 5','Subject 6','Subject 7','Subject 8','Subject 9','Subject 10','Subject 11','Subject 12']
                        subject_dict={i:StringVar() for i in range(1,int(no_of_subject.get())+1)}
                        th_fm_dict={i:StringVar() for i in range(1,int(no_of_subject.get())+1)}
                        pr_fm_dict = {i: StringVar() for i in range(1, int(no_of_subject.get()) + 1)}
                        credit_dict={i: StringVar() for i in range(1, int(no_of_subject.get()) + 1)}
                        #print(subject_dict)

                        for i in range(1,int(no_of_subject.get())+1):
                            space=Label(f6,text='\t\t\t\t',bg="cyan2")
                            space.grid(row=12,column=0)
                            space = Label(f6, text='\t\t\t\t',bg="cyan2")
                            space.grid(row=12, column=1)
                            info_label=Label(f6,text="Subject Name",font="arial 12 bold",bg="cyan2")
                            info_label.grid(row=12,column=2)

                            info_label = Label(f6, text="Full marks(TH)", font="arial 12 bold",bg="cyan2")
                            info_label.grid(row=12, column=3)
                            info_label = Label(f6, text="Full marks(PR)", font="arial 12 bold",bg="cyan2")
                            info_label.grid(row=12, column=4)
                            info_label = Label(f6, text="Credit Hour", font="arial 12 bold",bg="cyan2")
                            info_label.grid(row=12, column=5)

                            subject_label=Label(f6,text=subject_order[i-1],font="arial 12 bold",bg="cyan2")
                            subject_label.grid(row=13+i,column=1)
                            subject_entry=Entry(f6,textvariable=subject_dict[i],font="arial 12 bold",border=4,relief=SUNKEN,width=16)
                            subject_entry.grid(row=13+i,column=2)
                            th_fm_entry = Entry(f6, textvariable=th_fm_dict[i], font="arial 12 bold", border=4,relief=SUNKEN, width=16)
                            th_fm_entry.grid(row=13 + i, column=3)
                            pr_fm_entry = Entry(f6, textvariable=pr_fm_dict[i], font="arial 12 bold", border=4, relief=SUNKEN,width=16)
                            pr_fm_entry.grid(row=13 + i, column=4)
                            credit_entry = Entry(f6, textvariable=credit_dict[i], font="arial 12 bold", border=4, relief=SUNKEN,width=16)
                            credit_entry.grid(row=13 + i, column=5)
                            if os.path.exists(f"E:\\{classname}_full_marks.xlsx")==True:
                                ld3=op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
                                s3=ld3["Sheet"]
                                subject_dict[i].set(s3.cell(row=i+1,column=1).value)
                                th_fm_dict[i].set(s3.cell(row=i+1,column=2).value)
                                pr_fm_dict[i].set(s3.cell(row=i + 1, column=3).value)



                            credit_dict[i].set(4)

                        if os.path.exists(f"E:\\{classname}_full_marks.xlsx") == True:
                            ld24=op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
                            ws=ld24.active
                            ws.column_dimensions['A'].width="20"
                            ws.column_dimensions['B'].width = "20"
                            ws.column_dimensions['C'].width = "20"
                            ws.column_dimensions['D'].width = "20"
                            ws.column_dimensions['E'].width = "20"
                            ws.column_dimensions['F'].width = "20"
                            s24=ld24["Sheet"]
                            s24.cell(row=1, column=4).value = "Credit Hour"
                            s24.cell(row=1,column=5).value="No. of Student "
                            #s24.cell(row=1, column=6).value = "Section"
                            s24.cell(row=1, column=7).value = "No. of Subject "
                            s24.cell(row=2, column=5).value = total_no_of_students.get()
                            #s24.cell(row=2, column=6).value = section.get()
                            s24.cell(row=2, column=7).value = no_of_subject.get()
                            ld24.save(f"E:\\{classname}_full_marks.xlsx")
                            if s24.cell(row=2,column=4).value is not None:
                                for i in range(1,int(no_of_subject.get())+1):
                                    credit_dict[i].set(s24.cell(row=i+1,column=4).value)




                        #credit_dict[i].set(4)

                        save_button=Button(f6,text="Save and Continue",font="arial 12 bold",fg="maroon",bg="bisque",command=save)
                        save_button.grid(row=16+int(no_of_subject.get()),column=1,pady=10)

                        go_ahead_button.destroy()


                        #no_of_subject.set('')
                        #subject_name_list.append(subject_dict[i])
                        #th_fm_list.append(th_fm_dict[i])
                        #pr_fm_list.append(pr_fm_dict[i])"""
                else:
                    tmsg.showinfo("error","total no. of students should be less than or equal to 45")



            else:
                tmsg.showinfo("Error","No. of subject in a class can not be more than 12 ")

        else:
            tmsg.showinfo("Error", "no of students and subjects should be a positive integer greater than 0")
    else:
        tmsg.showinfo("Error","Input field can not be empty")


def nursery():
    global classname,classvalue
    classname="Nursery"
    classvalue = "NURSERY"
    class_work()
def lkg():
    global classname,classvalue
    classname="LKG"
    classvalue = "LKG"
    class_work()
def ukg():
    global classname,classvalue
    classname="UKG"
    classvalue = "UKG"
    class_work()


def class1():
    global classname,classvalue
    classname="class1"
    classvalue = 1
    class_work()


def class2():
    global classname,classvalue
    classname="class2"
    classvalue=2
    class_work()
def class3():
    global classname,classvalue
    classname="class3"
    classvalue = 3
    class_work()
def class4():
    global classname,classvalue
    classname="class4"
    classvalue = 4
    class_work()
def class5():
    global classname,classvalue
    classname="class5"
    classvalue = 5
    class_work()
def class6():
    global classname,classvalue
    classname="class6"
    classvalue = 6
    class_work()
def class7():
    global classname,classvalue
    classname="class7"
    classvalue = 7
    class_work()
def class8():
    global classname,classvalue
    classname="class8"
    classvalue = 8
    class_work()


def class_work():
    #print(os.path.exists(f"E:\\a.xlsx"))

    if os.path.exists(f"E:\\{classname}.xlsx")==False:
        wb=op.Workbook(f"E:\\{classname}.xlsx")
        wb.save(f"E:\\{classname}.xlsx")
    else:

        ld10=op.load_workbook(f"E:\\{classname}.xlsx")
        ws=ld10.active
        ws.column_dimensions['A'].width=25
        list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
                'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM',
                'AN', 'AO', 'AP']
        for i in range(2,len(list)+1):
            ws.column_dimensions[list[i-1]].width=18
        ld10.save(f"E:\\{classname}.xlsx")
        global nos_l,no_of_subject,no_of_subject_entry,f4,f5,class_l,no_l,total_no_of_students_entry,section,sec_l,section_entry,go_ahead_button,total_no_of_students
        f4=Frame(root)
        f4.configure(bg="cyan2")
        f4.place(x=0,y=0,width=2500,height=1500)

        f5=Frame(f4,bg="cyan2")
        f5.place(x=0,y=0)
        space1=Label(f5,text='\t\t\t\t',bg="cyan2")
        space1.grid(row=0,column=0)
        space1 = Label(f5, text='\t\t\t\t',bg="cyan2")
        space1.grid(row=0, column=1)

        back_button=Button(f5,text="Back to Previous Page",command=back,font="arial 12 bold",fg="maroon",bg="bisque")
        back_button.grid(row=0,column=2,pady=10)
        clear_button=Button(f5,text="Clear All Data",command=clear,font="arial 12 bold",fg="maroon",bg="bisque")
        clear_button.grid(row=0, column=3, pady=10)

        class_l=Label(f5,text=f"CLASS - {classvalue}",font="arial 13 bold",bg="cyan2")
        class_l.grid(row=6,column=2)


        no_l=Label(f5,text='Total No. Of Students',font="arial 13 bold",bg="cyan2")
        no_l.grid(row=7,column=2)
        total_no_of_students=StringVar()
        total_no_of_students_entry=Entry(f5,textvariable=total_no_of_students,width=20,font="arial 13 bold",border=4,relief=SUNKEN)
        total_no_of_students_entry.grid(row=7,column=3)



        nos_l = Label(f5, text='Total No. Of Subjects', font="arial 13 bold",bg="cyan2")
        nos_l.grid(row=9, column=2)
        no_of_subject = StringVar()
        no_of_subject_entry = Entry(f5, textvariable=no_of_subject, width=20, font="arial 13 bold", border=4, relief=SUNKEN)
        no_of_subject_entry.grid(row=9, column=3)
        #section.set("NA")

        if os.path.exists(f"E:\\{classname}_full_marks.xlsx")==True:
            ld456=op.load_workbook(f"E:\\{classname}_full_marks.xlsx")
            s456=ld456["Sheet"]
            total_no_of_students.set(s456.cell(row=2,column=5).value)
            #section.set(s456.cell(row=2, column=6).value)
            no_of_subject.set(s456.cell(row=2, column=7).value)


        go_ahead_button=Button(f5,text='Go Ahead',font="arial 13 bold",fg="maroon",bg="bisque",command=go_ahead)
        go_ahead_button.grid(row=10, column=3,pady=5)






















        ld=op.load_workbook(f"E:\\{classname}.xlsx")
        z=ld['Sheet']
        z.cell(row=1,column=1).value="NAME"
        z.cell(row=1, column=2).value = "ROLL NO."
        z.cell(row=1, column=3).value = "ADDRESS"
        ld.save(f"E:\\{classname}.xlsx")
        #os.startfile(f"E:\\a.xlsx")






root=Tk()
root.state("zoomed")
root.minsize(1300,800)
root.configure(bg="cyan2")
root.title("MKBS Marksheet Generator")
f1=Frame(root,bg="cyan2")
f1.pack()
l1=Label(f1,text="Welcome to MKBS Marksheet Generator",font="arial 30 bold",fg="maroon",bg="cyan2")
l1.pack()
f2=Frame(root,bg="cyan2")
f2.pack(anchor='n')
b1=Button(f2,text="Nursery",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=nursery)
b1.pack(anchor='n',side=LEFT,padx=30,pady=30)
b1=Button(f2,text="LKG",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=lkg)
b1.pack(anchor='n',side=LEFT,padx=30,pady=30)
b1=Button(f2,text="UKG",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=ukg)
b1.pack(anchor='n',side=LEFT,padx=30,pady=30)

f2=Frame(root,bg="cyan2")
f2.pack(anchor='n')
b1=Button(f2,text="Class 1",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=class1)
b1.pack(anchor='n',side=LEFT,padx=30,pady=30)
b1=Button(f2,text="Class 2",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=class2)
b1.pack(anchor='n',side=LEFT,padx=30,pady=30)
b1=Button(f2,text="Class 3",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=class3)
b1.pack(anchor='n',side=LEFT,padx=30,pady=30)
b1=Button(f2,text="Class 4",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=class4)
b1.pack(anchor='n',side=LEFT,padx=30,pady=30)
f3=Frame(root,bg="cyan2")
f3.pack(anchor='n')
b2=Button(f3,text="Class 5",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=class5)
b2.pack(anchor='n',side=LEFT,padx=30,pady=30)
b2=Button(f3,text="Class 6",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=class6)
b2.pack(anchor='n',side=LEFT,padx=30,pady=30)
b2=Button(f3,text="Class 7",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=class7)
b2.pack(anchor='n',side=LEFT,padx=30,pady=30)
b2=Button(f3,text="Class 8",font="arial 25 bold",padx=50,pady=30,fg="maroon",bg="bisque",command=class8)
b2.pack(anchor='n',side=LEFT,padx=30,pady=30)
root.mainloop()
