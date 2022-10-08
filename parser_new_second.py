from __future__ import annotations
from pdfminer.layout import LTTextBoxHorizontal
import pandas as pd
import csv
from pdfminer.high_level import extract_pages



class TheoryMarks:
    def __init__(self):
        self.marks = "0"

    def set_marks(self, marks: str):
        if(marks == "FF"):
            self.marks = "F"
        else:
            self.marks = marks

    def print(self) -> str:
        return self.marks


class LabMarks:
    def __init__(
        self,
        TW_marks: str = "---",
        PR_marks: str = "---",
        OR_marks: str = "---",
        Total_marks: str = "---",
    ):
        self.PR_marks = PR_marks  # format scored-marks/total-marks
        self.OR_marks = OR_marks
        self.TW_marks = TW_marks
        self.Total_marks = Total_marks

    def set_data(self, data: list[str]):
        self.TW_marks = data[3].strip()
        self.PR_marks = data[4].strip()
        self.OR_marks = data[5].strip()
        self.Total_marks = data[6].split("   ")[0].strip()

    def ret_data(self) -> list[str]:
        marks_list = []
        if self.TW_marks != "---":
            # print(self.TW_marks)
            # print(int(self.TW_marks.split("/")[1]))
            if(int(self.TW_marks.split("/")[1]) == 50):
          
                self.TW_marks = self.TW_marks.split("/")[0]
                marks_list.append(int(self.TW_marks)*2)
            elif (int(self.TW_marks.split("/")[1]) == 25):
               
                self.TW_marks = self.TW_marks.split("/")[0]
                marks_list.append(int(self.TW_marks)*4)
        else:
            marks_list.append("NA")
        if self.PR_marks != "---":
            #   print(self.PR_marks.split("/")[1])
              if(self.PR_marks =="10$/025"):
                self.PR_marks = 10
                marks_list.append(40) 
                #this above case exists solely because of an inconsistency in the ledger pdf, where the marks for the practicals are written as 10$ instead of 10/25
              elif(int(self.PR_marks.split("/")[1]) == 50):
             
                self.PR_marks = self.PR_marks.split("/")[0]
                marks_list.append(int(self.PR_marks)*2)
              elif (int(self.PR_marks.split("/")[1]) == 25):
               
                self.PR_marks = self.PR_marks.split("/")[0]
                marks_list.append(int(self.PR_marks)*4)
        else:
            marks_list.append("NA")
        if self.OR_marks != "---":
            if(int(self.OR_marks.split("/")[1]) == 50):
             
                self.OR_marks = self.OR_marks.split("/")[0]
                marks_list.append(int(self.OR_marks)*2)
            elif (int(self.OR_marks.split("/")[1]) == 25):
                self.OR_marks = self.OR_marks.split("/")[0]
                marks_list.append(int(self.OR_marks)*4)
        else:
            marks_list.append("NA")
        marks_list.append(self.Total_marks)
        return marks_list
    def print(self) -> int:
        return self.TW_marks

class CSVWriter:
    def __init__(self, csv_path: str):
        self.csv_path = csv_path
        self.csv_file = open(self.csv_path, "w", newline="")
        self.csv_writer = csv.writer(self.csv_file)
        self.csv_writer.writerow(
            [
                "Seat No",
                "Name",
                "EM3",
                "DSA",
                "SE",
                "MP",
                "PPL",
                "EM3 TW",
                "DSAL TW",
                "DSAL PR",
                "MPL TW",
                "MPL OR",
                "PBL2 TW",
                "COC TW",
                "SGPA",
               
            ]
        )
       

    def writeStudent(self, student:Student):
        self.csv_writer.writerow(student.tolist())
    

    def __del__(self):
        self.csv_file.close()


class Student:
    def __init__(self):
        self.full_name = ""
        self.seat_no = ""
        self.theory_marks_sub1 = TheoryMarks()
        self.theory_marks_sub2 = TheoryMarks()
        self.theory_marks_sub3 = TheoryMarks()
        self.theory_marks_sub4 = TheoryMarks()
        self.theory_marks_sub5 = TheoryMarks()
        self.lab_marks_sub1 = LabMarks()
        self.lab_marks_sub2 = LabMarks()
        self.lab_marks_sub3 = LabMarks()
        self.lab_marks_sub4 = LabMarks()
        self.lab_marks_sub5 = LabMarks()
        self.SGPA = 0

    def tolist(self) -> list[str, int]:
        lab1 = self.lab_marks_sub1.ret_data()
        lab2 = self.lab_marks_sub2.ret_data()
        lab3 = self.lab_marks_sub3.ret_data()
        lab4 = self.lab_marks_sub4.ret_data()
        lab5 = self.lab_marks_sub5.ret_data()
        return [
            self.seat_no,
            self.full_name,
            self.theory_marks_sub1.marks,
            self.theory_marks_sub2.marks,
            self.theory_marks_sub3.marks,
            self.theory_marks_sub4.marks,
            self.theory_marks_sub5.marks,
            lab1[0],
            lab2[0],
            lab2[1],
            lab3[0],
            lab3[2],
            lab4[0],
            lab5[0],
            self.SGPA,
        ]

    def clear(self):
        self.full_name = ""
        self.seat_no = ""
        self.theory_marks_sub1 = TheoryMarks()
        self.theory_marks_sub2 = TheoryMarks()
        self.theory_marks_sub3 = TheoryMarks()
        self.theory_marks_sub4 = TheoryMarks()
        self.theory_marks_sub5 = TheoryMarks()
        self.lab_marks_sub1 = LabMarks()
        self.lab_marks_sub2 = LabMarks()
        self.lab_marks_sub3 = LabMarks()
        self.lab_marks_sub4 = LabMarks()
        self.lab_marks_sub5 = LabMarks()
    def __str__(self) -> str:
        return f"{self.full_name} {self.seat_no} {self.theory_marks_sub1.print()} {self.theory_marks_sub2.print()} {self.theory_marks_sub3.print()} {self.theory_marks_sub4.print()} {self.theory_marks_sub5.print()} {self.lab_marks_sub1.ret_data()} {self.lab_marks_sub2.ret_data()} {self.lab_marks_sub3.ret_data()} {self.lab_marks_sub4.ret_data()} {self.lab_marks_sub5.ret_data()} {self.SGPA} \n"


class SmartParse:
    object_counter: int = 0
    counter: int = 0
    student: Student = Student()
    csv_writer: CSVWriter = CSVWriter("marks.csv")
    def parse_boxes(self , name_box:LTTextBoxHorizontal,marks_box:LTTextBoxHorizontal):
        try:
            for name in name_box:            
                self.student.full_name  = name.get_text().split(":")[2].split("    ")[0]
                self.student.seat_no = name.get_text().split(":")[1].split("NAME")[0]     
                break
        except:
            print("reached end")
            return 
        order_dict = {
            0: self.student.theory_marks_sub1,
            1: self.student.lab_marks_sub1,
            2: self.student.theory_marks_sub2,
            3: self.student.theory_marks_sub3,
            4: self.student.theory_marks_sub4,
            5: self.student.theory_marks_sub5,
            6: self.student.lab_marks_sub2,
            7: self.student.lab_marks_sub3,
            8: self.student.lab_marks_sub4,
            9: self.student.lab_marks_sub5,
        } 
        for text_line in marks_box:
           
            if (
            "CONFIDENTIAL" in text_line.get_text()
            or "COURSE" in text_line.get_text()
            or "SEM" in text_line.get_text()
            or "210260A" in text_line.get_text()
            or "210260B" in text_line.get_text()
        ):  # avoiding unwanted lines.
                continue
            parse_line = text_line.get_text()
            
            if self.counter in [0,2,3,4,5]:
                con_str = parse_line.split("*")[1]
                total_marks = list(map("".join, zip(*[iter(con_str)] * 9)))[6].split("   ")[
                    0
                ]  # splitting the line after * in 9 parts.
                

                order_dict[self.counter].set_marks(total_marks.strip())
                SmartParse.counter += 1
            elif "SGPA" in parse_line:
                self.student.SGPA = parse_line.split(":")[1].split(",")[0]
                #print(self.student)
 
                SmartParse.csv_writer.writeStudent(
                    self.student
                    ) 
                SmartParse.object_counter += 1 
                print(f"{self.object_counter} objects written")
                self.student.clear()
                SmartParse.counter = 0
            else:
                con_str = parse_line.split("*")[1]
                data = list(
                    map("".join, zip(*[iter(con_str)] * 9))
                )  # splitting the line after * in 9 parts.
                order_dict[self.counter].set_data(data)
                SmartParse.counter += 1
     

   
def getLTBoxCount(obj) -> int:
    count = 0 
    for element in obj:
        if(isinstance(element, LTTextBoxHorizontal)):
            count += 1
    return count 
for page_layout in extract_pages("se.pdf"):
  
    if getLTBoxCount(page_layout) == 5: 
        SmartParse().parse_boxes(page_layout._objs[1],page_layout._objs[2])
        SmartParse().parse_boxes(page_layout._objs[3],page_layout._objs[4])
    elif getLTBoxCount(page_layout) == 3:
        SmartParse().parse_boxes(page_layout._objs[1],page_layout._objs[2])
    else:
        print("End")
        break 
        
try:
    xl = pd.ExcelWriter(
        "se_marks.xlsx",
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_numbers": True}},
    )
    df = pd.read_csv("marks.csv")
    df.to_excel(xl ,index = False,na_rep = "NOF")
    xl.save()
except Exception as e:
    print(e)