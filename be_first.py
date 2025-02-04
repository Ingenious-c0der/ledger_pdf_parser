from __future__ import annotations
from pdfminer.layout import LTTextBoxHorizontal
import pandas as pd
import csv
from pdfminer.high_level import extract_pages

INPUT = "inputs/BECOMP2019.pdf"
OUTPUT = "generated/be_2025.csv"
class TheoryMarks:
    def __init__(self):
        self.marks = "NA"

    def set_marks(self, marks: str):
        if(marks == "FF"):
            self.marks = "F"
        elif "$" in marks:
            self.marks = marks.split("$")[0]
        else:
            self.marks = marks

    def print(self) -> int:
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
        if self.TW_marks != "---" and "AB" not in self.TW_marks and "$" not in self.TW_marks :
            # print(self.TW_marks)
            # print(int(self.TW_marks.split("/")[1]))
            if(int(self.TW_marks.split("/")[1]) == 50):
          
                self.TW_marks = self.TW_marks.split("/")[0]
                marks_list.append(int(self.TW_marks)*2)
            elif (int(self.TW_marks.split("/")[1]) == 25):
               
                self.TW_marks = self.TW_marks.split("/")[0]
                marks_list.append(int(self.TW_marks)*4)
            else: 
                # case for /100 
                # print(int(self.TW_marks.split("/")[0]))
                marks_list.append(int(self.TW_marks.split("/")[0]))
        elif "AB" in self.TW_marks :
            marks_list.append("AB")
        elif "$" in self.TW_marks :
            marks_list.append(int(self.TW_marks.split("$")[0]))
        else:
            marks_list.append("NA")
        
        if self.PR_marks != "---" and "AB" not in self.PR_marks and "$" not in self.PR_marks  :
              if(int(self.PR_marks.split("/")[1]) == 50):
             
                self.PR_marks = self.PR_marks.split("/")[0]
                marks_list.append(int(self.PR_marks)*2)
              elif (int(self.PR_marks.split("/")[1]) == 25):
               
                self.PR_marks = self.PR_marks.split("/")[0]
                marks_list.append(int(self.PR_marks)*4)
        elif  "AB" in self.PR_marks:
            marks_list.append("AB")
        elif "$" in self.PR_marks :
            marks_list.append(int(self.PR_marks.split("$")[0]))
        else:
            marks_list.append("NA")
        
        if self.OR_marks != "---" and  "AB" not in self.OR_marks and "$" not in self.OR_marks :
            if(int(self.OR_marks.split("/")[1]) == 50):
             
                self.OR_marks = self.OR_marks.split("/")[0]
                marks_list.append(int(self.OR_marks)*2)
            elif (int(self.OR_marks.split("/")[1]) == 25):
                self.OR_marks = self.OR_marks.split("/")[0]
                marks_list.append(int(self.OR_marks)*4)
        elif "AB"  in self.OR_marks:
            marks_list.append("AB")
        elif "$" in self.OR_marks :
            
            marks_list.append(int(self.OR_marks.split("$")[0]))
        else:
            marks_list.append("NA")
        marks_list.append(self.Total_marks)

        return marks_list

class CSVWriter:
    def __init__(self, csv_path: str):
        self.csv_path = csv_path
        self.csv_file = open(self.csv_path, "w",newline="")
        self.csv_writer = csv.writer(self.csv_file)
        self.csv_writer.writerow(
            [
                "Seat No",
                "Name",
                "DESIGN & ANALYSIS OF ALGO",
                "MACHINE LEARNING",
                "BLOCKCHAIN TECHNOLOGY ",
                "PERVASIVE COMPUTING",
                "MULTIMEDIA TECHNIQUES",
                "CYBER SEC & DIGITAL FORENSICS",
                "OBJ. ORIENTED MODL. & DESIGN",
                "INFORMATION RETRIEVAL",
                "MOBILE COMPUTING",
                "SOFTWARE TESTING & QUALITY ASSURANCE",
                "LABORATORY",
                "PRACTICE",
                "III",
                "LABORATORY",
                "PRACTICE",
                "IV",
                "PROJECT",
                "STAGE",
                "I",
                "SGPA",
                "Credits"
            ]
        )
        self.csv_writer.writerow(
            [
                " ",
                " ",
                " ",
                " ",
                " ",
                " ",
                " ",
                " ",
                " ",
                " ",
                " ",
                " ",
                "TW",
                "PR",
                "OR",
                "TW",
                "PR",
                "OR",
                "TW",
                "PR",
                "OR",
            ]
        )  # Header of the csv file

    def writeStudent(self, student):
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
        self.theory_marks_sub6 = TheoryMarks()
        self.theory_marks_sub7 = TheoryMarks()
        self.theory_marks_sub8 = TheoryMarks()
        self.theory_marks_sub9 = TheoryMarks()
        self.theory_marks_sub10 = TheoryMarks()
        self.lab_marks_sub1 = LabMarks()
        self.lab_marks_sub2 = LabMarks()
        self.lab_marks_sub3 = LabMarks()
        self.SGPA = 0
        self.Credits = 0 


    def tolist(self) -> list[str, int]:
        lab1 = self.lab_marks_sub1.ret_data()
        lab2 = self.lab_marks_sub2.ret_data()
        lab3 = self.lab_marks_sub3.ret_data()
        return [
            self.seat_no,
            self.full_name,
            self.theory_marks_sub1.print(),
            self.theory_marks_sub2.print(),
            self.theory_marks_sub3.print(),
            self.theory_marks_sub4.print(),
            self.theory_marks_sub5.print(),
            self.theory_marks_sub6.print(),
            self.theory_marks_sub7.print(),
            self.theory_marks_sub8.print(),
            self.theory_marks_sub9.print(),
            self.theory_marks_sub10.print(),
            lab1[0],
            lab1[1],
            lab1[2],
            lab2[0],
            lab2[1],
            lab2[2],
            lab3[0],
            lab3[1],
            lab3[2],
            self.SGPA,
            self.Credits
        ]

    def clear(self):
        self.full_name = ""
        self.seat_no = ""
        self.theory_marks_sub1 = TheoryMarks()
        self.theory_marks_sub2 = TheoryMarks()
        self.theory_marks_sub3 = TheoryMarks()
        self.theory_marks_sub4 = TheoryMarks()
        self.theory_marks_sub5 = TheoryMarks()
        self.theory_marks_sub6 = TheoryMarks()
        self.theory_marks_sub7 = TheoryMarks()
        self.theory_marks_sub8 = TheoryMarks()
        self.theory_marks_sub9 = TheoryMarks()
        self.theory_marks_sub10 = TheoryMarks()
        self.lab_marks_sub1 = LabMarks()
        self.lab_marks_sub2 = LabMarks()
        self.lab_marks_sub3 = LabMarks()
        self.SGPA = 0


class SmartParse:
    object_counter: int = 0
    counter: int = -1
    student: Student = Student()
    csv_writer: CSVWriter = CSVWriter(OUTPUT)

    def ordered_parse(self, parse_line: str):
        if (
            "CONFIDENTIAL" in text_line.get_text()
            or "COURSE" in text_line.get_text()
            or "SEM" in text_line.get_text()
            or "410301" in text_line.get_text()
            or "410501" in text_line.get_text()
            or "410401" in text_line.get_text()
            or "410402" in text_line.get_text()
            or "410249" in text_line.get_text()
            or "411701" in text_line.get_text()
            #sem 2 subjects, retake students 
            or "410250" in text_line.get_text()
            or "410251" in text_line.get_text()
            or "410252" in text_line.get_text()
            or "410253" in text_line.get_text()
            or "410254" in text_line.get_text()
            or "410255" in text_line.get_text()
            or "410256" in text_line.get_text()
            or "410257" in text_line.get_text()
            or "FE SGPA" in text_line.get_text()
            or "SE SGPA" in text_line.get_text()
            or "TOTAL GRADE" in text_line.get_text()
        ):  # avoiding unwanted lines.
            return
        order_dict = {
            -1: "Seat and Name",
            0: self.student.theory_marks_sub1,
            1: self.student.theory_marks_sub2,
            2: self.student.theory_marks_sub3,
            3: self.student.theory_marks_sub4, # 410244A
            4: self.student.theory_marks_sub5, # 410244B
            5: self.student.theory_marks_sub6, # 410244C
            6: self.student.theory_marks_sub7, # 410244D
            7: self.student.theory_marks_sub8, # 410245A
            8: self.student.theory_marks_sub9, # 410245C
            9: self.student.theory_marks_sub10, # 410245D
            10: self.student.lab_marks_sub1,
            11: self.student.lab_marks_sub2,
            12: self.student.lab_marks_sub3,
        }  # the lines noted and the corresponding objects parameters.

        if self.counter == -1:
            self.student.full_name = parse_line.split(":")[2].split("    ")[0]  # name
            seat_no = parse_line.split(":")[1].split(" ")[1]  # seat no
            if seat_no.endswith("NAME"):
                seat_no = seat_no[:-4]
            self.student.seat_no = seat_no
            SmartParse.counter += 1  # increment counter
            return

        if self.counter < 10 and self.counter > -1:
            if "*" not in parse_line:
                print(parse_line)
                index = parse_line.find("/")
                index-=3
                con_str = parse_line[index:]
                total_marks = list(map("".join, zip(*[iter(con_str)] * 9)))[2].split("   ")[
                    0
                ] # splitting the line after * in 9 parts.
            else: 
                con_str = parse_line.split("*")[1]            
                total_marks = list(map("".join, zip(*[iter(con_str)] * 9)))[2].split("   ")[
                        0
                    ]  # splitting the line after * in 9 parts.
            if (self.counter == 3):
                if("410244A" in parse_line): 
                    order_dict[self.counter].set_marks(total_marks)
                   
                elif("410244B" in parse_line):
                    order_dict[self.counter+1].set_marks(total_marks)
                   
                elif("410244C" in parse_line):
                    order_dict[self.counter+2].set_marks(total_marks)
                    
                elif("410244D" in parse_line):
                    order_dict[self.counter+3].set_marks(total_marks)
                SmartParse.counter = 7
            elif(self.counter == 7):
                if("410245A" in parse_line):
                    order_dict[self.counter].set_marks(total_marks)
                    
                elif("410245C" in parse_line):
                    order_dict[self.counter+1].set_marks(total_marks)
                    
                elif("410245D" in parse_line):
                    order_dict[self.counter+2].set_marks(total_marks)
                SmartParse.counter = 10
            else:
                order_dict[self.counter].set_marks(total_marks)
                SmartParse.counter += 1
           
        elif "SGPA" in parse_line:
                self.student.SGPA = parse_line.split(":")[1].split(",")[0]
                self.student.Credits = parse_line.split(":")[-1].strip()
                print(parse_line.split(":"))
                print(self.student.Credits)
                SmartParse.csv_writer.writeStudent(
                    self.student
                )  # writing the student to the csv file.
                SmartParse.counter = -1  # resetting the counter.
                SmartParse.object_counter += 1  # increasing the object counter.
                print(self.student.full_name)
                print(f"{self.object_counter} objects written")
                SmartParse.student.clear()  # clearing the student object.
        else:
            if "*" not in parse_line:
                    index = parse_line.find("---")
                    con_str = "   "+ parse_line[index:]
            else: 
                con_str = parse_line.split("*")[1]
            data = list(
                map("".join, zip(*[iter(con_str)] * 9))
            )  # splitting the line after * in 9 parts.
            order_dict[self.counter].set_data(data)
            SmartParse.counter += 1
          


for page_layout in extract_pages(INPUT):
    for element in page_layout:
        if isinstance(element, (LTTextBoxHorizontal)):
            if "PUNE" in element.get_text():
                continue  # avoiding unwanted blocks.
            else:
                for text_line in element:
                    SmartParse().ordered_parse(text_line.get_text())

# try:
#     xl = pd.ExcelWriter(
#         "be_marks.xlsx",
#         engine="xlsxwriter",
#         engine_kwargs={"options": {"strings_to_numbers": True}},
#     )
#     df = pd.read_csv(OUTPUT)
#     df.to_excel(xl ,index = False,na_rep = "NOF")
#     xl.save()
# except Exception as e:
#     print(e)
