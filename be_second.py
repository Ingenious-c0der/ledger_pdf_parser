from __future__ import annotations
import itertools
from pdfminer.layout import LTPage, LTTextBoxHorizontal
import pandas as pd
import csv
from pdfminer.high_level import extract_pages



INPUT  = "inputs/BE_2024.pdf"
OUTPUT = "generated/BE_2024.csv"

class TheoryMarks:
    def __init__(self):
        self.marks = "NA"

    def set_marks(self, marks: str):
        if(marks == "FF"):
            self.marks = "F"
        elif "$" in marks:
            self.marks = marks.split("$")[0]
        else:
            if("/" in marks):
                self.marks = marks.split("/")[0]
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
        if self.TW_marks != "---" and "AB" not in self.TW_marks and "$" not in self.TW_marks :
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
    def print(self) -> int:
        return self.TW_marks
    def __str__(self) -> str:
        return f"TW {self.TW_marks} PR {self.PR_marks} OR {self.OR_marks} TOT {self.Total_marks}"

class CSVWriter:
    def __init__(self, csv_path: str):
        self.csv_path = csv_path
        self.csv_file = open(self.csv_path, "w", newline="")
        self.csv_writer = csv.writer(self.csv_file)
        self.csv_writer.writerow(
            [
                "Seat No",
                "Name",
                "HIGH PERFORMANCE COMPUTING",
                "DEEP LEARNING",
                "NATURAL LANGUAGE PROCESSING",
                "IMAGE PROCESSING",
                "PATTERN RECOGNITION",
                "BUSINESS INTELLIGENCE",
                "LABORATORY ",
                "PRACTICE V",
                "LABORATORY PRACTICE VI",
                "PROJECT",
                "STAGE II",
                "FE SGPA",
                "SE SGPA",
                "TE SGPA",
                "BE SGPA",
                "CGPA",
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
                "TW",
                "PR",
                "TW",
                "TW",
                "OR",
            ]
        )  # Header of the csv file


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
        self.theory_marks_sub6 = TheoryMarks()
        self.lab_marks_sub1 = LabMarks()
        self.lab_marks_sub2 = LabMarks()
        self.lab_marks_sub3 = LabMarks()
        self.SGPA = 0
    #function added explicitly to calculate FE sgpa for cases where it is not available
    def computeScores(self):
        if(self.FE_SGPA == 0):
            print("Zero FE SGPA caught")
            print(self.CGPA, self.SE_SGPA, self.TE_SGPA,self.BE_SGPA)
            self.FE_SGPA = round((float(self.CGPA) * 4) - (float(self.SE_SGPA) + float(self.TE_SGPA) + float(self.BE_SGPA)),2)
            print("reconstructed fe sgpa ", self.FE_SGPA)
    def tolist(self) -> list[str, int]:
        lab1 = self.lab_marks_sub1.ret_data()
        lab2 = self.lab_marks_sub2.ret_data()
        lab3 = self.lab_marks_sub3.ret_data()
        return [
            self.seat_no,
            self.full_name,
            self.theory_marks_sub1.marks,#HPC
            self.theory_marks_sub2.marks,#DL
            self.theory_marks_sub3.marks, #NLP
            self.theory_marks_sub4.marks, #IP
            self.theory_marks_sub5.marks, #PR
            self.theory_marks_sub6.marks, #BI
            lab1[0],
            lab1[1],
            lab2[0],
            lab3[0],
            lab3[2],
            self.FE_SGPA,
            self.SE_SGPA,
            self.TE_SGPA,
            self.BE_SGPA,
            self.CGPA,
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
        self.lab_marks_sub1 = LabMarks()
        self.lab_marks_sub2 = LabMarks()
        self.lab_marks_sub3 = LabMarks()
        self.CGPA = 0
        self.FE_SGPA = 0
        self.SE_SGPA = 0
        self.TE_SGPA = 0
        self.BE_SGPA = 0


    def __str__(self) -> str:
        return f"{self.full_name} {self.seat_no} {self.theory_marks_sub1.print()} {self.theory_marks_sub2.print()} {self.theory_marks_sub3.print()} {self.theory_marks_sub4.print()} {self.theory_marks_sub5.print()} {self.theory_marks_sub6.print()} {self.lab_marks_sub1.ret_data()} {self.lab_marks_sub2.ret_data()} {self.lab_marks_sub3.ret_data()}  {self.SGPA} \n"


class SmartParse:
    object_counter: int = 0
    counter: int = 0
    sgpa_parsed: bool = False
    student_parsed: bool = False
    student: Student = Student()
    csv_writer: CSVWriter = CSVWriter(OUTPUT)


    def parse_boxes(self , name_box:LTTextBoxHorizontal,marks_box:LTTextBoxHorizontal,next_page_layout:LTPage):
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
            1: self.student.theory_marks_sub2,
            2: self.student.theory_marks_sub3,
            3: self.student.theory_marks_sub4,
            4: self.student.theory_marks_sub5,
            5: self.student.theory_marks_sub6,
            6: self.student.lab_marks_sub1,
            7: self.student.lab_marks_sub2,
            8: self.student.lab_marks_sub3,
        }
        for text_line in marks_box:

            if (
            "CONFIDENTIAL" in text_line.get_text()
            or "COURSE" in text_line.get_text()
            #or "CGPA" in text_line.get_text()
            or "SEM" in text_line.get_text()
            or "410257" in text_line.get_text()
            or "410503" in text_line.get_text() # honors ML,DS
            or "410403" in text_line.get_text() # honors INFO SYS MGMT
             or "410303" in text_line.get_text() # honors INFO SYS MGMT
            or "411703" in text_line.get_text()
        ):  # avoiding unwanted lines.
                continue
            parse_line = text_line.get_text()
            SmartParse.student_parsed = False
            if self.counter in [0,1,2,3]: #for theory marks

                if "*" not in parse_line:
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

                if(self.counter == 2):
                    sub_code = parse_line.split("*")[0].split(" ")[3]

                    #check if sub code contains A
                    if "A" in sub_code:
                        order_dict[2].set_marks(total_marks.strip())
                    else:
                        order_dict[3].set_marks(total_marks.strip())
                elif(self.counter == 3):

                    sub_code = parse_line.split("*")[0].split(" ")[3]

                    #check if sub code contains A
                    if "A" in sub_code:
                        order_dict[4].set_marks(total_marks.strip())
                    else:
                        order_dict[5].set_marks(total_marks.strip())
                else:

                    order_dict[self.counter].set_marks(total_marks.strip())
                self.counter += 1
                if(self.counter == 4):
                    self.counter = 6

            elif "CGPA" in parse_line:

                self.student.CGPA = parse_line.split(":")[-1].split()[0].strip()
                self.student.computeScores()
                # max till 2 decimal places
                #self.student.CGPA = round(sum(SGPA_array)/4,2)
                SmartParse.csv_writer.writeStudent(
                    self.student
                    )
                SmartParse.object_counter += 1
                print(f"{self.object_counter} objects written - {self.student.full_name}")
                self.student.clear()
                self.counter = 0
                SmartParse.student_parsed = True

            elif "SGPA" in parse_line:

                if (self.sgpa_parsed):
                    SGPA_parse_array_line = parse_line.split(":")

                    #to account for mising grading info in the ledger pdf for FE
                    missing_info = False
                    if("FE" not in parse_line):
                        missing_info = True
                        print("Missing info for fe")
                        self.student.FE_SGPA = 0.0
                        self.student.SE_SGPA = SGPA_parse_array_line[1].split("   ")[0].strip()
                        self.student.TE_SGPA = SGPA_parse_array_line[2].split("   ")[0].strip()
                    else:
                        self.student.FE_SGPA = SGPA_parse_array_line[1].split("   ")[0].strip()
                        self.student.SE_SGPA = SGPA_parse_array_line[2].split("   ")[0].strip()
                        self.student.TE_SGPA = SGPA_parse_array_line[3].split("   ")[0].strip()
                    SGPA_array = list(map(float,[self.student.FE_SGPA,self.student.SE_SGPA,self.student.TE_SGPA,self.student.BE_SGPA]))


                    self.sgpa_parsed = False
                else:
                    self.student.BE_SGPA = parse_line.split(":")[1].split(",")[0].strip()
                    if(self.student.BE_SGPA == "--"):
                        print("Student has failed in BE")
                        self.student.FE_SGPA = "NA"
                        self.student.SE_SGPA = "NA"
                        self.student.TE_SGPA = "NA"
                        self.student.CGPA = "--"
                        SmartParse.csv_writer.writeStudent(
                        self.student
                        )
                        SmartParse.object_counter += 1
                        print(f"{self.object_counter} objects written - {self.student.full_name}")
                        self.student.clear()
                        self.counter = 0
                        self.sgpa_parsed = False
                        SmartParse.student_parsed = True
                    else:
                        self.student.BE_SGPA = float(self.student.BE_SGPA)
                        self.sgpa_parsed = True

            else:
                # for labs

                if "*" not in parse_line:
                        index = parse_line.find("---")
                        con_str = "   "+ parse_line[index:]
                else:
                    con_str = parse_line.split("*")[1]
                data = list(
                    map("".join, zip(*[iter(con_str)] * 9))
                )  # splitting the line after * in 9 parts.

                order_dict[self.counter].set_data(data)

                self.counter += 1
        if(not SmartParse.student_parsed):
            #special case where only the cgpa is overflows on the next page
            #use the next page layout to complete the student object
            #considering that the next page layout is only minor overflow, hence page._objs[0]
            print("Overflow detected")
            overflow_box :LTTextBoxHorizontal = next_page_layout._objs[0]
            for line in overflow_box:
                parse_line  = line.get_text()
                print("Overflow2", parse_line)
                if("CGPA" in parse_line):
                    print("Reading overflow CGPA")
                    self.student.CGPA = parse_line.split(":")[-1].split()[0].strip()
                    print(self.student.CGPA)
                    self.student.computeScores()
                    # max till 2 decimal places
                    #self.student.CGPA = round(sum(SGPA_array)/4,2)
                    SmartParse.csv_writer.writeStudent(
                        self.student
                        )
                    SmartParse.object_counter += 1
                    print(f"{self.object_counter} objects written - {self.student.full_name}")
                    self.student.clear()
                    self.counter = 0
                    SmartParse.student_parsed = True
                    return





def getLTBoxCount(obj) -> int:
    count = 0
    for element in obj:
        if(isinstance(element, LTTextBoxHorizontal)):
            count += 1
    return count
try:

    page_iter, lookahead_iter = itertools.tee(extract_pages(INPUT),2)
    next(lookahead_iter)
    for page_layout in page_iter:
        next_page_layout = next(lookahead_iter,None)
        if getLTBoxCount(page_layout) == 5:
            SmartParse().parse_boxes(page_layout._objs[1],page_layout._objs[2], next_page_layout)
            SmartParse().parse_boxes(page_layout._objs[3],page_layout._objs[4], next_page_layout)
        elif getLTBoxCount(page_layout) == 3:
            SmartParse().parse_boxes(page_layout._objs[1],page_layout._objs[2], next_page_layout)
        else:
            print("Blank Page")
except Exception as e:
    print("Error : " + e)


# try:
#     xl = pd.ExcelWriter(
#         "generated/be_new_marks.xlsx",
#         engine="xlsxwriter",
#         engine_kwargs={"options": {"strings_to_numbers": True}},
#     )
#     df = pd.read_csv(OUTPUT)
#     df.to_excel(xl ,index = False,na_rep = "NOF")
#     xl.save()
# except Exception as e:
#     print(e)
