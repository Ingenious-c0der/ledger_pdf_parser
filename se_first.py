from __future__ import annotations
from pdfminer.layout import LTTextBoxHorizontal
import pandas as pd
import csv
from pdfminer.high_level import extract_pages


class TheoryMarks:
    def __init__(self):
        self.marks = 0

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
        if self.TW_marks != "---" and "AB" not in self.TW_marks and "$" not in self.TW_marks and "#" not in self.TW_marks :
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
        elif "#" in self.TW_marks:
            marks_list.append(int(self.TW_marks.split("#")[0]))
        else:
            marks_list.append("NA")
        
        if self.PR_marks != "---" and "AB" not in self.PR_marks and "$" not in self.PR_marks and "#" not in self.PR_marks  :
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
        elif "#" in self.PR_marks:
            marks_list.append(int(self.PR_marks.split("#")[0]))
        else:
            marks_list.append("NA")
        
        if self.OR_marks != "---" and  "AB" not in self.OR_marks and "$" not in self.OR_marks and "#" not in self.OR_marks:
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
        elif "#" in self.OR_marks:
            marks_list.append(int(self.OR_marks.split("#")[0]))
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
                "DISCRETE MATHEMATICS",
                "FUND. OF DATA STRUCTURES",
                "OBJECT ORIENTED PROGRAMMING",
                "COMPUTER GRAPHICS",
                "DIGITAL ELEC. & LOGIC DESIGN",
                "DATA",
                "STUCTURES",
                "LABORATORY",
                "OOP &",
                " COMP. ",
                "GRAPHICS LAB",
                "DIGITAL",
                " ELEC.",
                " LABORATORY",
                "BUSINESS",
                " COMMUNICATION",
                " SKILLS",
                "HUMANITY ",
                "& SOCIAL",
                " SCIENCE",
                  "SGPA",
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
                "TW",
                "PR",
                "OR",
                "TW",
                "PR",
                "OR",
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
            self.theory_marks_sub1.print(),
            self.theory_marks_sub2.print(),
            self.theory_marks_sub3.print(),
            self.theory_marks_sub4.print(),
            self.theory_marks_sub5.print(),
            lab1[0],
            lab1[1],
            lab1[2],
            lab2[0],
            lab2[1],
            lab2[2],
            lab3[0],
            lab3[1],
            lab3[2],
            lab4[0],
            lab4[1],
            lab4[2],
            lab5[0],
            lab5[1],
            lab5[2],
            self.SGPA
            
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
        self.SGPA = 0


class SmartParse:
    object_counter: int = 0
    counter: int = -1
    student: Student = Student()
    csv_writer: CSVWriter = CSVWriter("generated/SE_2023_MARKS.csv")

    def ordered_parse(self, parse_line: str):
        if (
            "CONFIDENTIAL" in text_line.get_text()
            or "COURSE" in text_line.get_text()
            or "SEM" in text_line.get_text()
            or "210251" in text_line.get_text()
        ):  # avoiding unwanted lines.
            return
        order_dict = {
            -1: "Seat and Name",
            0: self.student.theory_marks_sub1,
            1: self.student.theory_marks_sub2,
            2: self.student.theory_marks_sub3,
            3: self.student.theory_marks_sub4,
            4: self.student.theory_marks_sub5,
            5: self.student.lab_marks_sub1,
            6: self.student.lab_marks_sub2,
            7: self.student.lab_marks_sub3,
            8: self.student.lab_marks_sub4,
            9: self.student.lab_marks_sub5,
        }  # the lines noted and the corresponding objects parameters.

        if self.counter == -1:

            self.student.full_name = parse_line.split(":")[2].split("    ")[0]  # name
            self.student.seat_no = parse_line.split(":")[1].split(" ")[1]  # seat no
            SmartParse.counter += 1  # increment counter
            return

        if self.counter < 5 and self.counter > -1:
            if("*" not in parse_line):
                print("2 SEMS in one SEM detected,skipping line" )
                exit(0)
                return 
            con_str = parse_line.split("*")[1]
            total_marks = list(map("".join, zip(*[iter(con_str)] * 9)))[6].split("   ")[
                0
            ]  # splitting the line after * in 9 parts.

            order_dict[self.counter].set_marks(total_marks)
            SmartParse.counter += 1

        elif "SGPA1" in parse_line:
            self.student.SGPA = parse_line.split(":")[1].split(",")[0]
            SmartParse.csv_writer.writeStudent(
                self.student
            )  # writing the student to the csv file.
            SmartParse.counter = -1  # resetting the counter.
            SmartParse.object_counter += 1  # increasing the object counter.
            print(f"{self.object_counter,self.student.full_name,self.student.SGPA} objects written")
            SmartParse.student.clear()  # clearing the student object.

        else:
            print(self.counter)
            if("*" not in parse_line):
                print("2 SEMS in one SEM detected,skipping line = lab" )
                SmartParse.counter += 1
                return
            con_str = parse_line.split("*")[1]
            data = list(
                map("".join, zip(*[iter(con_str)] * 9))
            )  # splitting the line after * in 9 parts.
            order_dict[self.counter].set_data(data)
            SmartParse.counter += 1



for page_layout in extract_pages("inputs/SE_2023.pdf"):
    for element in page_layout:
        if isinstance(element, (LTTextBoxHorizontal)):
            if "PUNE" in element.get_text():
                continue  # avoiding unwanted blocks.
            else:
                for text_line in element:
                    SmartParse().ordered_parse(text_line.get_text())

xl = pd.ExcelWriter(
    "marks.xlsx",
    engine="xlsxwriter",
    engine_kwargs={"options": {"strings_to_numbers": True}},
)
pd.read_csv("generated/SE_2023_MARKS.csv").to_excel(xl, index=False)
xl.save()
