# ledger_pdf_parser
A Project for automating the job of reading ledger pdfs (custom) and storing the data in excel file.  

### Files and what they do 
Each of the file is named with the convention of {year}_{sem}.py and should be used for that particular sem & year's ledger, they are NOT cross-compatible. Code was never produced for first year ledgers.

### How does the conversion work 
Each of the file has the same underlying structure. 
#### Student class
- The Student Class which stores the marks, name, SGPA, credits and seat no. The amount of targetted elements to be extracted per student in the excel file can be modified here according to the requirements.
- Has a tolist method to created printable sequence in excel file incase you want to change the ordering or generate ret_data() (formatted data) for the lab marks.

#### TheoryMarks Class
- Theory marks class has method set_marks to take control over processing the raw targetting elements and default value

#### LabMarks Class
- Stores marks for Practicals, Orals and Term Work separately. ret_data() processes the string marks and returns excel entry ready list output. Its a bit big to match edge cases specific to that year/sem's ledger.


#### CSVWriter Class
- Is the interface between csv file and the program for writing to the file. The init method contains the first and second row as headers for the csv file. 

#### SmartParse Class
- is where all the conversion happens using the objects of the classes described above. the ordered_parse() function takes a string of input line (a single line). The function maintains state across invocations using the order_dict which sets values corresponding to the attributes to complete the student object. Once a student object is completely filled (all attributes have recieved some values) we first write that student to the output and then clear it for the next student. 
- The function flow is as follows, first we ignore the lines which do not have target content using an if statement, this is highly necessary to avoid false counter increase which might lead to missing some values. Now based on where the counter is at we know which line we are currently looking at and the target values to extract. This needs to be manually matched with the sequence in the ledger. Fill the student attributes (marks) line by line. Now inside this a lot of the logic may change based on the ledger format, but generally you want to split the input line into 9 parts if it is a subject/lab marks line, further if the subject is an elective you want to fill in different attributes with some wizardry such as reserving an id for each of the subject type eg 410244A vs 410244B etc. Remember to increment the counter to the next subject and not just the subject type! refer line 316 in be_first.py for reference of what I'm trying to explain. Once we see SGPA in the line, it means that the student was completely read, now we just need to calculate their credits and write it to the file!


#### Pdf page extraction
At the end we run a basic loop over the input pdf using pdf miner.six package. Element by element, we only want to read LTTextBoxHorizontal and run ordered_parse on each text_line in that element.




#### How can this be extended/ future scope. 
I see the following three ways to make this more robust and directly usable by end clients without any coding language. Ranked by practicality

1. Build an edge case library for each of the year/sem combo, classify the given input ledger pdf into one of the possible format matches, if none match you are out of luck! Else you can parse for the matching type. This can be done with trial and error at the code level. 

2. Remove the dependence of sequencing and format from the code level, send it to the gui level for the user to make choice of target element per column, making it generic without having the user code anything out/ understanding underlying code structure. This according to me is the best possible outcome scenario there is for this type of application.

3. AI? I am not sure if it can ever help here with 100% accuracy but one could try to use it for format matching to provide best effort output, which is acceptable incase of a completely new/unseen format. 