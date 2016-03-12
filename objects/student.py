from mechanize import *
from bs4 import BeautifulSoup
from xlsxwriter import *

class Student(object):
    """A YRDSB student"""
    def __init__(self, username, password, cc_password):
        """Username(student #), school password, and Career Cruising password"""
        self.username = username
        self.password = password
        self.cc_password = cc_password
        self.useragent = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US) AppleWebKit/534.3 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/534.3')]

    def get_cc_marks(self):
        """Returns courses from Career Cruising and their respective coursecodes, 
        marks, and credits as values in a list e.g. { "COURSENAME" : ["COURSECODE", MARK, CREDITS] }"""
        br = Browser()
        br.addheaders = self.useragent
        br.open("http://public.careercruising.com/en/")
        br.select_form(nr=0)
        br["Username"] = "york-" + self.username
        br["Password"] = self.cc_password
        br.submit()
        try:
            course_page = br.follow_link(url="/../my/course-plan")
        except ParseError:
            raise ValueError("Career Cruising username or cc_password is incorrect")
        bs = BeautifulSoup(course_page, "html.parser")
        raw_courses = bs.findAll("a", { "class" : "ChooseCourse" })
        course_data = {}
        for course in raw_courses:
            try:
                course_data[str(course["coursecode"]).strip("u")] = [str(course["coursename"]).strip("u"), float(course["grademark"]), float(course["creditvalue"])]
            except ValueError:
                None
        return course_data
        
    def unofficial_transcript(self):
        """One of the more interesting/useful functions here, this will
        save all past courses, including marks and credit values, into a
        pre-formatted excel spreadsheet"""
        course_data = self.get_cc_marks()
        transcript = Workbook("transcript.xlsx")
        sheet = transcript.add_worksheet()
        title_format = transcript.add_format({ "bold" : True, "fg_color" : "gray", "align" : "center", "border" : 1 })
        head_format = transcript.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "#d3d3d3" })
        data_format = transcript.add_format({ "bold" : True, "align" : "center", "border" : 1 })
        sheet.set_column(1, 3, 13)
        sheet.set_column(0, 0, 45)
        sheet.merge_range("A1:D1", "Transcript for %s" % self.username, title_format)
        for heading, location in (["Course Name", "A2"], ["Course Code", "B2"], ["Final Mark", "C2"], ["Credit Value", "D2"]):
            sheet.write(location, heading, head_format)
        row = 2
        for coursecode, data in course_data.items():
            sheet.write(row, 0, data[0], data_format)
            sheet.write(row, 1, coursecode, data_format)
            sheet.write(row, 2, data[1], data_format)
            sheet.write(row, 3, data[2], data_format)
            row+=1
    
    def get_ta_marks(self):
        """Returns courses from TeachAssist and their respective marks as a dictionary"""
        br = Browser()
        br.addheaders = self.useragent
        br.open("https://ta.yrdsb.ca/yrdsb/")
        br.select_form(name = "loginForm")
        br["username"] = self.username
        br["password"] = self.password
        mark_page = br.submit()
        bs = BeautifulSoup(mark_page, "html.parser")
        if bs.find("font", { "color" : "red" }) != None:
            raise ValueError("YRDSB username or password is incorrect")
        course_data = bs.findAll("table")[1].findAll("td", align=None)
        mark_data = bs.findAll("table")[1].findAll("td", align="right")
        courses = []
        marks = []
        for course in course_data:
            courses.append(str(course.text.split()[0]))
        for mark in mark_data:
            marks.append(str(mark.text.strip("\n \t")))
        course_info = dict(zip(courses, marks))
        return course_info
    
    def output_ta_marks(self):
        """Takes TeachAssist test/quiz/assignment marks and weightings for each current 
        course and saves them into a MS Excel spreadsheet. Note that this will output 
        your overall average regardless of whether your teacher has set it as visible,
        provided that marks and weights ARE visible"""
        br = Browser()
        br.addheaders = self.useragent
        br.open("https://ta.yrdsb.ca/yrdsb/")
        br.select_form(name = "loginForm")
        br["username"] = self.username
        br["password"] = self.password
        mark_page = br.submit()
        bs = BeautifulSoup(mark_page, "html.parser")
        if bs.find("font", { "color" : "red" }) != None:
            raise ValueError("YRDSB username or password is incorrect")
        course_links = bs.findAll("table")[1].findAll("a")
        markbook = Workbook("markbook.xlsx")
        for link in course_links:
            course_page = br.follow_link(url=link["href"])
            bs = BeautifulSoup(course_page, "html.parser")
            course = bs.find("h2").text
            marksheet = markbook.add_worksheet()
            title_format = markbook.add_format({ "bold" : True, "fg_color" : "gray", "align" : "center", "border" : 1 })
            knowledge_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "ffffaa" })
            knowledge_avg_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "ffffaa", "num_format" : "0.00%"  })
            thinking_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "c0fea4" })
            thinking_avg_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "c0fea4", "num_format" : "0.00%" })
            communication_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "afafff" })
            communication_avg_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "afafff", "num_format" : "0.00%" })
            application_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "ffd490" })
            application_avg_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1, "fg_color" : "ffd490", "num_format" : "0.00%" })
            head_format = markbook.add_format({ "bold" : True, "align" : "center", "border" : 1 })
            mark_format = markbook.add_format({ "align" : "center", "border" : 1, "num_format" : "0.00%" })
            weight_format = markbook.add_format({ "align" : "center", "border" : 1 })
            knowledge_weight = bs.find("tr", { "bgcolor" : "#ffffaa" }).find("td", { "align" : "right" }).text
            thinking_weight = bs.find("tr", { "bgcolor" : "#c0fea4" }).find("td", { "align" : "right" }).text
            communication_weight = bs.find("tr", { "bgcolor" : "#afafff" }).find("td", { "align" : "right" }).text
            application_weight = bs.find("tr", { "bgcolor" : "#ffd490" }).find("td", { "align" : "right" }).text
            marksheet.set_column(0, 7, 13)
            for heading, spread, form in (["Knowledge(weight=%s)" % knowledge_weight, "A2:B2", knowledge_format], ["Thinking(weight=%s)" % thinking_weight, "C2:D2", thinking_format], ["Communication(weight=%s)" % communication_weight, "E2:F2", communication_format], ["Application(weight=%s)" % application_weight, "G2:H2", application_format]):
                marksheet.merge_range(spread, heading, form)
            for column in [0, 2, 4, 6]:
                marksheet.write(3, column, "Mark", head_format)
                marksheet.write(3, column+1, "Weight", head_format)
            knowledge = bs.findAll("td", { "bgcolor" : "ffffaa", "align" : "center", "id" : None })
            k_marks = []
            k_marksum = 0
            k_weightsum = 0
            for cell in knowledge:
                words = cell.text.split()
                if len(words) > 0:
                    try:
                        num = float(words[0].strip("u"))
                        den = float(words[2].strip("u"))
                        weight = float(words[5].strip("u weight="))
                        k_marks.append([num/den, weight])
                        k_marksum+=(num/den)*weight
                        k_weightsum+=weight
                    except ValueError:
                        None
            row = 4
            col = 0
            marksheet.merge_range("A3:B3", k_marksum/k_weightsum, knowledge_avg_format)
            for mark, weight in k_marks:
                marksheet.write_number(row, col, mark, mark_format)
                marksheet.write_number(row, col+1, weight, weight_format)
                row+=1
            thinking = bs.findAll("td", { "bgcolor" : "c0fea4", "align" : "center", "id" : None })
            t_marks = []
            t_marksum = 0
            t_weightsum = 0
            for cell in thinking:
                words = cell.text.split()
                if len(words) > 0:
                    try:
                        num = float(words[0].strip("u"))
                        den = float(words[2].strip("u"))
                        weight = float(words[5].strip("u weight="))
                        t_marks.append([num/den, weight])
                        t_marksum+=(num/den)*weight
                        t_weightsum+=weight
                    except ValueError:
                        None
            row = 4
            col = 2
            marksheet.merge_range("C3:D3", t_marksum/t_weightsum, thinking_avg_format)
            for mark, weight in t_marks:
                marksheet.write_number(row, col, mark, mark_format)
                marksheet.write_number(row, col+1, weight, weight_format)
                row+=1
            communication = bs.findAll("td", { "bgcolor" : "afafff", "align" : "center", "id" : None })
            c_marks = []
            c_marksum = 0
            c_weightsum = 0
            for cell in communication:
                words = cell.text.split()
                if len(words) > 0:
                    try:
                        num = float(words[0].strip("u"))
                        den = float(words[2].strip("u"))
                        weight = float(words[5].strip("u weight="))
                        c_marks.append([num/den, weight])
                        c_marksum+=(num/den)*weight
                        c_weightsum+=weight
                    except ValueError:
                        None
            row = 4
            col = 4
            marksheet.merge_range("E3:F3", c_marksum/c_weightsum, communication_avg_format)
            for mark, weight in c_marks:
                marksheet.write_number(row, col, mark, mark_format)
                marksheet.write_number(row, col+1, weight, weight_format)
                row+=1
            application = bs.findAll("td", { "bgcolor" : "ffd490", "align" : "center", "id" : None })
            a_marks = []
            a_marksum = 0
            a_weightsum = 0
            for cell in application:
                words = cell.text.split()
                if len(words) > 0:
                    try:
                        num = float(words[0].strip("u"))
                        den = float(words[2].strip("u"))
                        weight = float(words[5].strip("u weight="))
                        a_marks.append([num/den, weight])
                        a_marksum+=(num/den)*weight
                        a_weightsum+=weight
                    except ValueError:
                        None
            row = 4
            col = 6
            marksheet.merge_range("G3:H3", a_marksum/a_weightsum, application_avg_format)
            for mark, weight in a_marks:
                marksheet.write_number(row, col, mark, mark_format)
                marksheet.write_number(row, col+1, weight, weight_format)
                row+=1
            average = (k_marksum/k_weightsum)*float(knowledge_weight.strip("%)")) + (t_marksum/t_weightsum)*float(thinking_weight.strip("%")) + (c_marksum/c_weightsum)*float(communication_weight.strip("%")) + (a_marksum/a_weightsum)*float(application_weight.strip("%"))
            round_average = str(float("%.2f" % average))
            marksheet.merge_range("A1:H1", course+" (current mark="+round_average+"%)", title_format)

    def current_average(self):
        """Gets current average based on TeachAssist marks"""
        courses = self.get_ta_marks()
        marks = []
        for mark in courses.values():
            for word in mark.split():
                try:
                    marks.append(float(word.strip("%")))
                except ValueError:
                    None
        average = sum(marks)/len(marks)
        round_average = float("%.2f" % average)
        return round_average
    
    def cumulative_average(self):
        """Gets cumulative average based on Career Cruising marks"""
        marks = self.get_cc_marks().values()
        credits = 0
        markSum = 0
        for mark in marks:
            markSum += mark[1]*mark[2]
            credits += mark[2]
        average = markSum/credits
        round_average = float("%.2f" % average)
        return round_average
        
    def gradelevel_average(self, grade):
        """Returns the average of all marks of the supplied grade level (9, 10, 11, 12, as integers)
        *NOTE: this does not return averages by year, but by the academic level of the courses"""
        course_data = self.get_cc_marks()
        markSum = 0
        credits = 0
        for course in course_data.keys():
            if grade == 9:
                if course[3] == "1":
                    markSum += course_data[course][1]*course_data[course][2]
                    credits += course_data[course][2]
            elif grade == 10:
                if course[3] == "2":
                    markSum += course_data[course][1]*course_data[course][2]
                    credits += course_data[course][2]
            elif grade == 11:
                if course[3] == "3":
                    markSum += course_data[course][1]*course_data[course][2]
                    credits += course_data[course][2]
            elif grade == 12:
                if course[3] == "4":
                    markSum += course_data[course][1]*course_data[course][2]
                    credits += course_data[course][2]
            else:
                raise ValueError("Please specify an integer grade of 9, 10, 11, or 12")
        average = markSum/credits
        round_average = float("%.2f" % average)
        return round_average
        
    def credits(self):
        """Returns the number of completed credits to date"""
        marks = self.get_cc_marks().values()
        credits = 0
        for mark in marks:
            credits += mark[2]
        return credits