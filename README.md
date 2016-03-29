# YRDSBMarkTools

I sorta just got bored over the course of a few days and so I ended up throwing this together. This library can be used to interact with mark information pulled from YRDSB TeachAssist and Career Cruising. I'm not entirely sure what it might be useful for, but enjoy if you do use it!

###Sample Usage

####Create New Student
```python
#import student class and functions from objects/student.py
from objects.student import *

#create new student object with student number, YRDSB password, and Career Cruising password
bob = Student("123456789", "password", "cc_password")
```

####Call Student Functions

```python
#this function creates an excel spreadsheet with all past courses, including marks and credits
bob.unofficial_transcript()
```
Sample Output:
![Sample1](https://github.com/haydengunraj/YRDSBMarkTools/blob/master/samples/Sample1.png?raw=true "Sample1")



### Dependencies

- [Python 2.7](https://www.python.org/downloads/)
- [BeautifulSoup4](http://www.crummy.com/software/BeautifulSoup/)
- [XlsxWriter](http://xlsxwriter.readthedocs.org/)
- [Mechanize] (http://wwwsearch.sourceforge.net/mechanize/)
