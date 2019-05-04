from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH


document = Document()
currentDT = datetime.datetime.now()
print ("Enter The Company Name:")
x = input()

p1 = document.add_heading('Ryan Nguyen', 0)
p1.add_run('\n')
info_address = p1.add_run('Vancouver,BC   778-922-2808  quangineer@gmail.com')
info_address.font.size = Pt(12)
p1.add_run('\n')
info_github = p1.add_run('Github: https://github.com/quangineer')
info_github.font.size = Pt(12)

p2 = document.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p2.add_run(currentDT.strftime('%d, %b %Y'))

p3 = document.add_paragraph('Dear Hiring Manager')
p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

apply = 'I would like to apply for the position of Data Analytics at '
p4 = document.add_paragraph(apply + x)
p4.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

introduce = 'I am a graduate student in Data Analytics Udacity nanodegree majoring in Data Analysis, Practical Statistics, Data Wrangling, Data Visualization. I have applied techniques in projects such as Explore Weather Trends, Investigate a Dataset, Analyze A/B Test Results, Communicate Data Findings.'
p5 = document.add_paragraph(introduce)
p5.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

my_project = 'I have experience in Python and SQL as evidenced by my projects on Github such as “Finding the Right Restaurants” and “AutoGenCoverLetter”. In these projects, I develop my technical skill in Python programming and applying Python-docx library to develop my automatic cover letter.'
p6 = document.add_paragraph(my_project)
p6.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

my_business_background = 'In addition, I came from business background with a strong knowledge of finance and accounting. I believe a combination between business background and a savvy of data analytics will benefit the company. With regards to my experience in analysis, I used to work at PwC where I was assigned to collect data from clients in financial institution and food & beverage industry in the fields of employees compensation packages, organizational structure, customers gender population. These information was then processed to have client get a better understanding of their human resource and customer target.'
p7 = document.add_paragraph(my_business_background)
p7.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

my_attribute_and_contact = 'During my past working experience, I proved myself a hard-working, goal-focused person who was always willing to go extra miles to learn new skills and open-minded to any environment. I would like to thank you very much for considering my application, and I would greatly look forward to an opportunity to interview for the position offered by ' + x + '.I can be reached by phone at 778-922-2808 or by email at quangineer@gmail.com. All my projects can be viewed at: https://github.com/quangineer.'
p8 = document.add_paragraph(my_attribute_and_contact)
p8.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

p9 = document.add_paragraph('Sincerely,')

p10 = document.add_paragraph('Ryan Nguyen')



document.save('demo.docx')

