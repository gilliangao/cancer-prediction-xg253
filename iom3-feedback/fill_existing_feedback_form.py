from docx import Document

SRC = "/Users/gillian/Documents/New project 2/iom3-feedback/original_feedback_form.docx"
OUT = "/Users/gillian/Documents/New project 2/iom3-feedback/Registration Application Feedback Form - Completed.docx"


def mark_yes_no(table, yes=True):
    table.rows[0].cells[0].text = "Yes  X" if yes else "Yes"
    table.rows[0].cells[1].text = "No" if yes else "No  X"


doc = Document(SRC)

# Yes / no selections
mark_yes_no(doc.tables[0], yes=True)   # Q1
mark_yes_no(doc.tables[1], yes=False)  # Q2
doc.tables[1].rows[1].cells[0].text = (
    "If Yes, please comment further:\nThe workshop covered the key points clearly and was helpful in understanding the process."
)
doc.tables[1].rows[1].cells[1].text = doc.tables[1].rows[1].cells[0].text

mark_yes_no(doc.tables[2], yes=True)   # Q3 mentor support
doc.tables[2].rows[1].cells[0].text = "If Yes, was your Mentor from your organisation?\nNo"
doc.tables[2].rows[1].cells[1].text = "If Yes, was your Mentor from your organisation?\nNo"

mark_yes_no(doc.tables[3], yes=True)   # Q4
doc.tables[3].rows[1].cells[0].text = "If No, please comment further:\nN/A"
doc.tables[3].rows[1].cells[1].text = "If No, please comment further:\nN/A"

mark_yes_no(doc.tables[4], yes=True)   # Q5
doc.tables[4].rows[1].cells[0].text = "If No, please comment further:\nN/A"
doc.tables[4].rows[1].cells[1].text = "If No, please comment further:\nN/A"

mark_yes_no(doc.tables[5], yes=True)   # Q6

mark_yes_no(doc.tables[6], yes=False)  # Q7
doc.tables[6].rows[1].cells[0].text = (
    "If Yes, please comment further:\nI felt I had a fair opportunity to explain my experience and respond to the questions raised."
)
doc.tables[6].rows[1].cells[1].text = doc.tables[6].rows[1].cells[0].text

mark_yes_no(doc.tables[7], yes=True)   # Q8
doc.tables[7].rows[1].cells[0].text = (
    "If No, please comment further:\nN/A"
)
doc.tables[7].rows[1].cells[1].text = doc.tables[7].rows[1].cells[0].text

mark_yes_no(doc.tables[8], yes=False)  # Q9
doc.tables[8].rows[1].cells[0].text = (
    "If Yes, please comment further:\nThe remote interview ran smoothly and I did not experience any technical issues."
)
doc.tables[8].rows[1].cells[1].text = doc.tables[8].rows[1].cells[0].text

mark_yes_no(doc.tables[9], yes=True)   # Q10

# Q11 satisfaction
for row in doc.tables[10].rows:
    row.cells[1].text = ""
doc.tables[10].rows[0].cells[1].text = "X"

# Q12 comments and metadata
doc.tables[11].rows[0].cells[0].text = (
    "Overall, I was very satisfied with the process. The information provided in advance was helpful, "
    "and the interview panel were very professional, polite, and supportive throughout the interview."
)

# Fill the text lines after labels
doc.paragraphs[48].text = "Dr Xingran (Gillian) Gao"
doc.paragraphs[51].text = "711359"
doc.paragraphs[55].text = "18 May 2026"

doc.save(OUT)
print(OUT)
