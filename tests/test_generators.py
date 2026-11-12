import unittest
from io import BytesIO
import pandas as pd
from openpyxl import load_workbook
from docx import Document
from scheduler.data import parse_meetings, normalize
from scheduler.scc import make_scc
from scheduler.word_signs import make_word_signs

class GeneratorsTest(unittest.TestCase):
    def setUp(self):
        self.df = normalize(pd.DataFrame([
            {'room':'SCC 101','subject':'SEIS','course_id':'630-01','course_time':'02/01 - 05/20; M 5:30pm-8:30pm @ SCC 101'},
            {'room':'OWS 200','subject':'ETLS','course_id':'601-01','course_time':'02/01 - 05/20; Th 02/15 9:00am-10:00am @ OWS 200'},
            {'room':'OWS 200','subject':'MATH','course_id':'101-01','course_time':'02/01 - 05/20; W 10:00am-11:00am @ OWS 200'},
        ]))
    def test_meetings(self):
        m = parse_meetings(self.df.iloc[1]['course_time'])
        self.assertEqual((m[0]['day'],m[0]['date'],m[0]['room']),('THU','02/15','OWS 200'))
    def test_excel(self):
        data, n = make_scc(self.df)
        self.assertEqual(n,1)
        wb = load_workbook(BytesIO(data))
        self.assertIn('SCC 101', wb.active['A1'].value)
    def test_separate_room_worksheets(self):
        df = normalize(pd.DataFrame([
            {"room":"SCC 101", "subject":"SEIS", "course_id":"630-01", "course_time":"02/01 - 05/20; M 5:30pm-8:30pm @ SCC 101"},
            {"room":"SCC 202", "subject":"SEIS", "course_id":"632-01", "course_time":"02/01 - 05/20; T 8:00am-9:00am @ SCC 202"},
            {"room":"SCC 202", "subject":"MATH", "course_id":"101-01", "course_time":"02/01 - 05/20; T 9:00am-10:30am @ SCC 202"},
        ]))
        data, n = make_scc(df, "Spring 2027")
        wb = load_workbook(BytesIO(data))
        self.assertEqual(n, 2)
        self.assertEqual(wb.sheetnames, ["SCC 101", "SCC 202"])
        self.assertIn("SEIS 630", str([c.value for row in wb["SCC 101"] for c in row]))
        self.assertNotIn("MATH 101", str([c.value for row in wb["SCC 101"] for c in row]))
        self.assertIn("MATH 101", str([c.value for row in wb["SCC 202"] for c in row]))

    def test_word(self):
        docs = make_word_signs(self.df)
        self.assertEqual(docs['OWS'][1],1)
        doc = Document(BytesIO(docs['OWS'][0]))
        values = ' '.join(p.text for p in doc.paragraphs)
        self.assertIn('MATH 101-01',values)
        self.assertIn('SPECIAL DATES',values)
        self.assertIn('OWS 200',values)

if __name__ == '__main__':
    unittest.main()

class OriginalWordLayoutTest(unittest.TestCase):
    def test_word_groups_evening_and_special_and_includes_other_subjects(self):
        df = normalize(pd.DataFrame([
            {'room':'OSS 110','subject':'SEIS','course_id':'630-01','course_time':'02/01 - 05/20; M 5:30pm-8:30pm @ OSS 110'},
            {'room':'OSS 110','subject':'HIST','course_id':'100-01','course_time':'02/01 - 05/20; M 9:00am-10:00am @ OSS 110; W 9:00am-10:00am @ OSS 110'},
            {'room':'OSS 110','subject':'ETLS','course_id':'746-01','course_time':'02/01 - 05/20; Th 03/17 9:00am-10:00am @ OSS 110'},
        ]))
        docs = make_word_signs(df)
        self.assertEqual(docs['OSS'][1],1)
        text = '\n'.join(p.text for p in Document(BytesIO(docs['OSS'][0])).paragraphs)
        for required in ('OSS 110','MW','HIST 100-01','EVENINGS','SEIS 630-01','SPECIAL DATES','ETLS 746-01'):
            self.assertIn(required,text)
