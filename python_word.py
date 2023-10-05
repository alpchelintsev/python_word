# pip install pypiwin32

# StringReplace replaces the text from variable SearchStr by text from ReplaceStr,
# ReplaceAll means to replace all occurrences,
# W is an object of Word.Application
def StringReplace(W, SearchStr, ReplaceStr, ReplaceAll):
  W.Selection.Find.ClearFormatting()
  W.Selection.Find.Replacement.ClearFormatting()
  if ReplaceAll:
    W.Selection.Find.Execute(SearchStr, False, False, False, False, False, True, 1, True, ReplaceStr, 2)
  else:
    W.Selection.Find.Execute(SearchStr, False, False, False, False, False, True, 1, True, ReplaceStr, 1)

# InsertFile inserts a document from a file with name in FN
def InsertFile(W, FN):
  W.Selection.EndKey(Unit = 6)
  W.Selection.InsertBreak(Type = 7)
  W.Selection.InsertFile(FileName = FN)

import win32com.client
Word = win32com.client.Dispatch("Word.Application")
Word.Documents.Open("F:\\my.docx")
StringReplace(Word, "{numstud}", "001", True)
InsertFile(Word, "F:\\doc2.docx")
Word.ActiveDocument.SaveAs(FileName = "F:\\1.docx", FileFormat = 16)
Word.ActiveDocument.Close()
Word.Quit()
