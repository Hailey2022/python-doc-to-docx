from win32com import client as wc
import os

path = r'D:\'
for top, dirs, files in os.walk(path):
    for file in files:
        file = os.path.join(top, file)
        if file.endswith('.doc'):
            word = wc.Dispatch("Word.Application")
            doc = word.Documents.Open(file)
            doc.SaveAs("{}x".format(file), 12)
            doc.Close()
            os.remove(file)
            word.Quit()
        print("Done!")
