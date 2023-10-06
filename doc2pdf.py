import sys
import os
import win32com.client


def main(filename):
    try:
        wdFormatPDF = 17
        inputFile = os.path.abspath(filename)
        # if ending is docx or doc its removed and replaced by pdf
        outputFile = os.path.abspath(os.path.splitext(filename)[0] + ".pdf")
        # open word file
        word_app = win32com.client.Dispatch('Word.Application')
        document = word_app.Documents.Open(inputFile)
        # and save i as pdf
        document.SaveAs(outputFile, FileFormat=wdFormatPDF)
        document.Close()
        word_app.Quit()
    except Exception as ex:
        print(ex)
        exit(-1)
    
    print(f"Done with conversion of {filename}")
    os.system ('pause')


if __name__ == "__main__":
    dropped_file = sys.argv[1]

    print(f"Converting file {dropped_file} to pdf. This application needs Word to be installed on the computer.")
    main(dropped_file)
