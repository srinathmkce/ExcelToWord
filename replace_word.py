import csv
import xlrd
import docx
import os


class ConvertExcelToWord:
    def __init__(self):
        self.template = None
        self.data = {}
        self.fulltext = ""
        self.word_files = []

    def find_template(self):
        # This is where we decide which template to pick
        template = {}
        template["ticket"] = "template/passenger_list.csv"
        self.template = template["ticket"]
        print "Picking template: {0}".format(self.template)
        return 1

    def read_template(self):
        # this method is to read the excel files and store the key value pairs
        # we need xlrd package for this 
        # Install it using `pip install xlrd`
        with open(self.template) as csv_file:
            for row in csv_file:
                split_data = row.split(",")
                self.data[split_data[0]] = split_data[1]
                
        return 1

    def find_doc_files(self):
        # This method is to find the list of docs present in directory
        # Which we are going use and replace the text
        for root, dir, files in os.walk("word_files"):
            for f in files:
                if f.endswith(".docx"):
                    self.word_files.append(f)

        return 1


    def replace_in_doc(self):
        # This method is to read the doc files
        # we need docx package for this
        # Install it using `pip install docx`
        for word_file in self.word_files:
            word_file = "word_files/{0}".format(word_file)
            print "Converting : ", word_file
            doc = docx.Document(word_file)
            text = []
            for para in doc.paragraphs:
                inline = para.runs
                for j in range(0, len(inline)):
                    for k, v in self.data.items():
                        inline[j].text = inline[j].text.replace(str(k), str(v))

            doc.save(word_file)
            print "Conversion Successful"

        return 1

    def run(self):
        if not self.find_template():
            raise Exception("Unable to pick the right template")

        if not self.read_template():
            raise Exception("Unable to read Excel template")

        if not self.find_doc_files():
            raise Exception("Unable to find the word files")
     
        if not self.replace_in_doc():
            raise Exception("Unable to read Word file")
        
        return 1

        
    

if __name__ == "__main__":
    a = ConvertExcelToWord().run()



