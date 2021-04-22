import docx
import csv
import pathlib
from os import path

if __name__ == "__main__":
    with open("respostas.csv") as f:
        reader = csv.DictReader(f)
        for line in reader:
            doc = docx.Document("rad.docx")
            #print("Checking: " + line["Date"] + ".docx")
            if line["SIAPE"] != "":
            	print(line["SIAPE"])
                #doc.save(line["Date"] + ".docx")"
