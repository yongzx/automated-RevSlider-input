import os
import zipfile

class FileOutput:

    def __init__(self):
        self.text = "slider_export.txt"

    def create_txt(self, content):
        with open(self.text, "w+") as f:
            f.truncate()
            f.write(content)
            print("Generated: {0}".format(self.text))

    def create_zip(self, num_slides):
        zf = zipfile.ZipFile('i{0}.zip'.format(num_slides), 'w')
        try:
            print('Zip {0} into i{1}.zip\n'.format(self.text,num_slides))
            zf.write(self.text)
        finally:
            zf.close()

    def remove(self):
        os.remove(self.text)
