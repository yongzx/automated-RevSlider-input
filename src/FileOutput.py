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

    def create_zip(self, slider, num_slides):
        zf = zipfile.ZipFile('{0}{1}.zip'.format(slider,num_slides), 'w')
        try:
            print('Zip {0} into {1}{2}.zip\n'.format(self.text,slider,num_slides))
            zf.write(self.text)
        finally:
            zf.close()

    def remove(self):
        os.remove(self.text)
