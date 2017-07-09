from package import *

if __name__ == "__main__":
    video = ExcelData()
    topic = TopicData()
    template = Output("template_head.txt",
                      "template_body1.txt", "template_end.txt")

    # obtain excel file and sheet name
    excel_name, excel_file = readFile()
    sheet_name, sheet = readSheet(excel_file)
    slider = readSlider()
    # processing information in the particular sheet
    video.read(topic, sheet)

    # output
    print("Total slides: ", len(video))
    template.export(excel_name, sheet_name, video, slider)
