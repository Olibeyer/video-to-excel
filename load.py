from ExcelRenderer import ExcelRenderer
import os 

if __name__ == "__main__":
    name = "Bad_Apple"
    video = 'video\Bad_Apple.mp4'
    result_path = rf".\result\{name}"
    renderer = ExcelRenderer(name=name, video=video)
    
    #renderer.load()
    chapter_files = os.listdir(result_path)
    chapters = []

    for file in chapter_files:
        chapters.append(renderer.load_workbook(os.path.join(result_path, file)))
    for chapter in chapters:
        renderer.format_workbook()
