from ExcelRenderer import ExcelRenderer
import os

if __name__ == "__main__":
    name = "Bad_Apple"
    video = 'video\Bad_Apple.mp4'
    result_path = rf"C:\Users\OliverBLauritsen\Documents\personal\ExcelRenderer\result\{name}"
    renderer = ExcelRenderer(name=name, video=video)

    chapter_files = os.listdir(result_path)
    chapters = []

    for i in range(len(chapter_files)):
        chapters.append(renderer.load_workbook(os.path.join(result_path, f"chapter{i}.xlsx")))
    for chapter in chapters:
        renderer.play(chapter)