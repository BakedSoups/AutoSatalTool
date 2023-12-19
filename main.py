from pptx import Presentation
import os 

def extract_text_from_slide(slide):
    text = ""
    for i, shape in enumerate(slide.shapes):
        if hasattr(shape, "text"):
            text += f"Shape {i+1}: {shape.text}\n"
    return text.strip()

def find_class_name(slide):
    for i in range(0, len(slide.shapes)):
        if hasattr(slide.shapes[i], "text"):
            print(slide.shapes[i].text)
            if slide.shapes[i].text == "[CLASS + NUMBER]":
                return i +1 
    return False

def change_class_name(slide, classPos,  new_text): 
    if presentation.slide[0].shapes[classPos].text != "":
        presentation.slide[0].shapes[classPos].text = new_text
       
def copy_slides(source_presentation, destination_presentation):
    for slide in source_presentation.slides:
        destination_presentation.slides.add_slide(slide)


def main():
    current_dir = os.getcwd()
    template_file_path = f"{current_dir}/template.pptx"
    template_presentation = Presentation(template_file_path)
    output_presentation = Presentation()
    copy_slides(template_presentation, output_presentation)
    first_slide = output_presentation.slides[0]
    slide_text = extract_text_from_slide(first_slide)
    
    print("Text from the first slide:")
    print(slide_text)

    # Find the position of the class name shape
    class_pos = find_class_name(first_slide)

    if not class_pos:
        raise Exception("Class name not found")
    else:
        # Change the class name on the first slide
        change_class_name(first_slide, class_pos, "CS 101")

    # Save the modified presentation
    output_file_path = f"{current_dir}/output.pptx"
    output_presentation.save(output_file_path)

if __name__ == "__main__":
    main()
