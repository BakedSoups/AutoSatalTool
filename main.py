from pptx import Presentation
from pptx.util import Inches, Pt
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
            if slide.shapes[i].text == "[CLASS + NUMBER]":
                return i + 1 
    return None

def change_class_name(slide, class_name):
    for i in range(0, len(slide.shapes)):
        if hasattr(slide.shapes[i], "text"):
            if slide.shapes[i].text == "[CLASS + NUMBER]":
                slide.shapes[i].text = class_name
                return True
    return False

def clone_slide(prs, slide):
    new_slide = prs.slides.add_slide(slide.slide_layout)
    for shape in slide.shapes:
        if shape.has_text_frame:
            if shape.text_frame.text:
                new_shape = new_slide.shapes.add_textbox(
                    left=shape.left,
                    top=shape.top,
                    width=shape.width,
                    height=shape.height
                )
                new_shape.text_frame.text = shape.text_frame.text
                for paragraph in shape.text_frame.paragraphs:
                    new_paragraph = new_shape.text_frame.add_paragraph()
                    new_paragraph.text = paragraph.text
                    new_paragraph.font.bold = paragraph.font.bold
                    new_paragraph.font.italic = paragraph.font.italic
                    new_paragraph.font.size = Pt(paragraph.font.size.pt) if paragraph.font.size else Pt(18)
                    new_paragraph.font.color.rgb = paragraph.font.color.rgb if paragraph.font.color else (0, 0, 0)  # Default to black color if None
        elif shape.has_image_frame:
            image_path = shape.image.filename
            new_shape = new_slide.shapes.add_picture(
                image_path,
                left=shape.left,
                top=shape.top,
                width=shape.width,
                height=shape.height
            )
    return new_slide

def main():
    current_dir = os.getcwd()

    template_file_path = f"{current_dir}/template.pptx"
    template_presentation = Presentation(template_file_path)

    template_first_slide = template_presentation.slides[0]
    template_slide_text = extract_text_from_slide(template_first_slide)

    change_class_name(template_first_slide, "CS 101")

    new_presentation = Presentation()

    # Clone the modified slide to the new presentation
    clone_slide(new_presentation, template_first_slide)

    new_presentation.save(f"{current_dir}/new_presentation.pptx") 
    
if __name__ == "__main__":
    main()
