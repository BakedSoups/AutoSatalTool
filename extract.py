from pptx import Presentation
import os 

def create_presentation_from_text(text):
    presentation = Presentation()

    # Split the text into slides
    slides_text = text.split("\n\n")

    for i, slide_text in enumerate(slides_text):
        # Create a new slide
        slide_layout = presentation.slide_layouts[0]  # You can change the layout as needed
        slide = presentation.slides.add_slide(slide_layout)

        # Add text to the slide
        shapes = slide.shapes
        textbox = shapes.add_textbox(left=0, top=0, width=1, height=1)
        frame = textbox.text_frame
        frame.text = slide_text.strip()

    return presentation

def save_presentation(presentation, file_path):
    presentation.save(file_path)

def main():
    # Specify the path to your text file
    currentdir = os.getcwd()
    text_file_path = "path/to/your/output_text.txt"

    # Read the text from the file
    with open(text_file_path, 'r', encoding='utf-8') as file:
        text = file.read()

    # Create a PowerPoint presentation from the text
    new_presentation = create_presentation_from_text(text)

    # Specify the path to save the new PowerPoint file
    new_pptx_file_path = "path/to/your/new_presentation.pptx"

    # Save the new PowerPoint presentation
    save_presentation(new_presentation, new_pptx_file_path)

if __name__ == "__main__":
    main()
