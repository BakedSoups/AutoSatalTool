from pptx import Presentation
import os 
import json
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
def pptx_to_json(pptx_path, json_path):
    presentation = Presentation(pptx_path)

    # Extract relevant information from each slide
    slides_data = []
    for slide in presentation.slides:
        slide_data = {
            'layout': slide.slide_layout.name,
            'shapes': []
        }

        for shape in slide.shapes:
            shape_data = {
                'type': shape.shape_type,
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height,
                'text': shape.text if hasattr(shape, 'text') else None
            }
            slide_data['shapes'].append(shape_data)

        slides_data.append(slide_data)

    # Write the extracted data to a JSON file
    with open(json_path, 'w') as json_file:
        json.dump(slides_data, json_file, indent=2)
def json_to_pptx(json_path, pptx_path):
    with open(json_path, 'r') as json_file:
        slides_data = json.load(json_file)

    presentation = Presentation()

    for slide_data in slides_data:
        layout_name = slide_data['layout']

        # Check if the layout exists
        if layout_name not in [layout.name for layout in presentation.slide_layouts]:
            # If not, check if it's a default layout and add it
            default_layouts = ['Title Slide', 'Title and Content', ...]  # Add default layout names
            if layout_name in default_layouts:
                presentation.slide_layouts.add_layout_by_name(layout_name)
            else:
                print(f"Error: Slide layout '{layout_name}' not found. Skipping slide.")
                continue

        slide_layout = presentation.slide_layouts[layout_name]
        slide = presentation.slides.add_slide(slide_layout)

        for shape_data in slide_data['shapes']:
            shape_type = shape_data['type']

            # Check if shape type is valid (e.g., 1 is for rectangles)
            if shape_type not in [1, 2, 3, 4, ...]:  # Add valid shape types
                print(f"Error: Invalid shape type '{shape_type}'. Skipping shape.")
                continue

            shape = slide.shapes.add_shape(
                shape_type,
                shape_data['left'],
                shape_data['top'],
                shape_data['width'],
                shape_data['height']
            )

            if 'text' in shape_data and shape.has_text_frame:
                shape.text_frame.text = shape_data['text']

    # Save the new PowerPoint presentation
    presentation.save(pptx_path)



def main():
    currentdir = os.getcwd() 
    pptx_file_path = f"{currentdir}/template.pptx"
    pptx_output_test = f"{currentdir}/output.json"
    pptx_output = f"{currentdir}/output.pptx"
    presentation = Presentation(pptx_file_path)
    first_slide = presentation.slides[0]
    

    slide_text = extract_text_from_slide(first_slide)
    print("Text from the first slide:")
    print(slide_text)

    pptx_to_json(pptx_file_path, pptx_output_test)
    json_to_pptx(pptx_output_test, pptx_output)
    print("Done!")


if __name__ == "__main__":
    main()