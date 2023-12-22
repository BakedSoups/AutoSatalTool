from pptx import Presentation
import os
import json

def pptx_to_xml(pptx_path, xml_path):
    presentation = Presentation(pptx_path)
    xml_data = []

    for i, slide in enumerate(presentation.slides):
        slide_data = {
            'layout_index': presentation.slide_layouts.index(slide.slide_layout),
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

        xml_data.append(slide_data)

    with open(xml_path, 'w') as xml_file:
        json.dump(xml_data, xml_file, indent=2)

def xml_to_pptx(xml_path, pptx_path):
    with open(xml_path, 'r') as xml_file:
        slides_data = json.load(xml_file)

    presentation = Presentation()

    for slide_data in slides_data:
        layout_index = slide_data['layout_index']

        try:
            new_layout = presentation.slide_layouts[layout_index]
        except IndexError:
            print(f"Error: Slide layout with index {layout_index} not found. Skipping slide.")
            continue

        slide = presentation.slides.add_slide(new_layout)

        for shape_data in slide_data['shapes']:
            shape_type = shape_data['type']

            if shape_type not in [1, 2, 3, 4]:
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

    presentation.save(pptx_path)

def main():
    current_dir = os.getcwd()
    template_pptx_path = os.path.join(current_dir, "template.pptx")
    xml_output_path = os.path.join(current_dir, "output.xml")
    converted_pptx_path = os.path.join(current_dir, "converted.pptx")

    pptx_to_xml(template_pptx_path, xml_output_path)
    xml_to_pptx(xml_output_path, converted_pptx_path)

if __name__ == "__main__":
    main()
