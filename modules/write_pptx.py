from pptx import Presentation
import os

# structure of input data:
# [['Word', ('path', 'name', slidenr)], ['Word', ('path', 'name', slidenr)], ['Word', ('path', 'name', slidenr)], ['Word', ('path', 'name', slidenr)]]

current_dir = os.path.dirname(__file__)
template_path = os.path.join(current_dir, "template.pptx")
filename = "search_results.pptx"

testlist = []
for i in range (9):
    string: str = f"Slide number {i} in Presentation Name"
    testlist.append(string)

def create_presentation(found_list):
    
    prs = Presentation(template_path)

    """Create a slide"""
    # pick slide layout
    title_slide_layout = prs.slide_layouts[0]
    # add a slide
    slide = prs.slides.add_slide(title_slide_layout)
    # add and fill title placeholder
    title = slide.shapes.title
    title.text = "Word found" # type: ignore

    """Iterate through the list of found results and fill placeholders accordingly"""
    # Highest placeholder ID: 18, Placeholder Name: Text Placeholder 17
    # Starting at 3 because that's the id of the first placeholder
    placeholder_counter: int = 3
    for result in found_list:
        for placeholder in slide.placeholders:
            if placeholder.shape_id == placeholder_counter:
                text_frame = placeholder.text_frame # type: ignore
                p = text_frame.paragraphs[0]
                run = p.add_run()
                run.text = result
                # this temporary link needs to be exchanged for a constructed hyperlink to the real file
                run.hyperlink.address = 'C:/Users/Fabian Timm/Documents/Github/pptxSearch/test/testpres.pptx#10'
        placeholder_counter += 1

    prs.save(filename)

if __name__ == "__main__":
    create_presentation(testlist)
    os.startfile("search_results.pptx") 