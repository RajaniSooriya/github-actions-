import os
import tempfile
from ppextractmodule import presentation

def test_process_function():
    # create a temporary directory to store extracted images
    with tempfile.TemporaryDirectory() as tmpdir:
        # create a test PowerPoint presentation file with three slides, each with one image
        test_file = os.path.join(tmpdir, "test.pptx")
        presentation = Presentation()
        for i in range(3):
            slide = presentation.slides.add_slide(presentation.slide_layouts[0])
            slide.shapes.add_picture("test_image.jpg", 0, 0)
        presentation.save(test_file)

        # run the process function on the test file
        process(test_file)

        # check that three image files were extracted
        assert len(os.listdir(tmpdir)) == 3

        # check that each extracted image file exists and has a non-zero size
        for i in range(3):
            image_file = os.path.join(tmpdir, f"image{i}.jpg")
            assert os.path.isfile(image_file)
            assert os.path.getsize(image_file) > 0
