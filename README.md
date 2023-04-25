# EXTRACT TEXT AND CREATE TABLE FROM IMAGE
#### Video Demo:  https://youtu.be/h_Wn_rFLUfs
## Description:

This program is written in Python with the purpose of extracting text from images and transform into table data. The program contains 2 files, the OCR_detector.py file used pytesseract, scikit-learn modules to identify text on image then save the data into .csv file, and the gui.py file create a user-friendly UI for the program.

### OCR_dectector.py
This file was originally created by Adrian Rosebrock on the website: https://pyimagesearch.com/2022/02/28/multi-column-table-ocr/. I found the code when I was looking for quick Python script to extract text from a financial statement image. 

In the original code, the creator use pytesseract module to extract every text that appear on the image. He then extract the x coordinates of every text and create `AgglomerativeClustering` model from the scikit-learn module to classify the x coordinates into groups. In other words, the model put texts that are vertically close to each other into columns. Each group (column) now have a number of text with different position in heights, which represents different rows in a column, the group is then sorted in y coordinates and concat to a numpy dataframe. The result is a dataframe containing columns and rows of the text from the image.

But this approach has a few drawbacks:
- Some images have columns that have blank rows, which lead to resulting dataframe have columns in different length, texts that should be on the same row is now mixed up and have to be manually edit. 

- Another feature was not utilized in the original code is the second way to use AgglomerativeClustering model provided by scikit-learn. The `AgglomerativeClustering` model has 2 ways to use, either input a distance-threshold variable or a n-cluster variable. The first method use distance-threshold input as a cut-off length to decide whether a text belongs to this column or another. And this length is measured in pixel, Because user don't always know exactly the width and length of an image in pixel, and images can varies in size, so it is hard to judge what is a good distance threshold. Compare this to the second method, which take a number n and the model will try to group the text into n groups. It means that we can specify the number of columns and rows to improve the structure of data table. Even though this method is more intunitive for the user than the first method, but it was not implemented.

From the original file, I have modified the main function to fix the drawbacks, introduced the ability to specify the number of rows or columns and a few more features such as removing unwanted characters, 

### gui.py
This file uses tkinter module to create the GUI in OOP (Object Oriented Programming) design. The reason I chose to use tkinter is because it is standard Python library, which is fast to implement and have community supports to help me resolve some of the problems I have encountered.

I wanted to create a GUI for this program so that not only myself can use it, but I can also share it to my friends and colleagues, of which most of them are not familar with using command line to execute the program. Moreover, the GUI also helps anyone interesting in my project in trying it out.

The main features of this file are: receiving input from different sources, adjusting options of the `AgglomerativeClustering` model, and using Microsoft Excel to open the .csv file.

The GUI has 3 button for input: input from computer file by opening a file dialog for the user to choose the image from, or input from clipboard by saving the image on the clipboard, or using the previous extracted image as the input. All of these types of input saved the image to a temporary jpg image and then open it using Pillow library.

The options available to the user to adjust `AgglomerativeClustering` model include: 
- Minimum number of columns/rows
- Distance thresh-hold for column/row
- Specific number of columns/rows
- Characters user want to remove from the text.

Finally, there is an option to open the .csv output in Excel right after the user had provided the image and OCR_detector.py has finished extracting data. But this option requires user to have Excel already installed, which I understand is not the case for everyone (Linux users for example) so if opening Excel raise an error, there is an try except clause to catch the error and notify with the user that they don't have Excel.

### Save user's settings
I found it is more convenient if the program remembers the value for each option, so that user don't have to adjust these whenever they want to use the program. There are many file type can help to achieve this, such as .ini, .yaml, .xml, ... I chose .ini file type for its simplicity and quick time to implement.

### Virtual envinroment (venv)
Because this program uses many library, some are not standard Python library, and the syntax of each library could change in a future major update, I chose to create venv so the program will be stable if myself or anyone wants to use the program in the future. This is where requirements.txt file comes in, the file contains the name and the version of the library being used in the program. Now Python will now exactly which version of each package to use, preventing new bugs from appearing.
