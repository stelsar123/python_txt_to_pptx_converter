# Requirements:
This script was build by using **Python 3.6** and the [python-pptx library](https://python-pptx.readthedocs.io/). It is recommended to study the library's documentation in order to use it more efficiently. This script can convert multiple .txt files, that follow a suggested format, to multiple .pptx files. Please see the files/test.txt file to understand the format. Watch out not to include blank spaces in the end of lines, as it can sometimes result in blank PowerPoint pages.

# How to use:

**Step 1:** Create one or more .txt files to convert to .pptx. As mentioned, please follow the suggested format. </br>
**Step 2:** Move your .txt files, to the /files folder. </br>
**Step 3:** Create a template .pptx slide, so all the others can inherit its properties, and place it in the project folder. There is an example input.pptx template in this repository. </br>
**Step 4:** Insert the path of your .txt files in the PATH variable, run the main.py script and insert the name of the template.</br>
**Step 5:** Debug by adjusting all the necessary parameters and values in the main.py script. Pay close attention at the Inches values in the add_slide function. </br>
**Step 6:** Repeat step 4 until satisfied with the result :)


