## Synopsis
This project intends to automate the creation of video layers in a Wordpress website through Wordpress RevSlider (Silder Revolution) plugin. 
The python file reads the excel document in which the raw information of Youtube videos (name, ID, duration) is stored and generates a text file that is ready to be imported into RevSlider.

## Motivation
My friend actively uses RevSlider to create video layers in his website (http://learnah.org/) from the video information (name, ID, duration) saved in an excel document. Hence, it becomes inconvenient for him to transfer the information from the excel document into the RevSlider plugin to create multiple video layers. 

## Instructions
1. Install OpenPyxl library `$ pip install openpyxl`.
2. Put python file *.py*, folder "*Templates*" , and your excel document "*.xlsx*" into the same folder.
3. Run the python file and input the name of the excel document which you want the python script to open. Remember to include .xlsx in your input.  
