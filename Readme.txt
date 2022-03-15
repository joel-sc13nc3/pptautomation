# Install requirements

1. Download and install python from https://www.python.org/downloads/
    1.a When installing python, select the option that sets the Environment variables
    1.b check python was properly installed by opening the cmd and run  the command [ python --version ] (do not include brackets). The python version should appear here
2. Install Libraries
    2.a Open CMD and install the libraries by typing [pip install name_of_variable] (do not include brackets).
    This needs to be done for the following list of libraries: pandas, numpy, thinkcell,regex,openpyxl

Note: This needs to be done only the first time python is installed


# Step by step to run the code
1. Add the analysis sheet (latest version) to the "Input/Analysis_Sheet"
2. Add the dataloader (latest version) to the "Input/Data_Loader"
3. Fill the Selections template located in the "Input/Naming"
4. Open the cmd
5. Select the source of the boxapp folder
    5.a Additional instructions to navigate to the source path
        5.a.1 type cd "full_path" (example: cd "C:\Users\Joel Ramirez\downloads\pptautomation")
        5.a.2 You need to find the main.py under the pptautomation folder
6. run the following instruction python main.py
7. Wait until following message appears "The ppttc file has been succesfully created"
8. The ppttc will be located in the output of the pptautomation folder