# Python for CPLEX Readme

The files included here are required in order to replicate our results. The original data used for the study is not uploaded.

We cover 3 sections in this readme document:
- The SOFTWARE required to run the python code,
- The FILES needed to run the python and cplex code,
- Some remarks about the python code.


FIRST PART: SOFTWARE

- IBM ILOG Optimization Studio should be installed (https://www.ibm.com/support/pages/downloading-ibm-ilog-cplex-optimization-studio-v1290). 
After installation, create an empty project and add the files mentioned below to its working directory (to solve the linear optimization problem). 

- MS Office Excel should be installed (to read and write .xlsx files).

- Python should be installed. Preferably install Anaconda (Anaconda | Individual Edition) and install Spyder Python IDE with python (to loop the ILOG Optimizer).
When using Anaconda or similar package managers, please note that any missing packages to run the code should be installed through the package manager. If you encounter such a problem, look up how install the specific package with the package manager and not the IDE. 


SECOND PART: FILES

Inside the python compiler directory (for example C:\Users\user_name\.spyder-py3\):

- loop_for_cplex.py

Inside the cplex compiler directory (for example C:\Users\user_name\opl\project_name\):

- data.xlsx (not provided)
- part.dat
- part.mod


THIRD PART: REMARKS 

You can inspect the python code with Microsoft Visual Studio Code or any other python IDE (such as Spyder).

Make sure to:

- Close Excel and ILOG Optimization Studio before running the code,

- Check that ALL pathing is correct – adjust each highlighted line in the code accordingly! (You can do this by using the search & replace function in any compiler or by pressing Ctrl + F and searching for “C:\\Users\\Reha\\opl\\thesis\\”. Replace each occurrence with “C:\\Users\\user_name\\opl\\project_name\\”.)
