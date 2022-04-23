
# Milan Lawyers data retrieval

This project is created to retrieve Milan Lawyers public information from a specific site with the only use of their name.


## Tech Stack

**Programming Language** Python

**Libraries** Selenium, Pandas, Openpyxl


## Usage/Examples

This program was created to obtain information from lawyers in Milan, such as:
-   Name
-	Studio Street
-	City
-	Province
-	ZIP code
-	PEC address
-	National ID

The following elements are required to run this program:
-	A version of Python 3 (the program was created using Python 3.10.4)
-	Selenium module required for internet browsing, searching, and copying information
-	Pandas module required to allow the program to read excel files
-	Openpyxl module required to allow the program to create and fill the results excel file
-	The correct Microsoft Edge WebDriver version for your build of Microsoft Edge (the driver for the Microsoft Edge version 100.0.1185.50 is already in the project)
-	The "Lawyer's name" file filled with a list of First and Last names of the lawyers that we want to retrieve the information(an example file can be inserted in the project that can be used).

The program starts by opening the excel file "Lawyers name" from which it extracts the name and surnames of lawyers.
It then uses the Microsoft Edge driver to access Albo Sfera in Milan. Enter the name and surname of the Lawyer in the corresponding field and press the search button.

![Captură 1](https://user-images.githubusercontent.com/46033329/164895742-11d3a2ac-a488-4cf0-8ed3-2fb9d80cd93a.JPG)
 

The program is set to fall asleep for a few seconds, time required by the website to find the lawyer.

![Captură 2](https://user-images.githubusercontent.com/46033329/164895758-27a2cfc7-75d7-4b0e-b06d-3bd1c0d2d28a.JPG))
  
Once the lawyer is found, the program press the button to get more details and then copy all the information needed to complete the Excel file with the results.

![Captură 3](https://user-images.githubusercontent.com/46033329/164895759-d4113c56-7944-4bd2-96f2-46446d7db93b.JPG)
  
![Captură 4](https://user-images.githubusercontent.com/46033329/164895760-b14f035d-bb2b-404e-8c8d-69e8000e3a61.JPG)
  

Finally, the program closes the edge pages to avoid overloading the computer. 
The program executes the respective action for each Lawyer from the list with names and surnames.

## Run Locally

Clone the project

```bash
  git clone https://github.com/OMarius97/Milan_Lawyers_data_retrival
```



Install dependencies

```bash
  pip install selenium
  pip install pandas
  pip install openpyxl
```

Go to the project directory

```bash
  cd Milan_Lawyers_data_retrival
```

Run the program

```bash
  Run the program main.py
```


## Authors

- [@Marius](https://www.github.com/OMarius97)


## Feedback

If you have any feedback, please reach out to me at mariusoana1997@gmail.com

