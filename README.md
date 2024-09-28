# Word Document Automation Project

This project automates the creation of Word (.docx) documents using predefined templates. The system reads data from an Excel file, processes it, and generates personalized documents for each student, inserting their name and grades for subjects like Math, Physics, and Chemistry.

## Description

The script automates the generation of Word documents for students using a `.docx` template and data stored in an `.xlsx` file. Documents are generated individually for each student, containing their name and grades in different subjects. The project uses the `docxtpl` library to manage Word templates and `pandas` to read Excel data.

## Project Files

- **`main.py`**: The main script that manages the automation process.
- **`template.docx`**: The Word template where each student's data is inserted.
- **`students.xlsx`**: The Excel file that contains student data and grades.
- **`output/`**: Folder where the generated files are saved, one Word file per student.

## Project Structure


```bash
├── input/
│   ├── template.docx         # Word template
│   ├── students.xlsx         # Student data
├── output/                   # Folder where generated files are saved
├── main.py                   # Main script
├── README.md                 # This file
└── requirements.txt          # Required libraries
```

## Usage

1. Ensure that the Excel file (`students.xlsx`) contains the following fields:
   - **Student_name**: Student's name.
   - **Math**: Grade for Math.
   - **Physics**: Grade for Physics.
   - **Chemistry**: Grade for Chemistry.

2. The Word file (`template.docx`) must have the appropriate placeholders for the data to be inserted, such as: `{{ student_name }}`, `{{ math_note }}`, `{{ physics_note }}`, `{{ chemistry_note }}`, `{{ name }}`, `{{ phone }}`, `{{ email }}`, `{{ date }}`.

3. After running the script, Word documents for each student will be saved in the `output/` folder, with the name format `student_<first_name>.docx`.

## Required Libraries

To run this project, you need to install the following Python libraries:

- [pandas](https://pandas.pydata.org/) - To handle Excel data manipulation.
- [docxtpl](https://docxtpl.readthedocs.io/en/latest/) - To work with Word document templates.
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) - To interact with Excel files (installed automatically with pandas).

You can install these by running the following command:

```bash
pip install pandas docxtpl openpyxl
```