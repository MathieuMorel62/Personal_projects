# <p align=center>🔍 Data Matching Assistant 🔍</p>

**Data Matching Assistant** is a Python application designed for reconciling data between two sheets of an Excel file. It caters to data professionals and anyone faced with the task of comparing and harmonizing large datasets. Using string similarity-based algorithms, this application facilitates the detection and matching of similar records, ensuring data accuracy and quality.

## 📖 Description and Features

Data Matching Assistant is designed to ease data reconciliation between two sheets of an Excel file. Here's what you can do with this program:

- **File Selection**: Easily import an Excel file for processing. Ensure the file is closed to avoid any conflicts.
- **Sheet and Column Selection**: Choose two specific sheets to compare and select corresponding columns for matching.
- **Matching Configuration**: Set a matching threshold to determine how similar two records should be to be considered a match.
- **Results Visualization**: Once matching is done, results are displayed in the interface, showing matching records and the match percentage.
- **Results Export**: You can opt to save the results to a new Excel file for further analysis.

## File Requirements

- **Format**: The file should be in .xlsx format.
- **Structure**: The file should contain at least two sheets with tabular data. Columns you wish to compare should have headers.
- **Size**: While the application can handle large files, for optimal performance, it's recommended not to exceed 10,000 rows per sheet.
- **Content**: Columns selected for matching should contain text. Numeric values, dates, or other data types are not recommended for string similarity-based matching.

## 📚 Table of Contents

- [Project Architecture](#project-architecture)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Support and Contribution](#support-and-contribution)
- [Main Contributor](#main-contributor)

## 🏗️ Project Architecture

The **Data Matching Assistant** application is modularly and intuitively structured:

- **Main Class - `DataMatcherApp`**: This is the heart of the application. It handles all user interactions, file imports, and the implementation of matching algorithms.
  - **User Interface**: Thanks to Tkinter, a user-friendly graphical interface is provided. It simplifies the selection of files, sheets, and columns while offering a clear visualization of the results.
  - **Data Management with Pandas**: `pandas` is used to import, process, and handle data from Excel files. This library provides great flexibility and ensures optimal tabular data manipulation.
  - **Matching Algorithms with Fuzzywuzzy**: The `fuzzywuzzy` library is used to perform matches based on string similarity. It's crucial for identifying records that might not be identical but are similar enough to be considered matches.

## 🛠️ Prerequisites

- Python 3.x
- pip3
- git

## 📥 Installation

1. **Clone the repository**: 

```bash
git clone https://github.com/MathieuMorel62/Personal_projects.git
```


2. **Navigate to the project directory**:

```bash
cd DataMatchingAssistant
```


3. **Install the dependencies**:

```bash
pip3 install -r requirements.txt
```


4. **Launch the application**:

```bash
python3 data_assistant.py
```

## 🖥️ Usage

<img width="100%" alt="Capture d’écran 2023-09-16 à 01 51 15" src="https://github.com/MathieuMorel62/Personal_projects/assets/113856302/939b07e5-a729-475d-8ba5-23184d75a3d7">

1. **Start the application**: 
Upon launching the application, you'll be presented with a graphical user interface. Ensure any Excel files you wish to use are closed to avoid conflicts.

2. **Select Excel files**: 
Click the relevant buttons to load your source and target files. These will be used for the matching process.

3. **Specify sheets and columns**: 
After loading the files, select the specific sheets you wish to compare from the dropdown menus. Next, select the corresponding columns for matching.

4. **Initiate the matching process**: 
After selecting sheets and columns, click the matching button. The application will then start the reconciliation process and display results upon completion.

5. **View results**: 
Once matching is completed, the results will be displayed in the interface. You can review the found matches and optionally export the results for further analysis.

## 🤝 Support and Contribution

If you have suggestions or bugs to report, please open a ticket. Contributions are also welcome; feel free to propose pull requests.

## 🚀 Main Contributor

- Mathieu Morel - [LinkedIn Profile](https://www.linkedin.com/in/mathieu-morel-9ab457261/)
