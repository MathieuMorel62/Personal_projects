# <p align=center>🔍 Data Matching Assistant 🔍</p>

**Data Matching Assistant** is a Python application designed for data reconciliation between two Excel files. It caters to data professionals and anyone faced with the task of comparing and harmonizing large datasets. By leveraging algorithms based on string similarity, this application facilitates the detection and pairing of similar records, ensuring data accuracy and quality.

## 📚 Table of Contents

- [Project Architecture](#project-architecture)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [License](#license)
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
