# Waste Management System using Python OOP

A terminal-based **Waste Management System** developed using **Python Object-Oriented Programming** and **Excel file storage**.
This project was completed in **December 2024** as part of the **CSE222: Object Oriented Programming II Lab** course.

The system helps manage waste categories, collection schedules, and waste collectors through a simple command-line interface. It uses Excel files for storing and retrieving records.

---

## Project Overview

Waste management is an important issue for modern communities, especially in urban areas where inefficient collection and disposal can create environmental and public health problems.

This project provides a lightweight and low-cost solution for small-scale waste management operations. It allows users to add waste types, schedule waste collection, assign collectors, view collection schedules, and delete records. The system is designed using Python OOP principles and uses the `openpyxl` library to store data in Excel files.

---

## Project Purpose

The main purpose of this project was to apply Python Object-Oriented Programming concepts to solve a real-world problem. The project demonstrates how classes, objects, lists, file handling, and external libraries can be used to build a structured terminal-based management system.

This project was developed to practice:

* Python fundamentals
* Object-Oriented Programming
* Class-based system design
* File-based data persistence
* Excel file handling using `openpyxl`
* Terminal-based user interaction
* Real-world problem-solving using programming

---

## Features

* Add waste type
* Schedule waste collection
* Assign collector to a collection schedule
* View collection schedules
* Delete waste record
* Delete schedule record
* Delete collector record
* Store waste data in Excel
* Store schedule data in Excel
* Store collector data in Excel
* Terminal-based menu system
* Object-oriented project structure

---

## Technologies Used

| Category             | Technology                                 |
| -------------------- | ------------------------------------------ |
| Programming Language | Python                                     |
| Library              | openpyxl                                   |
| Data Storage         | Excel files                                |
| Interface            | Terminal / Command Line                    |
| IDE                  | Visual Studio Code                         |
| Course               | CSE222: Object Oriented Programming II Lab |

---

## Course Information

| Item            | Details                            |
| --------------- | ---------------------------------- |
| Course Code     | CSE222                             |
| Course Title    | Object Oriented Programming II Lab |
| Project Type    | Mini Lab Project                   |
| Submission Date | 31 December 2024                   |
| Student         | Rashedul Hasan Shohan              |
| Institution     | Daffodil International University  |

---

## Project Structure

```text id="vhhqt1"
waste-management-system-python/
│
├── README.md
├── LICENSE
├── .gitignore
├── requirements.txt
│
├── src/
│   └── waste_management_system.py

```

---

## Main Classes

### `Waste`

Stores waste information such as:

* Waste ID
* Waste type
* Description

### `CollectionSchedule`

Stores collection schedule information such as:

* Schedule ID
* Waste type
* Collection date
* Location

### `WasteCollector`

Stores collector information such as:

* Collector ID
* Collector name
* Assigned routes

### `WasteManagementSystem`

Main controller class that manages:

* Waste records
* Collection schedules
* Collectors
* Excel file loading and saving
* Menu-based user interaction

---

## Installation and Setup

### 1. Clone the repository

```bash id="x851w1"
git clone https://github.com/itsrashedhasan/waste-management-system.git
cd waste-management-system-python
```

### 2. Create a virtual environment

For Windows PowerShell:

```powershell id="649q56"
python -m venv venv
.\venv\Scripts\activate
```

For Linux/macOS:

```bash id="kmecya"
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies

```bash id="fd679g"
pip install -r requirements.txt
```

### 4. Run the project

For Windows PowerShell:

```powershell id="aidlyj"
python .\src\waste_management_system.py
```

For Linux/macOS:

```bash id="7wf23i"
python src/waste_management_system.py
```

---

## How the System Works

When the program starts, it displays a terminal menu:

```text id="szs6dd"
Welcome to the Waste Management System!
1. Add Waste Type
2. Schedule Collection
3. Assign Collector
4. View Collection Schedules
5. Delete Waste
6. Delete Schedule
7. Delete Collector
8. Exit
```

The user can choose an option and perform the required task. Data is automatically saved into Excel files.

Generated Excel files:

```text id="zxlqpe"
wastes.xlsx
schedules.xlsx
collectors.xlsx
```

---

## Example Workflow

```text id="8z8edh"
1. Add a waste type
2. Enter waste ID, type, and description
3. Schedule a collection using the waste ID
4. Enter collection date and location
5. Assign a collector to the schedule
6. View all collection schedules
7. Exit the system
```

---

## Learning Outcomes

Through this project, I learned:

* How to design a class-based Python application
* How to apply Object-Oriented Programming principles
* How to use constructors, methods, and object relationships
* How to store and retrieve data using Excel files
* How to use the `openpyxl` library
* How to build a terminal-based menu system
* How to manage project documentation
* How to solve a practical problem using Python

---

## Limitations

* The system is terminal-based and has no graphical user interface
* Excel files are used instead of a full database
* Large-scale data handling may be limited
* Collector route assignments are not fully restored after restarting the program
* No real-time GPS tracking
* No web or mobile interface
* No advanced analytics or prediction system

---

## Future Improvements

* Add a graphical user interface
* Replace Excel storage with SQLite or MySQL
* Add login and user role management
* Add real-time GPS-based route tracking
* Add route optimization
* Add dashboard and reports
* Add web or mobile application support
* Add data visualization
* Add AI-based waste generation prediction

---

## License

This project is licensed under the MIT License.

---

## Author

**Rashedul Hasan Shohan**
BSc in Computer Science and Engineering
Daffodil International University
