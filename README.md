**ACADEMIC PERFORMANCE ANALYTICS AND AUTOMATED SCORE SCALING SYSTEMS**

INTRODUCTION

A full-stack web application that automates academic marks processing, statistical analysis, and dynamic score normalization. The system reads Excel-based student datasets, computes performance metrics, visualizes grade distribution, and generates structured downloadable reports with embedded charts.

FEATURES

Upload Excel file containing student marks

Automatic calculation of:
-maximum marks
-minimum marks
-mean
-standard deviation
-minimum passing marks (Grade 'DD' threshold)
Dynamic score scaling (Topper normalized to 100)
Sorted results (highest to lowest marks)
Interactive grade distribution bar chart
Downloadable Excel report including:
-processed data
-statistical summary table
-embedded chart visualization

TECH STACK
Python
Flask
Pandas
NumPy
OpenPyXL
HTML5
CSS3
Tailwind CSS
JavaScript
Chart.js

PROJECT STRUCTURE
marks_dashboard/
│
├── app.py
├── templates/
│   └── index.html
├── static/
│   └── uploads/
└── processed/

INSTALLATION AND SETUP
1. Clone the repository
2. Create a virtual environment
3. Install dependencies:
   pip install flask pandas numpy openpyxl
4. Run the application:
   python app.py
5. Open browser at:
   http://127.0.0.1:5000

OUTPUT
Interactive analytics dashboard
Downloadable processed Excel File
Embedded grade distribution chart
Automated statistical summary

