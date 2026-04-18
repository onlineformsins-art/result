const XLSX = require('xlsx');

// Create a new workbook
const wb = XLSX.utils.book_new();

// Sample Data
const data = [
  {
    "Index Number": "1001",
    "Name": "Alice Smith",
    "Mathematics": 85,
    "Science": 92,
    "English": 78,
    "History": 88,
    "Geography": 90,
    "Total": 433,
    "Average": 86.6,
    "Grade": "A"
  },
  {
    "Index Number": "1002",
    "Name": "Bob Johnson",
    "Mathematics": 65,
    "Science": 70,
    "English": 60,
    "History": 55,
    "Geography": 62,
    "Total": 312,
    "Average": 62.4,
    "Grade": "C"
  },
  {
    "Index Number": "1003",
    "Name": "Charlie Brown",
    "Mathematics": 95,
    "Science": 88,
    "English": 92,
    "History": 85,
    "Geography": 94,
    "Total": 454,
    "Average": 90.8,
    "Grade": "A+"
  }
];

// Convert data to worksheet
const ws = XLSX.utils.json_to_sheet(data);

// Add worksheet to workbook
XLSX.utils.book_append_sheet(wb, ws, "Results");

// Write to file
XLSX.writeFile(wb, 'results.xlsx');
console.log("results.xlsx created successfully!");
