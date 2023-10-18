const express = require('express');
const app = express();
const excel = require('exceljs');
const fs = require('fs');
const path = require('path');

// Use a data structure to store form submissions
const formSubmissions = [];

// Middleware to parse JSON and form data
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve your HTML file
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Handle form submission
app.post('/submit', (req, res) => {
  const { name, branch, section, q2, q3, q4, q5, q6, q7, q8, q9, q10, q11, q12 } = req.body;

  // Calculate the score here
  // Store the form data in your data structure
  const submissionData = {
    name,
    branch,
    section,
    q2,
    q3,
    q4,
    q5,
    q6,
    q7,
    q8,
    q9,
    q10,
    q11,
    q12,
  };
  formSubmissions.push(submissionData);

  res.send('Form submitted successfully.');
});

// Provide the admin interface to view form submissions
// Provide the admin interface to view form submissions
app.get('/admin', (req, res) => {
    // Create a workbook or open the existing Excel file if it exists
    const excelFilePath = 'all_submissions.xlsx';
    const workbook = new excel.Workbook();
  
    // Check if the file exists
    if (fs.existsSync(excelFilePath)) {
      workbook.xlsx.readFile(excelFilePath)
        .then(() => {
          const worksheet = workbook.getWorksheet(1); // Assuming it's the first worksheet
  
          // If it's a new session, write the headers
          if (worksheet.actualRowCount === 1) {
            worksheet.addRow(['Name', 'Branch', 'Section', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12']);
          }
  
          // Add all form submissions to the Excel file
          formSubmissions.forEach((submission) => {
            worksheet.addRow(Object.values(submission));
          });
  
          // Generate the Excel file
          return workbook.xlsx.writeFile(excelFilePath);
        })
        .then(() => {
          // Send the updated file for download
          res.download(excelFilePath, 'all_submissions.xlsx', (err) => {
            if (err) {
              console.error('Error sending Excel file:', err);
            }
          });
        })
        .catch((err) => {
          console.error('Error reading or generating Excel file:', err);
          res.status(500).send('Error reading or generating Excel file.');
        });
    } else {
      // If the file doesn't exist, handle this case as needed
      res.send('No form submissions yet.');
    }
  });
  
// Start the server
const port = 3000;
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
