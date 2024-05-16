# Automating-Training-Evaluation-Notifications-and-Reporting-with-Excel-Script-and-Power-Automate

🚀 I'm thrilled to share a recent project where I successfully integrated an Excel Script with Power Automate to streamline and automate our training evaluation process. Here's a detailed overview of how this solution works:

## Table of Contents
- [Project Overview](#project-overview)
- [Script Details](#script-details)
  - [Data Extraction and Filtering](#data-extraction-and-filtering)
  - [Date Validation](#date-validation)
  - [Integration with Power Automate](#integration-with-power-automate)
  - [Automated Reporting](#automated-reporting)
- [Benefits](#benefits)
- [Code](#code)


## Project Overview
This project leverages Excel Script and Power Automate to automate the training evaluation process. It dynamically generates and sends personalized emails with pre-filled Microsoft Forms for trainees to provide feedback, and compiles the responses into a report for analysis.

## Script Details

### Data Extraction and Filtering
The script identifies a specific table in an Excel worksheet and extracts key columns including Email, Completion Date, Training Subject, Training Topic, Institute, Status, and Trainee Name.
It processes each row to filter out entries based on several conditions:
- The training must not be an "Induction Training."
- The training institute must not be "Udemy."
- The status of the training must be marked as "Completed."
- The trainee must have a valid email address.

### Date Validation
The script calculates if the training completion date plus an additional three days matches the current date.
If this condition is met, it collects the necessary details into an array.

### Integration with Power Automate
The script's output is used to dynamically generate and send emails to the trainees.
Each email contains a link to a Microsoft Forms evaluation form, which is pre-filled with details such as the training subject, topic, institute, and trainee name.
This ensures that each trainee receives a personalized evaluation form, making it easier for them to provide feedback.

![image](https://github.com/symeon158/Automating-Training-Evaluation-Notifications-and-Reporting-with-Excel-Script-and-Power-Automate/assets/106148298/4662f7c6-ec45-40da-960b-48fc3822e2e8)

![image](https://github.com/symeon158/Automating-Training-Evaluation-Notifications-and-Reporting-with-Excel-Script-and-Power-Automate/assets/106148298/ddd4ab04-87fe-4e59-a07b-703e9d42ba41)


### Automated Reporting
The data collected from the evaluation forms is automatically compiled into a report.
This report provides valuable insights into the training effectiveness and helps in making data-driven decisions for future training programs.

## Benefits
- **Efficiency:** The entire process is automated, saving significant time and reducing the potential for human error. ⏱️
- **Personalization:** Trainees receive personalized emails and forms, improving their experience and the likelihood of receiving comprehensive feedback. 💌
- **Data-Driven Insights:** Automated reporting helps in quickly analyzing the effectiveness of training programs and identifying areas for improvement. 📊

This project demonstrates the power of integrating Excel Scripts with Power Automate to create a seamless, efficient, and effective process for managing training evaluations. It’s a significant step towards leveraging automation to enhance workflow efficiency and data-driven decision-making.

## Code

```typescript
function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    let table = sheet.getTable("Table4"); // Adjust to your table's name
    if (!table) {
        console.log("Table not found.");
        return;
    }

    // Retrieve columns by their names
    let emailColumn = table.getColumnByName("Email");
    let completionDateColumn = table.getColumnByName("Completion Date");
    let trainingSubjectColumn = table.getColumnByName("Training Subject");
    let trainingTopicColumn = table.getColumnByName("Training Topic");
    let instituteColumn = table.getColumnByName("Institute");
    let statusColumn = table.getColumnByName("Status"); // New column for status
    let traineeColumn = table.getColumnByName("Trainee (Surname & Name)"); // Retrieve Trainee column
    let dataToSend: (string[])[] = []; // Array of arrays

    // Get range and values between header and total
    let range = table.getRangeBetweenHeaderAndTotal();
    let values = range.getValues();
    let completionDateIndex = completionDateColumn.getIndex();
    let emailIndex = emailColumn.getIndex();
    let trainingSubjectIndex = trainingSubjectColumn.getIndex();
    let trainingTopicIndex = trainingTopicColumn.getIndex();
    let instituteIndex = instituteColumn.getIndex();
    let statusIndex = statusColumn.getIndex(); // Get index for the "Status" column
    let traineeIndex = traineeColumn.getIndex(); // Get index for the "Trainee" column

    let today = new Date();
    today.setHours(0, 0, 0, 0);

    const errorValues = ["#value!", "#n/a", "#ref!", "#div/0!", "#num!", "#name?", "#null!"];

    // Loop through each row to check conditions
    for (let i = 0; i < values.length; i++) {
        let row = values[i];
        let completionDateValue = row[completionDateIndex];
        let trainingSubjectValue = row[trainingSubjectIndex];
        let trainingTopicValue = row[trainingTopicIndex];
        let instituteValue = row[instituteIndex];
        let statusValue = row[statusIndex]; // Status value for the current row
        let traineeValue = row[traineeIndex]; // Trainee value for the current row
        let emailAddress = row[emailIndex];

        // Check if status is 'completed' and other conditions
        if (typeof completionDateValue === 'number' &&
            trainingSubjectValue !== "Induction Training" &&
            instituteValue !== "Udemy" &&
            statusValue === "Completed" &&
            typeof emailAddress === "string" &&
            emailAddress.trim() !== "" &&
            !errorValues.includes(emailAddress.trim().toLowerCase())) {

            let completionDate = convertExcelDateToJSDate(completionDateValue);
            completionDate.setHours(0, 0, 0, 0);
            let completionDatePlusThree = new Date(completionDate.getTime() + (3 * 24 * 60 * 60 * 1000)); // Adding 3 days to completion date

            // Compare the adjusted completion date with today
            if (completionDatePlusThree.getTime() === today.getTime()) {
                // Include trainee name in the data to send
                dataToSend.push([emailAddress, trainingSubjectValue, trainingTopicValue, instituteValue, traineeValue]);
            }
        }
    }
    console.log(dataToSend);
    // Return the `dataToSend` array to Power Automate
    return { dataToSend: dataToSend };
}

// Helper function to convert Excel serial date to JS Date object
function convertExcelDateToJSDate(serialDate: number): Date {
    const excelStartDate = new Date(Date.UTC(1899, 11, 30));
    const msPerDay = 24 * 60 * 60 * 1000;
    return new Date(excelStartDate.getTime() + serialDate * msPerDay);
}
