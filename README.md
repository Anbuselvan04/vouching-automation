# Exercise-14 Automate Vouching Process to Compare PDF and Excel Data

### Reg No : 212221040013

## AIM: 
To automate the vouching process by extracting data from PDF invoices, comparing it with Excel data, and marking entries as "matched" or "mismatch" based on the comparison.

## Activities Required:
1. Excel Application Scope
2. Read Range
3. For Each Row in Data Table
4. Assign
5. Read PDF with OCR
6. Write Text File
7. Multiple Assign
8. If
9. Log Message

## Procedure:
1. Open UiPath Studio and create a new project called `VouchingAutomation`.

2. Inside the **Main.xaml** file, use an **Excel Application Scope** activity to open the Excel file containing invoice data.

3. Add a **Read Range** activity within the Excel Application Scope to read the data into a DataTable variable called `sampleData`.

4. Add a **For Each Row in Data Table** activity to iterate through each row in `sampleData`.

5. Inside the loop, add a **Sequence** to perform the following actions for each invoice:
   
   - Use **Assign** activities to define `filename`, `textFilePath`, and `pdfText` variables, where:
     - `filename` holds the current row’s invoice number from Excel.
     - `textFilePath` specifies the path to save the text extracted from the PDF.
   
   - Use **Read PDF with OCR** to read the PDF file specified by `filename`, storing the extracted text in `pdfText`.
   
   - Use a **Write Text File** activity to save the `pdfText` content to a text file at `textFilePath`.

6. Add **Multiple Assign** activities to extract `invoiceDate`, `invoiceAmt`, `invoiceNo`, and set a boolean variable `isMatch` to `True`.

7. Add **If** conditions to verify if the extracted PDF data matches the corresponding fields in the Excel row:
   
   - For each field (`invoiceNo`, `invoiceDate`, and `invoiceAmt`), check if `pdfText` contains the required value. If not, set `isMatch` to `False` and log the missing field.

8. After checking each field, use **Multiple Assign** to update the Excel row:
   - Set the "Status" column to `isMatch`.
   - If there's a mismatch, log the missing fields in the "Remarks" column.

9. Finally, add a **Log Message** activity to display a confirmation message when the vouching process for each PDF file is complete.

10. Save and run the workflow.

## Workflow:
![image](https://github.com/user-attachments/assets/53e38d8e-7c28-4116-a5b0-3784d55d5820)

## Output:
![image](https://github.com/user-attachments/assets/f3dd4619-2f32-4af6-8cfe-a6f624b91163)

## Result:
Thus, the vouching automation process that compares PDF invoice data with Excel data has been successfully implemented.
