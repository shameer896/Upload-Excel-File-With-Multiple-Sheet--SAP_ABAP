# Reading Multiple Sheets from Excel in SAP ABAP

The standard function module `ALSM_EXCEL_TO_INTERNAL_TABLE` in SAP ABAP reads only the active sheet of an Excel file. This guide explains how to extend this functionality to read multiple sheets from an Excel file by creating a custom function module.

## Solution Overview

This solution is based on the method described in [this blog post](https://www.linkedin.com/pulse/como-pasar-varias-hojas-de-un-archivo-excel-tablas-en-gaona-cruz). The approach involves creating a custom function module that processes each sheet in an Excel file and stores the data into internal tables.

## Steps to Implement the Solution

### 1. Create a Custom Structure

1. Go to Transaction SE11 to create a custom structure.
2. Name the structure (e.g., `ZALSMEX_TABLINE_CUSTOM`).
3. Define the components of the structure as shown below. Ensure the component names match exactly to avoid errors.

   ![Structure Components](https://github.com/user-attachments/assets/5d334a3c-10c1-4c42-b1a9-4471fccc5c18)

### 2. Create a Function Group and Copy the Function Module

1. Go to Transaction SE80.
2. Select **Function Group** in the dropdown.
3. Enter a name for your function group (e.g., `ZALSMEX_CUSTOM`) and save.
4. Copy the function module `ALSM_EXCEL_TO_INTERNAL_TABLE` from the standard function group to your new function group:
   - In SE80, select the function group containing `ALSM_EXCEL_TO_INTERNAL_TABLE` (e.g., `ALSMEX`).
   - Right-click on the function group name and select **Copy**.
   - Enter your new function group name (`ZALSMEX_CUSTOM`) and follow the prompts to copy the function module.

5. Update the imported parameters:
   - Go to the **Import** tab of your copied function module.
   - Add the component `SHEETS` as shown below.

6. Update table parameters:
   - Under the **Table** tab, change the table type to the structure you created (`ZALSMEX_TABLINE_CUSTOM`).

7. Modify the function module code:
   - Update the code in the function module `ZALSM_EX_IT_CUSTOM` as required. Refer to or copy the code from [this file](<A HREF=''>ZALSM_EX_IT_CUSTOM</A>).

8. Edit include files:
   - Go to the **Include Folder** of your function group.
   - Make changes in the following files:
     - `LZALSMEX_CUSTOMF01`
     - `LZALSMEX_CUSTOMTOP`
   - Refer to the provided files or copy the content as needed.

9. Activate everything:
   - Ensure that all changes are saved and activated without errors.

### 3. Create a Test Report

1. Write a test report to execute the newly created function module.
2. Use the sample Excel file provided for testing:
   - [Sample Excel File for Testing](<A HREF=''>THIS IS THE EXCEL FILE YOU CAN CHOOSE FOR TESTING</A>)

## Additional Resources

- **Reference Code for Testing**: [Test Report Code](<A HREF=''>LETS TEST1</A>)

## Acknowledgements

A huge thanks to the developer who shared the initial solution [here](https://www.linkedin.com/pulse/como-pasar-varias-hojas-de-un-archivo-excel-tablas-en-gaona-cruz). Without their contribution, this extended solution would not have been possible.

---

I hope this guide helps you to implement the solution effectively. If you have any questions or need further assistance, feel free to reach out!
