GenerateReport() is the main subroutine that generates the report.
It assumes that you have two worksheets named "DataSheet" containing your data and "ReportTemplate" containing your report template.
It copies the template from "ReportTemplate" and pastes it into a new worksheet.
Then, it copies the data from "DataSheet" and pastes it into the new worksheet below the template.
Finally, it saves the generated report with a timestamp.
Error handling is implemented using On Error GoTo statements and ErrorHandler label to  handle runtime errors.
Password protection is added when saving the workbook using the Password parameter of the SaveAs method. we can  replace "YourPasswordHere" with the desired password our choice.
