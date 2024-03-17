app.get("/excel", async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    // sheet name
    const worksheet = workbook.addWorksheet(`${client_id} P&L (892) 2023-24`);

    // Add company name on the first row
    const row1 = worksheet.getRow(1);
    row1.getCell(1).value = "INDIRA SECURITES PVT. LTD.(NSE)";

    // Add report details on the second row
    const row2 = worksheet.getRow(2);
    row2.getCell(1).value = "892:Annual P&L";

    // Add client details on the third row
    const row3 = worksheet.getRow(3);
    row3.getCell(
      1
    ).value = `Client :${client_id} -${client_name} Year:${year} To Date :${to_date}`;

    // headers
    const row4 = worksheet.getRow(4);
    headers_data.forEach((ele, index) => {
      row4.getCell(index + 1).value = ele;
      row4.getCell(index + 1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "29A5DE" },
      };
    });

    // merging cells c:string r:number
    worksheet.mergeCells("A1:F1");
    worksheet.mergeCells("A2:F2");
    worksheet.mergeCells("A3:F3");

    // Alignment: To align the text within a cell.
    row1.getCell(1).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    row2.getCell(1).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    row3.getCell(1).alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    // Initialize row count
    let rowCount = 4; // Rows 1 to 4 have been written so far

    // Add data from dummy_data
    dummy_data.forEach((data) => {
      const row = worksheet.addRow(data);
      rowCount++;
    });

    // Increment rowCount for the new data rows
    rowCount += dummy_data.length;

    // Auto-fit column widths
    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const length = cell.value ? cell.value.toString().length : 10;
        if (length > maxLength) {
          maxLength = length;
        }
      });
      column.width = maxLength < 10 ? 10 : maxLength;
    });

    // Generate Excel file
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="employees.xlsx"' // you can change sheet name from here
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.log(error);
    res.status(500).json({
      success: false,
      message: error?.message || "Internal Server Error",
      error: error,
    });
  }
});
