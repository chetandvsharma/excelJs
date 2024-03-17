import express from "express";
import morgan from "morgan";
import ExcelJS from "exceljs";
import {
  client_id,
  client_name,
  to_date,
  headers_data,
  year,
  dummy_data,
} from "./dummyData.js";

const app = express();
app.use(morgan("dev"));

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

    // Add client details on the thired row
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

    // Add headers for employee details

    // Example employee details (you can replace this with your actual data)
    // const employees = [
    //   ["John Doe", "Manager", 5],
    //   ["Jane Smith", "Developer", 3],
    //   ["Alice Johnson", "HR", 7],
    // ];

    // Add employee details to the Excel file
    // employees.forEach((employee) => {
    //   worksheet.addRow(employee);
    // });

    // Apply formatting to numeric values
    // for (let i = 6; i <= employees.length + 1; i++) {
    //   worksheet.getCell(`C${i}`).numFmt = '0 "years"';
    // }

    // Initialize row count
    let rowCount = 4; // Rows 1 to 4 have been written so far

    // data listing gonna start from 5th row
    // const row5 = worksheet.getRow(5);
    // row5.getCell(1).value = "OP_ASSETS";
    // worksheet.mergeCells("B5:Q5");

    //  for (let i = 0; i < dummy_data.length; i+=1 ) {

    //  }

    Object.entries(dummy_data).forEach(([key, value], index) => {
      const row = worksheet.getRow(rowCount + 1);
      row.getCell(1).value = key;
      row.getCell(1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {argb: "DCECF4 "}
      }
      worksheet.mergeCells(`A${rowCount + 1}:Q${rowCount + 1}`);
      ++rowCount; // row added for heading of data like: OP_ASSETS
      console.log('heading row >> ', rowCount, key)
      for (let i = 0; i < value.length; i += 1) {
        const scripRow = worksheet.getRow(rowCount + 1);
        scripRow.getCell(1).value = `Scrip :${value[i].scrip_code}-${value[
          i
        ].scrip_name.toLocaleUpperCase()}`;
        worksheet.mergeCells(`A${rowCount + 1}:Q${rowCount + 1}`);
        scripRow.getCell(1).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "DAF7A6" }
        };

        console.log("scrip row ", rowCount + 1, value[i].scrip_name);
        ++rowCount; // new row for each entry of each scrip name
        console.log("row >> ", rowCount);
        Object.entries(value[i]).forEach(
          ([childKey, childValue], childIndex) => {
            const rowForValue = worksheet.getRow(rowCount + 1);
            rowForValue.getCell(childIndex + 1).value = childValue;
            // rowForValue.getCell(`${String.fromCharCode(65 + childIndex)}${rowCount + 1}`).value = childValue;
          }
        );
        ++rowCount; // for statoring each data of one heading
      }
      // calculating profit and loss
      const calculatingRow = worksheet.getRow(rowCount + 1);
      calculatingRow.getCell(1).value = 'Profit/Loss';
      calculatingRow.getCell(1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {argb: "DCECF4 "}
      }
      ++rowCount;
    });

    // Auto-fit column widths
    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const length = cell.value ? cell.value.toString().length : 6;
        if (length > maxLength) {
          maxLength = length;
        }
      });
      column.width = maxLength < 6 ? 6 : maxLength;
    });

    // Generate Excel file
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      'attachment; filename="employees.xls"' // you can change sheet name from here
    );

    await workbook.xlsx.write(res);
    res.end();
    // res.status(200).send('excel downloaded')
  } catch (error) {
    console.log(error);
    res.status(500).json({
      success: false,
      message: error?.message || "Internal Server Error",
      error: error,
    });
  }
});

app.listen(6999, () => console.log("server running"));
