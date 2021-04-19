import { Component, OnInit } from "@angular/core";
import * as XLSX from "@sheet/chartdemo";

type AOA = any[][];
@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.scss"],
})
export class AppComponent implements OnInit {
  title = "chartdemo";

  ngOnInit() {}

  data: AOA = [
    [1, 2],
    [3, 4],
  ];
  wopts: XLSX.WritingOptions = { bookType: "xlsx", type: "array" };
  fileName: string = "SheetJS.xlsx";

  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>evt.target;
    if (target.files.length !== 1) throw new Error("Cannot use multiple files");
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: "binary" });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = <AOA>XLSX.utils.sheet_to_json(ws, { header: 1 });
    };
    reader.readAsBinaryString(target.files[0]);
  }

  export(): void {
    /* generate worksheet */

    var data = [
      /*  A  B  C  D -- each series has its own X and Y values*/
      [1, 1, 0, 1],
      [2, 3, 3, 2],
      [3, 5, 8, 3],
      [4, 7, 15, 4],
      [5, 9, 24, 5],
    ];
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(data);
    var cs = XLSX.utils.aoa_to_sheet(data);

    cs["!type"] = "chart"; // mark as chartsheet
    cs["!title"] = "My Scatter"; //chart title
    cs["!legend"] = { pos: "b" }; // legend
    cs["!plot"] = []; // build up plot
    cs["!pos"] = { x: 0, y: 0, w: 400, h: 300 }; // location

    var scatter = {
      t: "scatter",
      ser: [],
    };

    /* X values in column A, Y values in column B, link to data worksheet */
    scatter.ser.push({
      name: "Data+Ref",
      /* column A is X, B is Y (raw is not specified) */
      cols: ["xVal", "yVal"],
      ranges: ["Data!A1:A5", "Data!B1:B5"],
      marker: {
        symbol: "triangle",
      },
    });

    /* X values in column D, Y values in column C, raw data */
    scatter.ser.push({
      name: "Data",
      /* column C is Y, D is X */
      cols: ["yVal", "xVal"],
      marker: null,
      linecolor: { rgb: "00FFFF" },
      labels: true,
      line: true,
      smooth: false,
    });

    /* third series, skip cache */
    scatter.ser.push({
      name: "Ref",
      cols: ["xVal", "yVal"],
      ranges: ["Data!A1:A5", "Data!C1:C5"],
      raw: true, // do not populate from cache
      marker: {
        symbol: "x",
        color: { rgb: "FF00FF" },
      },
      /* multiple trendlines can be attached to a given series */
      trendlines: [
        {
          t: "linear",
          eq: true /* show equation */,
          r2: true /* show R2 value */,
        },
        {
          t: "quadratic",
          color: { rgb: "00FF00" },
        },
      ],
    });

    cs["!plot"].push(scatter);
    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, cs, "Chart");
    XLSX.utils.book_append_sheet(wb, ws, "Data");
    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }
}
