import { Component, OnInit, ViewEncapsulation } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as fs from 'file-saver';
import { MessageService } from 'primeng/api';
import { Workbook } from 'exceljs';
import {dataTable} from './../Data'
@Component({
  selector: 'app-excel-comp',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './excel-comp.component.html',
  styleUrls: ['./excel-comp.component.css'],
  providers: [MessageService],
  encapsulation: ViewEncapsulation.None
})
export class ExcelCompComponent implements OnInit {
  DataList: any = [];
  Thinborder: Object = {
    top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
  }
  SideBoldTopborder: Object = {
    top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'thin' },
      right: { style: 'medium' },
  }
  SideBoldborder: Object = {
    top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'medium' },
  }
  font11SiZe: Object = {
      size: 11,
      bold: true,
      color: { argb: '4775d8' },
      name :  'Calibri' ,
  }
  constructor(){}
  ngOnInit() {
   this.DataList = dataTable
 console.log('Help the ',this.DataList)
 }
  DownloadFile() {
   let workbook = new Workbook();
    let worksheetInflowThree = workbook.addWorksheet('Projected Vs Actual');
    let Project1STATEMENTRow = worksheetInflowThree.addRow([]);
    Project1STATEMENTRow.getCell(1).value = "Projected Vs Actual";
    Project1STATEMENTRow.getCell(1).font=this.font11SiZe
    Project1STATEMENTRow.getCell(1).alignment = {
      horizontal:'center'
    }
    worksheetInflowThree.mergeCells('A1', 'C1');

    let Project1PERIODRow = worksheetInflowThree.addRow([]);
     Project1PERIODRow.getCell(1).value = "Project: " + this.DataList[0].Project;
     Project1PERIODRow.getCell(1).font= this.font11SiZe
    Project1PERIODRow.getCell(1).alignment = {
      horizontal:'left'
    }
    worksheetInflowThree.mergeCells('A2', 'C2');

    let Project1NameRow = worksheetInflowThree.addRow([]);
     Project1NameRow.getCell(1).value = "Financial Year :  "+ this.DataList[0].Fin;
     Project1NameRow.getCell(1).font={
       size: 11,
       bold: true,
       name :  'Calibri' ,
    }
    Project1NameRow.getCell(1).alignment = {
      horizontal: 'left',
      vertical : "middle"
    }
    worksheetInflowThree.getRow(3).height = 19.50;
    worksheetInflowThree.mergeCells('A3', 'C3'); 

    let InflowRow = worksheetInflowThree.addRow([]);
    InflowRow.getCell(1).value = "Inflow";
    InflowRow.getCell(1).font={
       size: 11,
       bold: true,
       name :  'Calibri' ,
    }
    InflowRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'c6e0b4' },
      bgColor: { argb: '' },
    }
    InflowRow.getCell(1).border = this.SideBoldTopborder
    InflowRow.getCell(1).alignment = {
      horizontal:'center'
    }
    worksheetInflowThree.mergeCells('A4', 'C4');

    let SiteName1Row = worksheetInflowThree.addRow([]);
    SiteName1Row.getCell(1).value = "Site Name";
    SiteName1Row.getCell(1).font={
       size: 10,
       bold: true,
       name :  'Calibri' ,
    }
    SiteName1Row.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'bdd7ee' },
      bgColor: { argb: '' }
    }
    SiteName1Row.getCell(1).border = this.Thinborder
     
    SiteName1Row.getCell(1).alignment = {
      horizontal:'center'
    }

    SiteName1Row.getCell(2).value = "Projected Inflow";
    SiteName1Row.getCell(2).font={
       size: 10,
       bold: true,
       name :  'Calibri' ,
    }
    SiteName1Row.getCell(2).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'bdd7ee' },
      bgColor: { argb: '' }
    }
    SiteName1Row.getCell(2).border = this.Thinborder
    SiteName1Row.getCell(2).alignment = {
      horizontal:'center'
    }

    SiteName1Row.getCell(3).value = "Actual Inflow";
    SiteName1Row.getCell(3).font={
       size: 10,
       bold: true,
       name :  'Calibri' ,
    }
    SiteName1Row.getCell(3).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'bdd7ee' },
      bgColor: { argb: '' }
    }
    SiteName1Row.getCell(3).border = this.SideBoldborder
    SiteName1Row.getCell(3).alignment = {
      horizontal:'center'
    }
    worksheetInflowThree.getColumn(1).width = 45.71;
    worksheetInflowThree.getColumn(2).width = 23.14;
    worksheetInflowThree.getColumn(3).width = 23.14;
    worksheetInflowThree.getRow(5).height = 12.75;

        const data8:any = [];
        this.DataList[0].inflow.forEach((ele:any) => {
          data8.push(Object.values(ele))
        });
    data8.forEach((d:any )=> {
      const row = worksheetInflowThree.addRow(d);
      for (let i = 0; i < d.length; i++) {
        row.getCell(i + 1).border = this.Thinborder
        row.getCell(3).border = this.SideBoldborder
        row.getCell(i + 1).alignment = {
          horizontal: 'left',
          vertical: 'middle',
          wrapText: true
        }
        row.getCell(i + 1).font = {
          size: 10
        }
      }
      row.getCell(2).alignment = {
        horizontal: 'right',
        vertical: 'middle',
        wrapText: true
      }
      row.getCell(3).alignment = {
        horizontal: 'right',
        vertical: 'middle',
        wrapText: true
      }
    }
        );

     let TotalInflowAmountBillRow = worksheetInflowThree.addRow([])
      TotalInflowAmountBillRow.getCell(1).value = "Total Inflow";
      TotalInflowAmountBillRow.getCell(1).border = this.Thinborder
    TotalInflowAmountBillRow.getCell(1).alignment = {
        horizontal:'right'
    }
    TotalInflowAmountBillRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'c6e0b4' },
        bgColor: { argb: '' },
      }
      TotalInflowAmountBillRow.getCell(2).alignment = {
        horizontal:'right'
      }
      TotalInflowAmountBillRow.getCell(1).font={
        size: 10,
        bold: true,
        name :  'Calibri' ,
      }
     TotalInflowAmountBillRow.getCell(2).value = Number(this.DataList[0].TotalINPRJ);
      TotalInflowAmountBillRow.getCell(2).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'c6e0b4' },
        bgColor: { argb: '' },
      }
      TotalInflowAmountBillRow.getCell(2).border = this.Thinborder
      TotalInflowAmountBillRow.getCell(3).value = Number(this.DataList[0].TotalPRji);
      TotalInflowAmountBillRow.getCell(3).alignment = {
          horizontal:'right'
        }
      TotalInflowAmountBillRow.getCell(3).border = this.SideBoldborder
      TotalInflowAmountBillRow.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'c6e0b4' },
        bgColor: { argb: '' },
    }
    let BlankRow6 = worksheetInflowThree.addRow([])
        BlankRow6.getCell(3).border = {
         right: { style: 'medium' },
      }
    let OutflowRow = worksheetInflowThree.addRow([])
        OutflowRow.getCell(1).value = "Outflow";
        OutflowRow.getCell(1).font={
          size: 11,
          bold: true,
          name :  'Calibri' ,
        }
        OutflowRow.getCell(1).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'f8cbad' },
          bgColor: { argb: '' },
        }
        OutflowRow.getCell(1).border = this.SideBoldborder
        OutflowRow.getCell(1).alignment = {
          horizontal:'center'
        }
        const mergeRowValue1 = (v:any) => {
          return  data8.length + v
        }
    worksheetInflowThree.mergeCells('A' + mergeRowValue1(8), 'C' + mergeRowValue1(8));
    
    let SiteName2Row = worksheetInflowThree.addRow([]);
    SiteName2Row.getCell(1).value = "Site Name";
    SiteName2Row.getCell(1).font={
       size: 10,
       bold: true,
       name :  'Calibri' ,
    }
    SiteName2Row.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'bdd7ee' },
      bgColor: { argb: '' }
    }
    SiteName2Row.getCell(1).border = this.Thinborder
    SiteName2Row.getCell(1).alignment = {
      horizontal:'center'
    }

    SiteName2Row.getCell(2).value = "Projected Outflow";
    SiteName2Row.getCell(2).font={
       size: 10,
       bold: true,
       name :  'Calibri' ,
    }
    SiteName2Row.getCell(2).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'bdd7ee' },
      bgColor: { argb: '' }
    }
    SiteName2Row.getCell(2).border = this.Thinborder
    SiteName2Row.getCell(2).alignment = {
      horizontal:'center'
    }

    SiteName2Row.getCell(3).value = "Actual Outflow";
    SiteName2Row.getCell(3).font={
       size: 10,
       bold: true,
       name :  'Calibri' ,
    }
    SiteName2Row.getCell(3).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'bdd7ee' },
      bgColor: { argb: '' }
    }
    SiteName2Row.getCell(3).border = this.SideBoldborder
    SiteName2Row.getCell(3).alignment = {
      horizontal:'center'
    }
    SiteName2Row.height = 12.75

        const data9:any = [];
        this.DataList[0].outflow.forEach((ele:any) => {
          data9.push(Object.values(ele))
        });
    data9.forEach((l:any) => {
      const row = worksheetInflowThree.addRow(l);
      for (let i = 0; i < l.length; i++) {
        row.getCell(i + 1).border = this.Thinborder
        row.getCell(3).border = this.SideBoldborder
        row.getCell(i + 1).alignment = {
          horizontal: 'left',
          vertical: 'middle',
          wrapText: true
        }
        row.getCell(i + 1).font = {
          size: 10
        }       
      }
      row.getCell(2).alignment = {
        horizontal: 'right',
        vertical: 'middle',
        wrapText: true
      }
      row.getCell(3).alignment = {
        horizontal: 'right',
        vertical: 'middle',
        wrapText: true
      }
    }
    );
    
    let TotalOutflowAmountBillRow = worksheetInflowThree.addRow([])
      TotalOutflowAmountBillRow.getCell(1).value = "Total Outflow ";
      TotalOutflowAmountBillRow.getCell(1).border = this.Thinborder
     TotalOutflowAmountBillRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8cbad' },
        bgColor: { argb: '' },
      }
    TotalOutflowAmountBillRow.getCell(1).alignment = {
        horizontal:'right'
      }
      TotalOutflowAmountBillRow.getCell(2).alignment = {
        horizontal:'right'
      }
      TotalOutflowAmountBillRow.getCell(1).font={
        size: 10,
        bold: true,
        name :  'Calibri' ,
      }
      TotalOutflowAmountBillRow.getCell(2).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8cbad' },
        bgColor: { argb: '' },
      }
      TotalOutflowAmountBillRow.getCell(2).border = this.Thinborder
      TotalOutflowAmountBillRow.getCell(3).value = Number(this.DataList[0].TotalActout);
      TotalOutflowAmountBillRow.getCell(3).alignment = {
          horizontal:'right'
        }
      TotalOutflowAmountBillRow.getCell(3).border = this.SideBoldborder
      TotalOutflowAmountBillRow.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f8cbad' },
        bgColor: { argb: '' },
      }
      let BlankRow7 = worksheetInflowThree.addRow([])
          BlankRow7.getCell(3).border = {
          right: { style: 'medium' },
      }
    
    let DifferenceRow = worksheetInflowThree.addRow([])
    DifferenceRow.getCell(1).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'medium' },
      right: { style: 'thin' },
    };
    DifferenceRow.getCell(1).value = "Difference (Inflow - Outflow)"
    DifferenceRow.getCell(1).font={
       size: 10,
       bold: true,
       color : { argb: 'f6f7f7' },
       name :  'Calibri' ,
    }
    DifferenceRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2f75b5' },
        bgColor: { argb: '' },
    }
    DifferenceRow.getCell(1).alignment = {
      horizontal:'right'
    }
    DifferenceRow.getCell(1).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'medium' },
    };
    DifferenceRow.getCell(2).value = Number(this.DataList[0].TotalINPRJ);
    DifferenceRow.getCell(2).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'medium' },
    };
    DifferenceRow.getCell(2).alignment = {
      horizontal:'right'
    }
    DifferenceRow.getCell(3).value =  Number(this.DataList[0].TotalPRji - this.DataList[0].TotalActout)
    DifferenceRow.getCell(3).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'medium' },
      right: { style: 'medium' },
    };
    DifferenceRow.getCell(3).alignment = {
      horizontal:'right'
    }
    DifferenceRow.height = 13.57


    //Generate & Save Excel File
      workbook.xlsx.writeBuffer().then((data) => {
        let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        fs.saveAs(blob, 'Projected_V_Actual'+'.xlsx');
      })
  } 
}

