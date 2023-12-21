import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as JsonToXML from "js2xmlparser";
import { DomSanitizer } from '@angular/platform-browser';
type AOA = any[][];

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})


export class AppComponent {
  fileUrl:any;
  jsonInvoice:any;
  jsonObj:any;
  title = 'xml_convertor';
  data: AOA = [];
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'SheetJS.xlsx';
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];
debugger
      /* save data */
      this.data = <AOA>(XLSX.utils.sheet_to_json(ws, { header: 1 }));
      console.log("data:",this.data);
      this.data.map(res=>{
        if(res[0] === "no"){
          console.log(res[0]);
        }else{
          console.log(res[0]);
        }
      })
console.log(JSON.stringify(this.data));

this.xml=JSON.stringify(this.data);

    };
    reader.readAsBinaryString(target.files[0]);
  }




  onFileChange_1(ev:any) {
    let workBook : any;
    let jsonData = null;
    const reader = new FileReader();
    const file = ev.target.files[0];
    reader.onload = (event) => {
      const data = reader.result;
      workBook = XLSX.read(data, { type: 'binary' });
      jsonData = workBook.SheetNames.reduce((initial : any, name:any) => {
        const sheet = workBook.Sheets[name];
        initial[name] = XLSX.utils.sheet_to_json(sheet);
        return initial;
      }, {});
      debugger
      const dataString = JSON.stringify(jsonData);
    this.jsonInvoice = dataString.slice(0, 300).concat("...");

    console.log(jsonData);
    this.jsonObj=jsonData['Sheet1'][0];
    this.bindXml(this.jsonObj);
      
    }
    reader.readAsBinaryString(file);
  }


  export(): void {

    this.xml=JsonToXML.parse("person", this.data);
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }




  obj = {
    Invoice:{
    firstName: "John",
    lastName: "Smith",
    dateOfBirth: new Date(1964, 7, 26),
    address: {
      "@": {
        type: "home"
      },
      streetAddress: "3212 22nd St",
      city: "Chicago",
      state: "Illinois",
      zip: 10000
    },
    phone: [
      {
        "@": {
          type: "home"
        },
        "#": "123-555-4567"
      },
      {
        "@": {
          type: "cell"
        },
        "#": "890-555-1234"
      },
      {
        "@": {
          type: "work"
        },
        "#": "567-555-8901"
      }
    ],
    email: "john@smith.com"
  }
  };


  invoiceobj = {
    Invoice:{
      InvoiceHeader:{
          CustomerID: "",
          InvoiceType: "",
          InvoiceNum:'',
          PONum: '',
          AttachmentList:{
            Attachment:{
              FileName:''
            }

          },
          UserDefined11:'',
          UserDefined12:'',
          InvoiceParty:{
            Role:'',
            ContactName:'',
            ContactPhone:''

          }
    },
   
    InvoiceLine: []=[],
      // {
      //   LineNum:'1',
      //   Quantity:'943.68',
      //   UnitPrice:'1.00',
      //   POLineNum:'5',
      //   POSchedNum:'1',
      //   ServiceStartDate:'20221121',
      //   ServiceEndDate:'20221121'
      // },
      // {
      //   LineNum:'2',
      //   Quantity:'92.68',
      //   UnitPrice:'2.00',
      //   POLineNum:'6',
      //   POSchedNum:'3',
      //   ServiceStartDate:'20221123',
      //   ServiceEndDate:'20221123'
      // },
    //],
    InvoiceSummary: {
      InvoiceTotNetVal:'',
      InvoiceTotGrossVal:''

    }
  }
  };

  bindXml(obj:any)
  {
this.invoiceobj.Invoice.InvoiceHeader.CustomerID=obj.CustomerID;
this.invoiceobj.Invoice.InvoiceHeader.InvoiceType=obj.InvoiceType;
this.invoiceobj.Invoice.InvoiceHeader.InvoiceNum=obj.InvoiceNum;
this.invoiceobj.Invoice.InvoiceHeader.PONum=obj.PONum;
this.invoiceobj.Invoice.InvoiceHeader.AttachmentList.Attachment.FileName=obj.FileName;
this.invoiceobj.Invoice.InvoiceHeader.UserDefined11=obj.UserDefined11;
this.invoiceobj.Invoice.InvoiceHeader.UserDefined12=obj.UserDefined12;
this.invoiceobj.Invoice.InvoiceHeader.InvoiceParty.Role=obj.Role;
this.invoiceobj.Invoice.InvoiceHeader.InvoiceParty.ContactName=obj.ContactName;
this.invoiceobj.Invoice.InvoiceHeader.InvoiceParty.ContactPhone=obj.ContactPhone;

let items= {
  
    LineNum:obj.LineNum,
    Quantity:obj.Quantity,
    UnitPrice:obj.UnitPrice,
    POLineNum:obj.POLineNum,
    POSchedNum:obj.POSchedNum,
    ServiceStartDate:obj.ServiceStartDate,
    ServiceEndDate:obj.ServiceEndDate
}

this.invoiceobj.Invoice.InvoiceLine.push(items as never);
this.invoiceobj.Invoice.InvoiceSummary.InvoiceTotNetVal=obj.InvoiceTotNetVal;
this.invoiceobj.Invoice.InvoiceSummary.InvoiceTotGrossVal=obj.InvoiceTotGrossVal;

this.xml=JsonToXML.parse("InvoiceList", this.invoiceobj);


const data = this.xml;
    const blob = new Blob([data], { type: 'application/octet-stream' });

    this.fileUrl = this.sanitizer.bypassSecurityTrustResourceUrl(window.URL.createObjectURL(blob));

  }

  xml:string="";
   constructor(private sanitizer: DomSanitizer) {
    // this.xml=JsonToXML.parse("InvoiceList", this.invoiceobj);
  }
}