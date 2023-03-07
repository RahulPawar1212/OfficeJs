import { Component, EventEmitter, Output } from "@angular/core";

/* global console, Excel */
var testupdate = "";
@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})

export default class AppComponent {
  welcomeMessage = "Welcomesss";
  colortxt = "green";
  testarray = ["aa","bb","cc","dd"];
  testarray2;
  testarray3:any = [];
  @Output() reviewSubmitted = new EventEmitter<string>();

  async run() {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";
        this.welcomeMessage = "heloow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

 

  async  registerClickHandler() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      const range1 = sheet.getRange("G4:K4")
      range1.load("values");
      await context.sync()
      
      //console.log(range1)
      //console.log(JSON.stringify(range1.values, null, 4));
      // var markets =   JSON.stringify(option2[i]).slice(1,-1).toString().replaceAll(/},{/g,',')

      sheet.onSingleClicked.add((event) => {
        return Excel.run((context) => {
          console.log(
            `Click detected at ${event.address} (pixel offset from upper-left cell corner: ${event.offsetX}, ${event.offsetY})`
          );

          this.testarray2 = JSON.stringify(range1.values, null, 4);
          this.testarray2 =  this.testarray2.slice(1,-1)
          this.testarray3 = eval(this.testarray2)
      
          console.log(this.testarray3)

         

          this.onchange (`${event.address}`);
          this.colortxt = "red";  
            Promise.resolve(this.onchange (`${event.address}`)).then(v=>
            {
              this.welcomeMessage = v;
            })
          
          //this.reviewSubmitted.emit(this.welcomeMessage);
          
          console.log(this.welcomeMessage);
          return context.sync();
        });
      });
  
      console.log("The worksheet click handler is registered.");
  
      await context.sync();
    });
  }

   onchange(texts: string) {
    this.welcomeMessage = texts
    document.getElementById("p1").innerHTML = texts
    return texts
  }
  
}
