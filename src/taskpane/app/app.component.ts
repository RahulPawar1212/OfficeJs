import { Component, EventEmitter, Output } from "@angular/core";
import { wendydata } from 'interfaces';
import { map } from 'rxjs';

/* global console, Excel */
var testupdate = "";
@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})

export default class AppComponent {
  itle = 'xml_handler';
  testarray:string[] = ["aa","bb","cc","dd"];
  //map = new Map();

  questioner = [
    { 
      "id": "1",
      "qid": "Q1",
      "type": "",
      "title": "",
      "question": "",
      "instructionForRespondent": "",
      "answersRowOptions": [1,2,3,4],
      "scaleColumnOptions": [],
      "frageomrade": "",
      "CommentsForProgrammer": [],
    },
    { 
      "id": "2",
      "qid": "Q2",
      "type": "",
      "title": "",
      "question": "",
      "instructionForRespondent": "",
      "answersRowOptions": [],
      "scaleColumnOptions": [],
      "frageomrade": "",
      "CommentsForProgrammer": [],
    }
  ];
  
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

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  selectedCountry: String = "--Choose Country--";

expression:string = "";
    
newData:wendydata[] = JSON.parse(JSON.stringify(this.questioner))


updateExp(qid:any,answr:any) {
  this.expression = "Expression : " + "f('" + qid + "')=='" + answr+"'"  
}

questionid:string = "";

answers:any = []
changeQid(qid: any) { //Angular 11
  this.questionid = qid.target.value;
  //this.states = this.Countries.find(cntry => cntry.name == country).states; //Angular 8
  this.answers = this.newData.find((ans: any) => ans.qid == qid.target.value)?.answersRowOptions; //Angular 11
  this.updateExp(qid.target.value,"")
  console.log(this.answers)
  
}

changeAnswer(answers:any) { //Angular 11
  //this.states = this.Countries.find(cntry => cntry.name == country).states; //Angular 8
  //this.answers = this.newData.find((ans: any) => ans.qid == qid.target.value)?.answersRowOptions; //Angular 11
  this.updateExp(this.questionid,answers.target.value)
  console.log(answers.target.value)
  
}


  async registerClickHandler() {
    await Excel.run(async (context) => {


    })

  }}