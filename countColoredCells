function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Nells Functions')
    .addItem('Update Color Counts', 'countIfsColoredCells')
    .addToUi();
}

// countifs colored cells
function countIfsColoredCells(countRange,colorRef1,colorRef2,criteria1ApplyOverRange,criteria1Range,criteria2ApplyOverRange,criteria2Range) {
  const activeRange = SpreadsheetApp.getActiveRange();
  const activeSheet = activeRange.getSheet();
  const formula = activeRange.getFormula();


//extract arguments from function
  const argsMatch = formula.match(/\((.+)\)/);
    if (!argsMatch) {
      throw new Error("Invalid formula format. Expected =FUNCTION(range,colorCell,criteria1ApplyOverRange,criteria1Range,criteria2ApplyOverRange,criteria2Range)");
    }


// split arguments into placement range, target color 1 and 2, critera 1 range, criteria 1, criteria 2 range, and criteria 2
  const args = argsMatch[1].split(/[,;]/);
  Logger.log("Arguments: " + args);
  Logger.log("Arg array length: " + args.length);

  const rangeA1Notation = args[0].trim();
  const color1CellA1Notation = args[1].trim();
  const color2CellA1Notation = args[2].trim();
  const crit1RangeToApplyA1Notation = args[3].trim(); // Range in which we apply C1 to (in master sheet)
  const crit1RangeA1Notation = args[4].trim(); // Array of C1s to apply (ie districts)
  const crit2RangeToApplyA1Notation = args[5].trim(); // Range in which we apply C2 to (in master sheet)
  const crit2RangeA1Notation = args[6].trim(); // Array of C2s to apply (ie M or F)

//  Logger.log("Criteria 1: " + crit1RangeA1Notation);
//  Logger.log("Criteria 2: " + crit2RangeA1Notation);


// get data
  const range = activeSheet.getRange(rangeA1Notation);
  const bg = range.getBackgrounds();

  const colorCell1 = activeSheet.getRange(color1CellA1Notation);
  const targetColor1 = colorCell1.getBackground();

  const colorCell2 = activeSheet.getRange(color2CellA1Notation);
  const targetColor2 = colorCell2.getBackground();

  const crit1Val = activeSheet.getRange(crit1RangeA1Notation).getValues(); // Range of crit 1 to loop over
  const c1RangeToApplyVals = activeSheet.getRange(crit1RangeToApplyA1Notation).getValues(); // Range in which to apply crit 1 to also loop over
//  const c1rangeVals = c1range.getValues(); // Array of 

  const crit2Val = activeSheet.getRange(crit2RangeA1Notation).getValues(); // Range of crit 2 to loop over
  const c2RangeToApplyVals = activeSheet.getRange(crit2RangeToApplyA1Notation).getValues(); // Range in which to apply crit 2 to also loop over 
//  const c2rangeVals = c2range.getValues();


// count colors
  let count = 0;
  let counts_array = [];
  let one_m = 0;
  let indices = [];
//  Logger.log("bg array length: " + bg.length);
  Logger.log("Crit1 val: " + crit1Val);
  Logger.log("Crit2 val: " + crit2Val);
  
// loop over Criteria 1 first, then Criteria 2 (for each Crit 1, ie. 1F, 1M, 2F, 2M ...)
for (let m=0; m<crit1Val.length; m++){
  // Logger.log("M: " + crit1Val[m]);

  for (let n=0; n<crit2Val.length; n++){
    //Logger.log("N: " + n);

    // reset count 
    count = 0;
    // loop over background color arrays 
    for (let i=0; i<bg.length; i++){
      for (let j=0; j<bg[i].length; j++){
        // Apply Crit 1 instance to Crit 1 range (districts) and Crit 2 instance to Crit 2 range (gender slot)
        if( c1RangeToApplyVals[i][j] == crit1Val[m] && c2RangeToApplyVals[i][j] == crit2Val[n]){
          one_m++;
          if ( bg[i][j] == targetColor1 || bg[i][j] == targetColor2){
            indices.push(i);
            count++;
          }
        }
      }
    }
    // move on to next criteria; save count to array
    counts_array.push(count);
  }
}

//  Logger.log("1Ms: " + one_m); 
//  Logger.log(indices); 
  const col_counts = counts_array.map(item => [item]); 
  Logger.log("Final array: " + col_counts);

  return col_counts;

}
