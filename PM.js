function myFunction() {

    var app = SpreadsheetApp.getActive();
    var outputsheet = app.getSheetByName("Sheet1");
    var lastRow = outputsheet.getLastRow() - 1;
    var lastColumn = outputsheet.getLastColumn();
    var allsheet = outputsheet.getRange(1, 1, lastRow, lastColumn).getValues();
    
    var artifactID = outputsheet.getRange(2, 2, lastRow, 1).getValues();
    var estimatedEffort = outputsheet.getRange(2, 9, lastRow, 1).getValues();
    var remainingEffort = outputsheet.getRange(2, 10, lastRow, 1).getValues();
    var actualEffort = outputsheet.getRange(2, 11, lastRow, 1).getValues();
    var trackerName = outputsheet.getRange(2, 19, lastRow, 1).getValues();
    var dependencyParent = outputsheet.getRange(2, 18, lastRow, 1).getValues();
   
  //change the range to make it accessble to any column possible possition that the user could change.
  //all the ranges will changes until ranges in the gauge chart.
  
    var  progressStatus = outputsheet.getRange(2,21,lastRow,1).getValues();
      for (var i = 0; i < trackerName.length; i++) {
      
      var estEffort = parseFloat(estimatedEffort[i][0]);
      var remEffort = parseFloat(remainingEffort[i][0]);
      var actEffort = parseFloat(actualEffort[i][0]);
  
      if (
        isNaN(estEffort) || isNaN(remEffort) || isNaN(actEffort) ||
        estEffort < 0 || remEffort < 0 || actEffort < 0 ||
        estimatedEffort[i][0] === "" || remainingEffort[i][0] === "" || actualEffort[i][0] === "" ||
        estimatedEffort[i][0] === String || remainingEffort[i][0] === String || actualEffort[i][0] === String
      ) {
  
        outputsheet.getRange(i + 2, 11).activate();
        //SpreadsheetApp.getUi().alert("ERROR! Invalid or empty value in Estimated/Remaining/Actual Effort in row's number ("+(i+2)+") Cell value must be a non-negative number.");
        outputsheet.getRange(2+i,1, 1, 25).setBackground("red");
        SpreadsheetApp.getActiveSpreadsheet().toast("ERROR! Invalid or empty value in Estimated/Remaining/Actual Effort in row's number ("+(i+2)+") Cell value must be a non-negative number.");
        continue; // Skip to the next iteration if any of the values are invalid or empty
      }
      else{
        outputsheet.getRange(2+i,22, 1,4).setBackground("green");
      }
  
        var pro = progressPercentages(remEffort, actEffort);
        var lOver = leftOver(remEffort, actEffort);
        var effortVariance = projectedEffortVariance(estEffort, remEffort, actEffort);
  
        if (!isNaN(pro)) {
          outputsheet.getRange(2 + i, 22).setValue(pro.toFixed(2));
        }
        else {
          SpreadsheetApp.getActiveSpreadsheet().toast("the cell number " + (2 + i) + " in column prgress percentage is empty because its value is NaN");
          outputsheet.getRange(2+i,1, 1,25).setBackground("red");
          continue;
        }
  
        if (!isNaN(lOver)) {
          outputsheet.getRange(2 + i, 23).setValue(lOver.toFixed(2));
        }
        else {
          SpreadsheetApp.getActiveSpreadsheet().toast("the cell number " + (2 + i) + " in column leftover is empty because its value is NaN");
          outputsheet.getRange(2+i,1, 1,25).setBackground("red");
          continue;
        }
  
        outputsheet.getRange(2 + i, 24).setValue(effortVariance.toFixed(2));
        var pstatus = effortVariance;
  
        if (pstatus == 1) {
          Logger.log("Ahead");
          outputsheet.getRange(2 + i, 25).setValue("Ahead");
        }
        else if (pstatus == 0) {
          Logger.log("Completed");
          outputsheet.getRange(2 + i, 25).setValue("Completed");
        }
        else if (pstatus == -1) {
          Logger.log("Behind");
          outputsheet.getRange(2 + i, 25).setValue("Behind");
        }
        else {
          Logger.log("N/A");
          outputsheet.getRange(2 + i, 25).setValue("N/A");
        }
    
      }
  var iterator2 = 0;
  for (var i = 0; i < artifactID.length; i++) {
    if (trackerName[i] == "Story") {
      var idRange;
      var pRange;
  
      if (dependencyParent[i][0] !== " ") {
        var iter = 2;
        pRange = outputsheet.getRange(iter+i, 22, 1, 1);
        idRange = outputsheet.getRange(iter + i, 2, 1, 1);
      } else {
        continue;
      }
      
      if (artifactID[i][0].length !== 0 && pRange.getValue() !== "") {
        var chart = outputsheet.newChart()
          .setChartType(Charts.ChartType.GAUGE)
          .addRange(idRange)
          .addRange(pRange)
          .setPosition(3+lastRow + (6*iterator2), 1 , 0, 0)
          .setOption('height', 120)
          .setOption('width', 120)
          .setOption('redFrom', 0)
          .setOption('redTo', 50)
          .setOption('yellowFrom', 50)
          .setOption('yellowTo', 75)
          .setOption('greenFrom', 75)
          .setOption('greenTo', 100)
          .setOption("theme", "maximized")
          .build();
        
        outputsheet.insertChart(chart);
        iterator2 += 2;
      } else {
        continue;
      }
    } else {
      continue;
    }
  
  }
  var iterator = 0;
  var unique_repeated;
  var  progressStatus2 = outputsheet.getRange(2,25,lastRow,1).getValues();
  //Logger.log(progressStatus2);
  for (var i = 0; i < artifactID.length; i++) {
    var valuesToSet =[];
    var countrow = [];
  
    for (var j = 0; j < dependencyParent.length; j++){
    var count2 = 0;
  
      if (artifactID[i][0] === dependencyParent[j][0]){
      
        countrow.push(progressStatus2[j]);
        valuesToSet.push.apply(valuesToSet, countrow);
        count2 =valuesToSet.length;
        //Logger.log(dependencyParent[j] +" "+  count2 + " and status is " +progressStatus[j]);
        outputsheet.getRange(2,50+i,valuesToSet.length,1).setValues(valuesToSet);  
        unique_repeated =artifactID[i][0];
        var lastcounter= countrow.lastIndexOf(countrow[countrow.length-1]); 
        
        if(valuesToSet.length %  lastcounter === 0){
          var pieChartBuilder = outputsheet.newChart()
            .addRange(outputsheet.getRange(2,50+ i,count2,1)) //
            .setChartType(Charts.ChartType.PIE)
            .setOption('pieSliceText', 'value')
            .setPosition(lastRow+9 + (12* iterator), 1, 0, 0)
            .setOption('title', unique_repeated)
            .setOption('legend', { textStyle: { color: 'blue', fontSize: 12 } })
            .setOption('width', 220)
            .setOption('height', 125)
            .setOption('applyAggregateData', 0)
            .build();
            
          // Insert the pie chart
          //outputsheet.insertChart(pieChartBuilder);
          iterator += 1;
  }
  //outputsheet.getRange(2,50+i,valuesToSet.length,1).setValues(valuesToSet);  
  valuesToSet.splice(0);      
  }
  else{
        continue;
      }
  }
  // outputsheet.getRange(2,23+i,count2,1).clear();
  }
  }
  
  
  function progressPercentages(rem1, act1) {
  
    var sum1 = rem1 + act1;
    var proPercentages = (act1 / sum1) * 100;
    return proPercentages;
  
  }
  
  function leftOver(rem2, act2) {
    var sum2 = rem2 + act2;
    var leftOver = (rem2 / sum2) * 100;
    return leftOver;
  
  }
  
  function projectedEffortVariance(est, rem3, act3) {
    var sum3 = rem3 + act3;
    var varience = est - sum3;
    return varience;
  }