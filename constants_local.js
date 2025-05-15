const localNamedRanges = {
  "formulaRunArchiveDriverId": {
    "sheetName":"Run Archive",
    "headerName":"Driver ID"
  },
  "formulaRunArchiveVehicleId": {
    "sheetName":"Run Archive",
    "headerName":"Vehicle ID"
  },
  "formulaRunArchiveRunId": {
    "sheetName":"Run Archive",
    "headerName":"Run ID"
  },
  "formulaRunArchiveRunKey": {
    "sheetName":"Run Archive",
    "headerName":"Run Key"
  },
    "formulaRunArchiveTotalVehicleMiles": {
    "sheetName":"Run Archive",
    "headerName":"Total Vehicle Miles"
  },
  "formulaRunArchiveTotalVehicleHours": {
    "sheetName":"Run Archive",
    "headerName":"Total Vehicle Hours"
  },
"formulaTripArchiveDriverId": {
    "sheetName":"Trip Archive",
    "headerName":"Driver ID"
  },
  "formulaTripArchiveVehicleId": {
    "sheetName":"Trip Archive",
    "headerName":"Vehicle ID"
  },
  "formulaTripArchiveRunId": {
    "sheetName":"Trip Archive",
    "headerName":"Run ID"
  },
  "formulaTripArchiveGuests": {
    "sheetName":"Trip Archive",
    "headerName":"Guests"
  },
  "formulaTripArchiveRiderCount": {
    "sheetName":"Trip Archive",
    "headerName":"Rider Count"
  },
  "formulaTripArchiveRunKey": {
    "sheetName":"Trip Archive",
    "headerName":"Run Key"
  },
  "formulaTripArchiveEstHours": {
    "sheetName":"Trip Archive",
    "headerName":"Est Hours"
  },
  "formulaTripArchiveSumOfRunEstHours": {
    "sheetName":"Trip Archive",
    "headerName":"Sum of Run Est Hours"
  },
  "formulaTripArchiveRunDuration": {
    "sheetName":"Trip Archive",
    "headerName":"Run Duration"
  },
  "formulaTripArchiveRunDurationPortion": {
    "sheetName":"Trip Archive",
    "headerName":"Run Duration Portion"
  },
  "formulaTripArchiveBillableRunDuration": {
    "sheetName":"Trip Archive",
    "headerName":"Billable Run Duration"
  },
  "formulaTripArchiveEstMiles": {
    "sheetName":"Trip Archive",
    "headerName":"Est Miles"
  },
  "formulaTripArchiveSumOfRunEstMiles": {
    "sheetName":"Trip Archive",
    "headerName":"Sum of Run Est Miles"
  },
  "formulaTripArchiveRunMileage": {
    "sheetName":"Trip Archive",
    "headerName":"Run Mileage"
  },
  "formulaTripArchiveRunMileagePortion": {
    "sheetName":"Trip Archive",
    "headerName":"Run Mileage Portion"
  },
}
const localNamedRangesToRemove = []

const localSheetsToRemove = []
const localSheets = []
const localSheetsWithHeaders = []

const localColumnsToRemove = {}
const localColumns = {
  "Run Archive": {
    "Run Key": {
      headerFormula: `={"Run Key";MAP(formulaRunArchiveRunDate, formulaRunArchiveDriverId,formulaRunArchiveVehicleId,formulaRunArchiveRunId,LAMBDA(runDate,driverId,vehicleId,runId,IF(COUNTBLANK(runDate,driverId,vehicleId)>0,"",TEXT(runDate,"mm/dd/yyyy")&"-"&driverid&"-"&vehicleId&"-"&runId)))}`,
      numberFormat: "@"
    },
    "Trip Count": {
      headerFormula: `={"Trip Count";MAP(formulaRunArchiveRunKey, LAMBDA(runKey,IF(COUNTBLANK(runKey)>0,"",COUNTIFS(formulaTripArchiveRunKey, runKey))))}`,
      numberFormat: "0"
    },
    "Rider Count": {
      headerFormula: `={"Rider Count";MAP(formulaRunArchiveRunKey, LAMBDA(runKey,IF(COUNTBLANK(runKey)>0,"",SUMIFS(formulaTripArchiveRiderCount,formulaTripArchiveRunKey, runKey))))}`,
      numberFormat: "0"
    },
  },
  "Trip Archive": {
    "Trip Sequence": {
      headerFormula: `={"Trip Sequence";MAP(formulaTripArchiveTripDate, formulaTripArchiveCustomerId, formulaTripArchivePuTime, LAMBDA(tripDate,customerId, PuTime, IF(ISBLANK(tripDate),"",COUNTIFS(formulaTripArchiveTripDate, tripDate, formulaTripArchiveCustomerId, customerId, formulaTripArchivePuTime, "<" & PuTime)+1)))}`,
      numberFormat: "0"
    },
    "Trip Reverse Sequence": {
      headerFormula: `={"Trip Reverse Sequence";MAP(formulaTripArchiveTripDate, formulaTripArchiveCustomerId, formulaTripArchivePuTime, LAMBDA(tripDate,customerId, PuTime, IF(ISBLANK(tripDate),"",COUNTIFS(formulaTripArchiveTripDate, tripDate, formulaTripArchiveCustomerId, customerId, formulaTripArchivePuTime, ">" & PuTime)+1)))}`,
      numberFormat: "0"
    },
    "Run Key": {
      headerFormula: `={"Run Key";MAP(formulaTripArchiveTripDate, formulaTripArchiveDriverId,formulaTripArchiveVehicleId,formulaTripArchiveRunId,LAMBDA(tripDate,driverId,vehicleId,runId,IF(COUNTBLANK(runDate,driverId,vehicleId)>0,"",TEXT(tripDate,"mm/dd/yyyy")&"-"&driverid&"-"&vehicleId&"-"&runId)))}`,
      numberFormat: "@"
    },
    "Rider Count": {
      headerFormula: `={"Rider Count";MAP(formulaTripArchiveTripDate, formulaTripArchiveGuests, LAMBDA(tripDate, guests, IF(COUNTBLANK(tripDate)>0,"",1+guests)))}`,
      numberFormat: "0"
    },
    "Run Duration": {
      headerFormula: `={"Run Duration";map(formulaTripArchiveRunKey,LAMBDA(runKey,XLOOKUP(runKey,formulaRunArchiveRunKey,formulaRunArchiveTotalVehicleHours,0)))}`,
      numberFormat: "[h]:mm"
    },
    "Sum of Run Est Hours": {
      headerFormula: `={"Sum of Run Est Hours";MAP(formulaTripArchiveRunKey, LAMBDA (runKey,SUMIF(formulaTripArchiveRunKey,runKey,formulaTripArchiveEstHours)))}`,
      numberFormat: "[h]:mm"
    },
    "Run Duration Portion": {
      headerFormula: `={"Run Duration Portion";MAP(formulaTripArchiveEstHours,formulaTripArchiveSumOfRunEstHours,LAMBDA(estHours,sumOfEstHours,IF(COUNTBLANK(estHours,sumOfEstHours)>0,"",estHours/sumOfEstHours)))}`,
      numberFormat: "0.00%"
    },
    "Billable Run Duration": {
      headerFormula: `={"Billable Run Duration";MAP(formulaTripArchiveRunDuration,formulaTripArchiveRunDurationPortion,LAMBDA(runDuration,runDurationPortion,IF(COUNTBLANK(runDuration,runDurationPortion)>0,"",runDuration*runDurationPortion)))}`,
      numberFormat: "[h]:mm"
    },
    "Billable Run Seconds": {
      headerFormula: `={"Billable Run Seconds";MAP(formulaTripArchiveBillableRunDuration,LAMBDA(brd,IF(COUNTBLANK(brd)>0,"",brd*86400)))}`,
      numberFormat: "#,##0"
    },
    "Run Mileage": {
      headerFormula: `={"Run Mileage";map(formulaTripArchiveRunKey,LAMBDA(runKey,XLOOKUP(runKey,formulaRunArchiveRunKey,formulaRunArchiveTotalVehicleMiles,0)))}`,
      numberFormat: "0.00"
    },
    "Sum of Run Est Miles": {
      headerFormula: `={"Sum of Run Est Miles";MAP(formulaTripArchiveRunKey, LAMBDA (runKey,SUMIF(formulaTripArchiveRunKey,runKey,formulaTripArchiveEstMiles)))}`,
      numberFormat: "0.00"
    },
    "Run Mileage Portion": {
      headerFormula: `={"Run Mileage Portion";MAP(formulaTripArchiveEstMiles,formulaTripArchiveSumOfRunEstMiles,LAMBDA(estHours,sumOfEstHours,IF(COUNTBLANK(estHours,sumOfEstHours)>0,"",estHours/sumOfEstHours)))}`,
      numberFormat: "0.00%"
    },
    "Billable Run Mileage": {
      headerFormula: `={"Billable Run Mileage";MAP(formulaTripArchiveRunMileage,formulaTripArchiveRunMileagePortion,LAMBDA(runMileage,runMileagePortion,IF(COUNTBLANK(runMileage,runMileagePortion)>0,"",runMileage*runMileagePortion)))}`,
      numberFormat: "0.00"
    },
  }
}
