function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("MetricLog") || ss.insertSheet("MetricLog");

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Timestamp", "Site", "Shift", "Type of Entry", "Quantity", "Summary"]);
  }

  const values = e.values;
  const timestamp = values[0];
  const typeOfEntry = values[1];
  const shift = values[5];
  const quantity = parseInt(values[6]);
  const email = values[7];
  const site = email.split("@")[0].toLowerCase();

  const summary = `${site} - ${typeOfEntry}: ${quantity} (Shift: ${shift})`;

  logSheet.appendRow([
    timestamp,
    site,
    shift,
    typeOfEntry,
    quantity,
    summary
  ]);

  updateDailySummary();
  updateWeeklySummary();
  createDailyChart();
  updateDashboard();
}

function updateDailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("MetricLog");
  if (!logSheet) return;

  const summarySheet = ss.getSheetByName("DailySummary") || ss.insertSheet("DailySummary");
  const data = logSheet.getDataRange().getValues();
  data.shift();

  summarySheet.clearContents();
  summarySheet.appendRow(["Date", "Site", "Type of Entry", "Total Quantity"]);

  const summaryMap = {};

  data.forEach(row => {
    const [timestamp, site, shift, type, quantity] = row;
    const date = new Date(timestamp);
    const dateKey = date.toISOString().split("T")[0];
    const key = `${dateKey}_${site}_${type}`;
    summaryMap[key] = (summaryMap[key] || 0) + Number(quantity);
  });

  for (const key in summaryMap) {
    const [date, site, type] = key.split("_");
    summarySheet.appendRow([date, site, type, summaryMap[key]]);
  }
}

function updateWeeklySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("MetricLog");
  const weeklySheet = ss.getSheetByName("WeeklySummary") || ss.insertSheet("WeeklySummary");

  const data = logSheet.getDataRange().getValues();
  data.shift();

  weeklySheet.clearContents();
  weeklySheet.appendRow(["Week Start", "Site", "Type of Entry", "Total Quantity"]);

  const summaryMap = {};

  data.forEach(row => {
    const [timestamp, site, shift, type, quantity] = row;
    const date = new Date(timestamp);
    const day = date.getDay();
    const weekStart = new Date(date);
    weekStart.setDate(date.getDate() - day);
    weekStart.setHours(0, 0, 0, 0);
    const weekKey = weekStart.toISOString().split("T")[0];
    const key = `${weekKey}_${site}_${type}`;
    summaryMap[key] = (summaryMap[key] || 0) + Number(quantity);
  });

  for (const key in summaryMap) {
    const [weekStart, site, type] = key.split("_");
    weeklySheet.appendRow([weekStart, site, type, summaryMap[key]]);
  }
}

function createDailyChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DailySummary");
  if (!sheet) return;

  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const range = sheet.getRange(`A2:D${lastRow}`);
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(2, 6, 0, 0)
    .setOption("title", "Daily Totals by Site and Entry Type")
    .build();

  sheet.insertChart(chart);
}

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("DailySummary");
  const dashboard = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");

  const data = summarySheet.getDataRange().getValues();
  const headers = data.shift();
  const today = new Date().toISOString().split("T")[0];

  const todayData = data.filter(row => row[0] === today);

  dashboard.clearContents();
  dashboard.appendRow(["Site", "Type of Entry", "Total Quantity"]);
  todayData.forEach(row => {
    dashboard.appendRow([row[1], row[2], row[3]]);
  });

  dashboard.autoResizeColumns(1, 3);
}

function testSubmit() {
  const mockEvent = {
    values: [
      new Date().toString(),         
      "Ingress",                     
      "3/29/2025",                   
      "5-15 min",                    
      "Routine Patrol",             
      "GV",                          
      "4",                           
      "Location3@company.com"       
    ]
  };

  onFormSubmit(mockEvent);
}

