const Excel = require("exceljs");

const workbook = new Excel.Workbook();
const sheet = workbook.addWorksheet("Test Sheet 1");

// fetch sheet by name
const worksheet = workbook.getWorksheet("Test Sheet 1");

worksheet;
