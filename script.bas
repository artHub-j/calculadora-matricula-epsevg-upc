function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Calculo Matricula Menu")
    .addItem("Recalcular", "recalculateTot")
    .addToUi();
}

function recalculateTot() {
  superats();
  dedicacio();
  preu_matricula();
}

var cellPairs = [
  "D4", "E4", "D6", "E6", "D9", "E9", "D11", "E11", "D14", "E14",
  "D16", "E16", "D19", "E19", "D20", "E20", "D23", "E23", "G4", "H4",
  "G6", "H6", "G9", "H9", "G11", "H11", "G14", "H14", "G16", "H16",
  "G19", "H19", "G20", "H20", "G23", "H23", "J4", "K4", "J6", "K6",
  "J9", "K9", "J11", "K11", "J14", "K14", "J16", "K16", "J19", "K19",
  "J20", "K20", "J23", "K23", "J24", "K24", "M4", "N4", "M6", "N6",
  "M9", "N9", "M11", "N11", "M14", "N14", "M16", "N16", "M19", "N19",
  "M20", "N20", "M23", "N23", "M24", "N24", "P9", "Q9", "P11", "Q11",
  "P14", "Q14", "P16", "Q16", "P19", "Q19", "P20", "Q20", "P23", "Q23",
  "P24", "Q24", "P25", "Q25", "S23", "T23"
];

function preu_matricula() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var totalPrice = 0;

  for (var i = 0; i < cellPairs.length; i += 2) {
    var creditCell = sheet.getRange(cellPairs[i]);
    var selCell = sheet.getRange(cellPairs[i + 1]);

    var credits = creditCell.getValue();
    var selection = selCell.getValue();

    if (selection == "X") {
      var color = selCell.getBackground().toLowerCase(); // Convert to lowercase for case-insensitive comparison

      if (color == "#f3f3f3") {
        // 1a Vegada
        totalPrice += 18.46 * credits;
      } else if (color == "#e8b8b0") {
        // 2a Vegada
        totalPrice += 28.0 * credits;
      } else if (color == "#e17d6d") {
        // 3a Vegada
        totalPrice += 65.0 * credits;
      } else if (color == "#aa1a0b") {
        // 4ta Vegada
        totalPrice += 88.0 * credits;
      }
    }
  }

  totalPrice += 90.66;

  if (totalPrice == 90.66) sheet.getRange("C32").setValue(0);
  else sheet.getRange("C32").setValue(totalPrice);

  if (totalPrice == 90.66) return 0;
  else return totalPrice;
}

function superats() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var superats = 0;

  for (var i = 0; i < cellPairs.length; i += 2) {
    var creditCell = sheet.getRange(cellPairs[i]);
    var selCell = sheet.getRange(cellPairs[i + 1]);

    var credits = creditCell.getValue();
    var selection = selCell.getValue();
    var color = selCell.getBackground().toLowerCase(); // Convert to lowercase for case-insensitive comparison

    if (color == "#b6d7a8" && selection == "-") {
      superats += credits;
    }
  }

  // Update a cell with the total price (optional)
  sheet.getRange("C27").setValue(superats);

  return superats;
}

function dedicacio() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var hores = 0;

  for (var i = 0; i < cellPairs.length; i += 2) {
    var creditCell = sheet.getRange(cellPairs[i]);
    var selCell = sheet.getRange(cellPairs[i + 1]);

    //var credits = creditCell.getValue();
    var selection = selCell.getValue();

    if (selection == "X") {
      var color = creditCell.getBackground().toLowerCase(); // Convert to lowercase for case-insensitive comparison

      if (color == "#fee599") {
        hores += 187.5;
      } else if (color == "#ff6d01") {
        hores += 450;
      } else {
        hores += 150;
      }
    }
  }

  // Update a cell with the total price (optional)
  sheet.getRange("C30").setValue(hores);

  return hores;
}
