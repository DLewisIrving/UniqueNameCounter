function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Check if the edit was in Column B
  if (range.columnStart === 2 && range.getSheet().getName() === "Sheet1") {
    const productName = range.getValue();
    const dataRange = sheet.getRange(1, 3, sheet.getLastRow(), 2);
    const data = dataRange.getValues();
    
    // Check if the product is already listed in Column C
    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === productName) {
        // Increment the count in Column D
        data[i][1] = (data[i][1] || 0) + 1;
        found = true;
        break;
      }
    }
    
    // If not found, add the new product and start the count at 1
    if (!found) {
      const newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 3).setValue(productName);
      sheet.getRange(newRow, 4).setValue(1);
    } else {
      // Update the range with the new counts
      dataRange.setValues(data);
    }
  }
}
