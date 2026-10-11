function getTransactionData(sheet, type, theLatestRow)
{
  let valueCol;

  if (type === "Number Of Transactions") {
    valueCol = 5;
  } else if (type === "Total Amount") {
    valueCol = 6;
  } else {
    throw new Error("Invalid input file");
  }

  const headerCell = sheet.getRange(5, valueCol).getValue().toString().trim();
  if (headerCell !== type) {
    throw new Error("Invalid input file");
  }

  let row = 6;

  let OMtoMTN_LRD  = 0;
  let OMtoMTN_USD  = 0;
  let MTNtoOM_LRD  = 0;
  let MTNtoOM_USD  = 0;
  let GovToOM_LRD  = 0;
  let GovToOM_USD  = 0;
  let GovToMTN_LRD = 0;
  let GovToMTN_USD = 0;

  while (row <= theLatestRow) {
    const colA = sheet.getRange(row, 1).getValue().toString().trim();
    const colC = sheet.getRange(row, 3).getValue().toString().trim();
    const colG = sheet.getRange(row, 7).getValue().toString().trim();
    const val  = parseFloat(sheet.getRange(row, valueCol).getValue());
    if (val == null || isNaN(val) || val === "")
    {
      throw new Error("Check your input file");
    }

    if (colG === 'LRD') {
      if (colA === 'mwallet0383' && colC === 'mwallet0096') OMtoMTN_LRD  += val; // OM → MTN
      if (colA === 'mwallet0096' && colC === 'mwallet0383') MTNtoOM_LRD  += val; // MTN → OM
      if (colA === 'gov0448'     && colC === 'mwallet0383') GovToOM_LRD  += val; // Gov → OM
      if (colA === 'gov0448'     && colC === 'mwallet0096') GovToMTN_LRD += val; // Gov → MTN

    } else if (colG === 'USD') {
      if (colA === 'mwallet0383' && colC === 'mwallet0096') OMtoMTN_USD  += val; // OM → MTN
      if (colA === 'mwallet0096' && colC === 'mwallet0383') MTNtoOM_USD  += val; // MTN → OM
      if (colA === 'gov0448'     && colC === 'mwallet0383') GovToOM_USD  += val; // Gov → OM
      if (colA === 'gov0448'     && colC === 'mwallet0096') GovToMTN_USD += val; // Gov → MTN
    }
    row++;
  }

  return [
    OMtoMTN_LRD,
    OMtoMTN_USD,
    MTNtoOM_LRD,
    MTNtoOM_USD,
    GovToOM_LRD,
    GovToOM_USD,
    GovToMTN_LRD,
    GovToMTN_USD
  ];
}

function  getTheValues(sheet, theLatestRow, exchangeRate)
{
  // Get the transaction Counts
  [
    OMtoMTN_LRD_Count,
    OMtoMTN_USD_Count,
    MTNtoOM_LRD_Count,
    MTNtoOM_USD_Count,
    GovToOM_LRD_Count,
    GovToOM_USD_Count,
    GovToMTN_LRD_Count,
    GovToMTN_USD_Count
  ] = getTransactionData(sheet, "Number Of Transactions", theLatestRow);

  // Get the transaction Amounts
  [
    OMtoMTN_LRD_Amount,
    OMtoMTN_USD_Amount,
    MTNtoOM_LRD_Amount,
    MTNtoOM_USD_Amount,
    GovToOM_LRD_Amount,
    GovToOM_USD_Amount,
    GovToMTN_LRD_Amount,
    GovToMTN_USD_Amount
  ] = getTransactionData(sheet, "Total Amount", theLatestRow);
  
  // Calculation Trn Count
  P2PLRDTransactionCount = OMtoMTN_LRD_Count + MTNtoOM_LRD_Count;
  G2PLRDTransactionCount = GovToOM_LRD_Count + GovToMTN_LRD_Count;
  P2PUSDTransactionCount = OMtoMTN_USD_Count + MTNtoOM_USD_Count;
  G2PUSDTransactionCount = GovToOM_USD_Count + GovToMTN_USD_Count;

  // Calculation Trn Amount
  P2PLRDTransactionAmount = OMtoMTN_LRD_Amount + MTNtoOM_LRD_Amount;
  P2PUSDTransactionAmount = OMtoMTN_USD_Amount + MTNtoOM_USD_Amount;
  G2PLRDTransactionAmount = GovToOM_LRD_Amount + GovToMTN_LRD_Amount;
  G2PUSDTransactionAmount = GovToOM_USD_Amount + GovToMTN_USD_Amount;

  // Total
  totalLRDTransactionCount = P2PLRDTransactionCount + G2PLRDTransactionCount;
  totalUSDTransactionCount = P2PUSDTransactionCount + G2PUSDTransactionCount;
  totalLRDTransactionAmount = P2PLRDTransactionAmount + G2PLRDTransactionAmount;
  totalUSDTransactionAmount = P2PUSDTransactionAmount + G2PUSDTransactionAmount;
  totalTransactionCount = totalLRDTransactionCount + totalUSDTransactionCount;
  totalTransactionAmount = (totalLRDTransactionAmount / exchangeRate) + totalUSDTransactionAmount;

  values =  [
              OMtoMTN_LRD_Count, OMtoMTN_USD_Count, MTNtoOM_LRD_Count, MTNtoOM_USD_Count,
              GovToOM_LRD_Count, GovToOM_USD_Count, GovToMTN_LRD_Count, GovToMTN_USD_Count,
              OMtoMTN_LRD_Amount, OMtoMTN_USD_Amount, MTNtoOM_LRD_Amount, MTNtoOM_USD_Amount,
              GovToOM_LRD_Amount, GovToOM_USD_Amount, GovToMTN_LRD_Amount, GovToMTN_USD_Amount,
              P2PLRDTransactionCount, P2PUSDTransactionCount, G2PLRDTransactionCount, G2PUSDTransactionCount,
              P2PLRDTransactionAmount, P2PUSDTransactionAmount, G2PLRDTransactionAmount, G2PUSDTransactionAmount,
              totalLRDTransactionCount,  totalUSDTransactionCount, totalLRDTransactionAmount, 
              totalUSDTransactionAmount, totalTransactionCount,     totalTransactionAmount
            ];
  if (OMtoMTN_LRD_Count == 0 || OMtoMTN_USD_Count == 0 || MTNtoOM_LRD_Count == 0 || MTNtoOM_USD_Count == 0 || OMtoMTN_LRD_Amount == 0 || OMtoMTN_USD_Amount == 0 || MTNtoOM_LRD_Amount == 0 || MTNtoOM_USD_Amount == 0)
  {
    throw new Error("Check your input file");
  }
  return (values);
}