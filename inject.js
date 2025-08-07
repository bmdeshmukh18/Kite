var node = document.createElement("form"); 
var highDivNode = document.createElement("div");
var lowDivNode = document.createElement("div");
var slNode = document.createElement("div");
slNode.id = "myExcelDiv";

var tfHeaderNode = document.createElement("th");
var percentHeaderNode = document.createElement("th");
var investedHeaderNode = document.createElement("th");
var slHeaderNode = document.createElement("th");

//var ExcelData;
var SLData = {};

//onload = priceHL();

//window.addEventListener('load', priceHL());


window.onload = function () {
    
    console.log("Access Token Starting");
    
    // Step 5: Initiate token retrieval
    getAccessTokenFromUrl(); // Start the process
    
    //priceHL(); //Make sure the function fires as soon as the page is loaded
    //setTimeout(priceHL, 10000); //Then set it to run again after ten minutes
    
};

function addKeyValue(jsonObj, key, slValue, timeFrame) {
  var value = [];
  value.push(slValue);
  value.push(timeFrame);
  
  jsonObj[key] = value; // Adds a key-value pair to the JSON object
  
  return jsonObj;
}

function priceHL(ExcelData){
  var len;
  
  var high = 0;
  var low = 0;
  var highStock = "";
  var lowStock = "";
  
  var holdingPage = document.getElementsByClassName("data-table fold-header sticky")[0];
  var holdingTableWrap = holdingPage.childNodes[4];
  var holdingTable = holdingTableWrap.childNodes[0];
  var holdingTableBody = holdingTable.childNodes[2];
  
  len = holdingTableBody.children.length-1;
  
  //var totalInvestElement = document.getElementsByClassName("value text-bold text-pagetitle")[0].textContent;
  var totalInvestElement = document.getElementsByClassName("value")[0].textContent;
  var totalInvestment = parseFloat(totalInvestElement.replaceAll(",",""));
  
  //len = (document.getElementsByClassName("data-table fold-header sticky")[0].childNodes[4].childNodes[0].childNodes[2].children.length)-1;
  
  //console.log("len = " + len);
  
  var stock;
  var dayChange;
  
  
  for(var j=0; j<ExcelData.length;j++){
    var st = ExcelData[j].values[0][2];  // stock column in Excel sheet
    var sl = ExcelData[j].values[0][3];  // SL/TSL column in Excel sheet
    var tf = ExcelData[j].values[0][16]; // setup column in Excel sheet
    
    //console.log("st = " + ExcelData[j].values[0][2] + ", SL = " + ExcelData[j].values[0][3]);
    
    //SLData[st] = sl;
    SLData = addKeyValue(SLData, st, sl, tf);
  }
  
  console.log("SL DATA : " + JSON.stringify(SLData));
    
  
  for(var i=0; i<= len*2; i++){
    
    var tfNode = document.createElement("td");
    var percentNode = document.createElement("td");
    var investedNode = document.createElement("td");
    var slNode = document.createElement("td");
    
    if(i%2 == 0){
      
      var holdingRow = holdingTableBody.childNodes[i];
      
      var holdingRowEleInstrument = holdingRow.childNodes[2];
      var holdingRowEleInstrumentRef = holdingRowEleInstrument.childNodes[0];
      var holdingRowEleSpan = holdingRowEleInstrumentRef.childNodes[0];
      
      var holdingRowEleDayChange = holdingRow.childNodes[18];
      var holdingRowEleDayChangeRef = holdingRowEleDayChange.childNodes[0];
      
      var holdingPriceEle = holdingRow.childNodes[6];
      var holdingPrice = parseFloat(holdingPriceEle.innerText.replaceAll(",","")).toFixed(2);
      
      var ltpPriceEle = holdingRow.childNodes[8];
      var ltpPrice = parseFloat(ltpPriceEle.innerText.replaceAll(",","")).toFixed(2);
      
      var qtyEle = holdingRow.children[1];
      
      stock = holdingRowEleSpan.textContent;
      //dayChange = holdingRowEleDayChangeRef.dataset.balloon;
      dayChange = holdingRowEleDayChangeRef.textContent.replaceAll("\n","").replaceAll("\t","");
      
      //var holdingCurrent  = holdingRow.childNodes[12];
      //CurrentVal = parseFloat(holdingCurrent.textContent.replaceAll("\n","").replaceAll("\t","").replace(",",""));
      
      //var holdingPL  = holdingRow.childNodes[14];
      // = parseFloat(holdingPL.textContent.replaceAll("\n","").replaceAll("\t","").replace(",",""));
      
      //var InvestedVal = (CurrentVal - PnL).toFixed(2);
      var holdingInvested = holdingRow.childNodes[10];
      InvestedVal = parseFloat(holdingInvested.textContent.replaceAll("\n","").replaceAll("\t","").replace(",",""));
      var InvestedPercent = ((InvestedVal/totalInvestment)*100).toFixed(2);
      
      
      //Change Stock Name to stock:Invested value
      //holdingRowEleSpan.innerHTML = stock + " : " + InvestedVal;//  + " : " + InvestedPercent + "%";
      
      //var chngArr = dayChange.split(" ");
      //var arr1Value = chngArr[1].replace("(","").replace(")","").replace("%","").replace("+","");
      var arr1Value = dayChange.replace("(","").replace(")","").replace("%","").replace("+","");
      
      if(parseFloat(arr1Value) > high){
        high = arr1Value;
        highStock = stock;
      }
      
      //console.log("Day Low = " + lowStock + " % = " + low + ", arr1Value = " + arr1Value);
      //console.log("Type arr1Value = " + typeof arr1Value);

      if(parseFloat(arr1Value) < low){
        low = arr1Value;
        lowStock = stock;
      }
      
      
      if(stock in SLData){
	      var slValue = SLData[stock][0];
	      switch(SLData[stock][1].toString()){
	        case "DAILY":
	          slValue = (SLData[stock][0] * 0.995).toFixed(2);
	        case "WEEKLY":
	          slValue = (SLData[stock][0] * 0.99).toFixed(2);
	        case "MONTHLY":
	          slValue = (SLData[stock][0] * 0.98).toFixed(2);
	      }
      }
      else slValue = 0;
      
      if(stock in SLData){
	      tfNode.innerHTML = SLData[stock][1].toString();
      }
      else{
      	tfNode.innerHTML = "NONE";
      }
      holdingRow.insertBefore(tfNode, qtyEle);
      
      if(stock in SLData){
      	slNode.innerHTML = slValue;
      }
      else{
      	slNode.innerHTML = 00;
      }
      holdingRow.insertBefore(slNode, qtyEle);
      
      if(stock in SLData){
      	percentNode.innerHTML = InvestedPercent.toString() + "%";
      }
      else{
      	percentNode.innerHTML = " ";
      }
      holdingRow.insertBefore(percentNode, qtyEle);
      
      
      /*if(stock in SLData){
      	investedNode.innerHTML = InvestedVal;
      }
      else{
      	investedNode.innerHTML = " ";
      }
      holdingRow.insertBefore(investedNode, qtyEle);
      */
      
      /*
      if(SLData[stock] !== undefined && parseFloat(holdingPrice) < slValue){
        console.log(stock + " holdingPrice = "+holdingPrice+",SL = "+SLData[stock][0] +",slValue = " + slValue);
        */
      if(SLData[stock] !== undefined && parseFloat(ltpPrice) < slValue){
        console.log(stock + " ltpPrice = "+ltpPrice+",SL = "+SLData[stock][0] +",slValue = " + slValue);
        holdingRow.className = "sl-row";
        
        // Create a MutationObserver specific to this row
        (function(row) {
            const observer = new MutationObserver(function(mutationsList) {
                mutationsList.forEach(function(mutation) {
                    if (mutation.attributeName === "class" && !row.classList.contains("sl-row")) {
                        row.classList.add("sl-row"); // Reapply the class if removed
                    }
                });
            });

            // Start observing changes to this row's class attribute
            observer.observe(row, { attributes: true });
        })(holdingRow);  // Pass the current row into the IIFE
      }
      
      //console.log("Day High = " + highStock + " % = " + high);
    }//if i%2 end
    
    
  }//for end
  
  var holdingTableHeader = holdingTable.childNodes[0];
  var holdingHeaderRow = holdingTableHeader.childNodes[0];
  var qtyHeader = holdingHeaderRow.children[1];
  
  tfHeaderNode.innerHTML = "Setup";
  slHeaderNode.innerHTML = "SL/TSL";
  percentHeaderNode.innerHTML = "Inv %";
  //investedHeaderNode.innerHTML = "Inv.Val";
  
  
  holdingHeaderRow.insertBefore(tfHeaderNode, qtyHeader);
  holdingHeaderRow.insertBefore(slHeaderNode, qtyHeader);
  holdingHeaderRow.insertBefore(percentHeaderNode, qtyHeader);
  //holdingHeaderRow.insertBefore(investedHeaderNode, qtyHeader);
  

  node.appendChild(highDivNode);
  node.appendChild(lowDivNode);
  
  node.classList.add("form-bhushan", "grid-container");
  highDivNode.classList.add("high-node","text-green", "text-bold", "text-pagetitle", "grid-item");
  lowDivNode.classList.add("low-node","text-red", "text-bold", "text-pagetitle", "grid-item");
  
  //Event Listener to make node draggable
  let isDragging = false;
  let startX, startY;
  
  
  node.addEventListener("mousedown", (event) => {
    isDragging = true;
    startX = event.clientX - node.offsetLeft;
    startY = event.clientY - node.offsetTop;
  });
  document.addEventListener("mousemove", (event) => {
    if (!isDragging) return;
  
    const x = event.clientX - startX;
    const y = event.clientY - startY;
  
    node.style.left = x + "px";
    node.style.top = y + "px";
  });
  
  document.addEventListener("mouseup", (event) => {
    isDragging = false;
  });
  
  document.getElementsByClassName("app page-holdingsEq")[0].appendChild(node);
  
  highDivNode.innerHTML = highStock + ": " + high + "%";
  lowDivNode.innerHTML = lowStock + ": " + low + "%";
  
}


// Add scroll event listener to make the element sticky
document.addEventListener('scroll', function(){
  
  if ( document.body.scrollTop >= 1 || document.scrollingElement.scrollTop >= 1 || document.documentElement.scrollTop >= 1 ){
     
      node.classList.add('custom-fixed');
  } else {
      
      node.classList.remove('custom-fixed')
  }
  
});



const clientId = '9d1172e4-c600-437e-924c-588635bd2e21'; // Your application client ID
const redirectUri = 'https://kite.zerodha.com/holdings/equity'; // Change to your redirect URI
const authUrl = `https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent('https://graph.microsoft.com/.default')}&prompt=select_account`;

let accessToken = null;

// Step 1: Initiate login if no access token is present
function initiateLogin() {
    window.location.href = authUrl;
}

// Step 2: Capture the access token from the URL
function getAccessTokenFromUrl() {
    const hash = window.location.hash;
    if (hash) {
        const params = new URLSearchParams(hash.replace('#', '?'));
        accessToken = params.get('access_token');
        if (accessToken) {
            console.log('Access Token:', accessToken);
            fetchExcelData(accessToken);
            
            scheduleTokenRefresh();
        } else {
            console.error('Access token not found in URL.');
            initiateLogin(); // Retry login if token not found
        }
    } else {
        initiateLogin(); // Retry login if no hash found
    }
}

// Step 3: Schedule token refresh
function scheduleTokenRefresh() {
    // Refresh 5 minutes before expiration (assuming token expires in 1 hour)
    setTimeout(() => {
        console.log("Access token expired, reloading page to refresh.");
        initiateLogin();
    }, 55 * 60 * 1000);
}

// Step 4: Fetch Excel table data
async function fetchExcelData(token) {
    const filePath = 'Documents/.MyHoldings.xlsx'; // Change to your file path
    const tableName = 'TradeDataTable'; // Change to your table name

    const apiUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/rows`;

    const response = await fetch(apiUrl, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
        },
    });

    if (!response.ok) {
        const errorResponse = await response.json();
        console.error('Failed to fetch Excel data:', errorResponse);

        // If unauthorized, try to refresh the token by reloading
        if (response.status === 401) {
            console.log("Unauthorized. Reloading to obtain new token.");
            initiateLogin();
        }
        return;
    }

    const data = await response.json();
    console.log('Excel Data:', data["value"][1]["values"]);
    ExcelData = data["value"];
    console.log('Excel Data fetchExcelData:', ExcelData.length);
    
    console.log('starting priceHL:', ExcelData.length);
    priceHL(data["value"]);
}

// Reload the page every 5 minutes (300000 milliseconds)
//setInterval(function() {
    //location.reload(); // Refreshes the page
//}, 300000); // 5 minutes
