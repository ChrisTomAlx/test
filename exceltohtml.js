if (typeof require !== 'undefined') XLSX = require('sheetchart');

module.exports = function (RED) {

	
	function isSheetIndicesValid(msg) {
		if (typeof (msg.payload.sheetIndices) != 'undefined' && Array.isArray(msg.payload.sheetIndices) && msg.payload.sheetIndices.length != 0) // return html of only specified sheets
		{
			return true;
		}
		else {
			return false;
		}
	}


	function getSheetFromHtml(wb, i) {
		var sheetToHtml;
		try {
			ws = wb.Sheets[wb.SheetNames[i]];
			sheetToHtml = XLSX.utils.sheet_to_html(ws);
		}
		catch (error) {
			sheetToHtml = "";
		}
		return sheetToHtml;
	}


	function exceltohtml(config) {
		RED.nodes.createNode(this, config);
		this.propDataType = config.propDataType;
		this.propSheetData = config.propSheetData;
		var node = this;
		node.on('input', function (msg) {

			var wb;
			var isError = false;
			try {
				// Check if mandatory inputs dataType and sheetData has been passed in
				if ((typeof (msg.payload.dataType) === 'string' || typeof (node.propDataType) === 'string') && (typeof (msg.payload.sheetData) != 'undefined' || typeof (node.propSheetData) != 'undefined')) {
					
					var datatypelocal = (typeof(msg.payload.dataType) === 'string') ? msg.payload.dataType : node.propDataType;
					var sheetdatalocal = (typeof(msg.payload.sheetData) != 'undefined') ? msg.payload.sheetData : node.propSheetData;

					// sheet Option to remove cell styling of every cell
					var keepCellStyles = msg.payload.keepCellStyles === false ? msg.payload.keepCellStyles : true;
					
					// sheet Option to remove cell styling of empty cells
					var keepEmptyCellStyles = msg.payload.keepEmptyCellStyles === false ? msg.payload.keepEmptyCellStyles : true;

					// file
					if (datatypelocal.toLowerCase().trim() == "file" && typeof sheetdatalocal === 'string')  
					{
						try
						{
							wb = XLSX.readFile(sheetdatalocal, { cellStyles: keepCellStyles, sheetStubs: keepEmptyCellStyles });	
						}
						catch (e)
						{
							if (!isError) 
							{
								msg.payload = {status: 400,message: "Error 400 : Possibly wrong sheetData specified." };
								node.error(msg);
								isError = true;
							}
						}	
					}
					// arraybuffer
					else if (datatypelocal.toLowerCase().trim() == "arraybuffer" && Buffer.isBuffer(sheetdatalocal)) 
					{	
						try
						{
							wb = XLSX.read(sheetdatalocal, { type: "array", cellStyles: keepCellStyles, sheetStubs: keepEmptyCellStyles });
						}
						catch (e)
						{
							if (!isError) 
							{
								msg.payload = {status: 400,message: "Error 400 : Possibly wrong sheetData specified." };
								node.error(msg);
								isError = true;
							}
						}
					}
					// buffer
					else if (datatypelocal.toLowerCase().trim() == "buffer" && Buffer.isBuffer(sheetdatalocal)) 
					{	
						try
						{
							wb = XLSX.read(sheetdatalocal, { type: "buffer", cellStyles: keepCellStyles, sheetStubs: keepEmptyCellStyles });
						}
						catch (e)
						{
							if (!isError) 
							{
								msg.payload = {status: 400,message: "Error 400 : Possibly wrong sheetData specified." };
								node.error(msg);
								isError = true;
							}
						}
					}
					// Neither a proper file or a buffer
					else 
					{
						if (!isError) 
						{
							msg.payload = {status: 400,message: "Error 400 : Wrong dataType/sheetData specified." };
							node.error(msg);
							isError = true;
						}
					}
				}
				// type of file or data was not specified by user
				else 
				{
					if (!isError) 
					{
						msg.payload = { status: 400, message: "Error 404 : No dataType/sheetData specified." };
						node.error(msg);
						isError = true;
					}
				}

				// Outputs
				var allSheetNames = [];
				var reqSheetNames = [];
				var excelAsHtml = [];

				// Styling variables
				var bodyDivStyle = ".bodydiv {font-family:Arial; max-width:100%; max-height:100%; overflow: auto;}";
				var tabScrollBarStyle = ".tab::-webkit-scrollbar { height: 0px; } .tab::-webkit-scrollbar-track { -webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.3); border-radius: 10px; } .tab::-webkit-scrollbar-thumb { border-radius: 10px; -webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.5); }"
				var tabStyleOnHover = "";	//".tab:hover {overflow: auto}" 
				var tabStyle = ".tab { scrollbar-width: none;overflow: auto; border: 1px solid #ccc; background-color: #f1f1f1; }";
				var buttonDivPos = "sticky;position:-webkit-sticky;float:right";

				if(msg.payload.noParent === true){
					bodyDivStyle = ".bodydiv {font-family:Arial; margin-bottom:100px;}";
					buttonDivPos = "fixed";
				}


				// Style, Div and Script for completely formatted HTML output
				var style = "<style> "+bodyDivStyle+" "+tabScrollBarStyle+" "+tabStyleOnHover+" "+tabStyle+" /* Style the buttons inside the tab */ .tab button { background-color: inherit; float: left; border: none; outline: none; cursor: pointer; padding: 14px 16px; transition: 0.3s; font-size: 17px; } /* Change background color of buttons on hover */ .tab button:hover { background-color: #ddd; } /* Create an active/current tablink class */ .tab button.active { background-color: #ccc; } /* Style the tab content */ .tabcontent { display: none; padding: 6px 12px;-webkit-animation: fadeEffect 1s; animation: fadeEffect 1s; border: 1px solid #ccc; border-top: none; } /* Fade in tabs */ @-webkit-keyframes fadeEffect { from {opacity: 0;} to {opacity: 1;} } @keyframes fadeEffect { from {opacity: 0;} to {opacity: 1;} } </style>";
				
				var buttondivarrow = "<div style='flex-direction: row;display: flex;'><button style='width: 45px;' onclick='shift(-1)'>❮</button> <button style='width: 45px;' onclick='shift(+1)'>❯</button></div>";
				var buttondiv = "<div style='width:100%;overflow: auto;position:"+buttonDivPos+";left:0;bottom: 0;flex-direction: row;display: inline-flex;'>"+buttondivarrow+"<div id ='btncontainer' class='tab' style = 'display: inline-flex; white-space: nowrap;width:100%;'>";
				
				//var buttondiv = "<div id='btncontainer' class='tab' style = 'position:"+buttonDivPos+";left:0;bottom: 0;display: inline-flex; white-space: nowrap;width:100%;'>";
				var script = 		  "<script>";
					script = script + "function openSheet(evt, sheetName) {";
					script = script + "var i, tabcontent, tablinks;";
					script = script + "tabcontent = document.getElementsByClassName('tabcontent');";
					script = script + "for (i = 0; i < tabcontent.length; i++) {";
					script = script + "tabcontent[i].style.display = 'none';}";

					script = script + "tablinks = document.getElementsByClassName('tablinks');";
					script = script + "for (i = 0; i < tablinks.length; i++) {tablinks[i].className = tablinks[i].className.replace(' active', '');}";
					script = script + "document.getElementById(sheetName).style.display = 'block';";
					script = script + "evt.currentTarget.className += ' active';}";

					script = script + "document.getElementById('defaultOpen').click();";

					// function to scroll on button click
					script = script + "function shift(direction) {";
					script = script + "var container = document.getElementById('btncontainer');";
					script = script + "var scrollsize = document.getElementsByClassName('tablinks')[0].scrollWidth;";
					script = script + "var start = container.scrollLeft;";
					script = script + "var change = scrollsize*9/10; var duration = 175;";
					script = script + "currentTime = 0; increment = 25;";

					script = script + "var animateScroll = function(){ currentTime += increment;var val;";
					script = script + "if (direction == 1) {val = start + (1-(duration-currentTime)/duration)*change;}";
					script = script + "else {val = start - (1-(duration-currentTime)/duration)*change;}";
					script = script + "container.scrollLeft = val;";
					script = script + "if(currentTime < duration) { setTimeout(animateScroll, increment); } };";
					script = script + "animateScroll(); }";
					

					script = script + "</script>";

				// Allow user to edit styles and scripts
				if(typeof msg.payload.customHtmlStyle === 'string') {
					style = msg.payload.customHtmlStyle;
				}
				
				if(typeof msg.payload.customHtmlScript === 'string') {
					script = msg.payload.customHtmlScript;
				}					


				// Get all sheet names
				for (var i = 0; i < wb.SheetNames.length; i++) {
					allSheetNames.push(wb.SheetNames[i]);
				}


				// If user has requested for only specific sheets
				if (isSheetIndicesValid(msg)) {
					// Cycle through req sheet indixes
					for (var i = 0; i < msg.payload.sheetIndices.length; i++) 
					{
						reqSheetNames.push(wb.SheetNames[msg.payload.sheetIndices[i]]);

						if (msg.payload.sheetIndices[i] < wb.SheetNames.length && Number.isInteger(msg.payload.sheetIndices[i]) && msg.payload.sheetIndices.length <= wb.SheetNames.length) // sheet index validated
						{
							if (msg.payload.retPure === true) {
								excelAsHtml.push(getSheetFromHtml(wb, msg.payload.sheetIndices[i]));
							}
							else {
								if (typeof (excelAsHtml[0]) == 'undefined') {
									excelAsHtml[0] = "<html>" + style + "<body><div class='bodydiv'>";
									buttondiv = buttondiv + " <button class='tablinks' onclick=\"openSheet(event, 'a" + i + "')\" id='defaultOpen'>" + reqSheetNames[i] + "</button>";
								}
								else { buttondiv = buttondiv + " <button class='tablinks' onclick=\"openSheet(event, 'a" + i + "')\">" + reqSheetNames[i] + "</button>"; }

								excelAsHtml[0] = excelAsHtml[0] + "<div id='a" + i + "' class='tabcontent'>" + getSheetFromHtml(wb, msg.payload.sheetIndices[i]) + "</div>";

								if (i == msg.payload.sheetIndices.length - 1) { excelAsHtml[0] = excelAsHtml[0] + buttondiv + "</div></div>" + script + "</body></html>"; }
							}
						}
						else {
							if (!isError) 
							{
								msg.payload = { status: 400, message: "Inputed Sheet index is wrong" };
								node.error(msg);
								isError = true;
							}
						}
					}
				}
				// If sheet indices not available provide HTML of all sheets
				else 
				{
					// Cycle through all sheets
					for (var i = 0; i < wb.SheetNames.length; i++) 
					{
						if (msg.payload.retPure === true) {
							excelAsHtml.push(getSheetFromHtml(wb, i));
						}
						else 
						{
							if (typeof (excelAsHtml[0]) == 'undefined') {
								excelAsHtml[0] = "<html>" + style + "<body><div class='bodydiv'>";
								buttondiv = buttondiv + " <button class='tablinks' onclick=\"openSheet(event, 'a" + i + "')\" id='defaultOpen'>" + allSheetNames[i] + "</button>";
							}
							else { buttondiv = buttondiv + " <button class='tablinks' onclick=\"openSheet(event, 'a" + i + "')\">" + allSheetNames[i] + "</button>"; }

							excelAsHtml[0] = excelAsHtml[0] + "<div id='a" + i + "' class='tabcontent'>" + getSheetFromHtml(wb, i) + "</div>";

							if (i == wb.SheetNames.length - 1) { excelAsHtml[0] = excelAsHtml[0] + buttondiv + "</div></div>" + script + "</body></html>"; }
						}
					}
				}
			}
			catch (e) {
				if (!isError) {
					msg.payload = { status: 500, error: e, message: "Oops something went wrong! Possibly caused by the Script" };
					node.error(msg);
					isError = true;
				}
			}

			if (!isError) {
				var retPure = msg.payload.retPure;
				msg.payload = {};
				msg.payload.allSheetNames = allSheetNames;
				msg.payload.reqSheetNames = reqSheetNames;
				if (retPure === true) {
					msg.payload.excelAsHtml = excelAsHtml;
				} else {
					msg.payload.excelAsHtml = excelAsHtml[0];
				}
			}


			node.send(msg);

		});
	}
	RED.nodes.registerType("exceltohtml", exceltohtml);
}
