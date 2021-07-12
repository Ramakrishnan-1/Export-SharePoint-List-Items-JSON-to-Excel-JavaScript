$(document).ready(function() {
	$("#pageTitle").hide();

	fnGetListNames();//Bind custom lists to radio buttons 

	$("#exportListBtn").on("click", function() {
        var selectedVal = $("input[type='radio']:checked").val();
		if(selectedVal){
            $("#exportListBtn").hide();
            $("#alertMsg").text('File Ready To Download!');
            $("#alert").css("display", "block");
            fnExport2Excel(selectedVal);
        } else{
            alert('Select a list to export.!');
        }
	});
	
});

function fnGetListNames() {
	var htmlContent = '';
	var listNamesID = document.getElementById("listNamesID");
	$.ajax({
		url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/",
		type: "GET",
		headers: {
			"Accept": "application/json;odata=verbose",
		},
		success: function(data, status, xhr) {
			var dataResults = data.d.results;
			for(var i = 0; i < dataResults.length; i++) {
				if(dataResults[i].BaseTemplate === 100 && dataResults[i].Hidden === false) {
					var listTitle = dataResults[i].Title;
					htmlContent += '<tr>' + '<td><input type="radio" id="ValueID' + i + '" name="listNames" value="' + listTitle + '"></td><td>' + listTitle + '</td></tr>';
				}
			}
			listNamesID.innerHTML = htmlContent;
		},
		error: function(xhr, status, error) {
			console.log("Failed to Get List Names");
		}
	});
}

function fnExport2Excel(selectedValue) {
	$.ajax({
		url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getByTitle('"+selectedValue+"')/items?",
		type: "GET",
		headers: {
			"Accept": "application/json;odata=verbose"
		},
		success: function(data, status, xhr) {
			var listData = data.d.results;
			var fileName = selectedValue+'.xlsx';
            var jsonToSheet = XLSX.utils.json_to_sheet(listData);
            var excelBook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(excelBook, jsonToSheet, selectedValue);
            XLSX.writeFile(excelBook, fileName);
		},
		error: function(xhr, status, error) {
			console.log("Failed to download");
		}
	});
}