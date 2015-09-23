/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */

/// <reference path="../App.js" />

(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			populateWorkbookWithRegionalSalesData();

			$('#consolidate-sheets').click(consolidateSheetsData);
		});
	};

	// Populates the workbook with some sales data in three new sheets
	function populateWorkbookWithRegionalSalesData() {

		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {

			// Create a proxy object for the worksheets collection
			var sheets = ctx.workbook.worksheets;

			// Queue commands to add three new worksheets to the workbook
			var sheet1 = sheets.add("Americas");
			var sheet2 = sheets.add("Asia");
			var sheet3 = sheets.add("Europe");

			// Add sample Quarterly Sales data to each sheet

			// Sheet 1

			// Queue commands to set the title in the sheet and format it
			sheet1.getRange("A1:A1").values = "Quarterly Sales Report - Americas";
			sheet1.getRange("A1:A1").format.font.name = "Century";
			sheet1.getRange("A1:A1").format.font.size = 26;


			// Sample data for Sheet1
			var sheet1values = [["Product", "January", "February", "March"],
						  ["Frames", 1, 5, 10],
						  ["Saddles", 400, 323, 276],
						  ["Brake levers", 12000, 8766, 8456],
						  ["Chains", 1550, 1088, 692],
						  ["Mirrors", 225, 600, 923],
						  ["Spokes", 6005, 7634, 4589]];

			// Queue a command to write the sample data to Sheet1
			sheet1.getRange("A2:D8").values = sheet1values;

			// Queue a command to format the header row of the range
			sheet1.getRange("A2:D2").format.font.bold = true;

			// Sheet2

			// Queue commands to set the title in the sheet and format it
			sheet2.getRange("A1:A1").values = "Quarterly Sales Report - Asia";
			sheet2.getRange("A1:A1").format.font.name = "Century";
			sheet2.getRange("A1:A1").format.font.size = 26;


			// Sample data for Sheet2
			var sheet2values = [["Product", "January", "February", "March"],
						  ["Frames", 2, 2, 2],
						  ["Saddles", 678, 678, 456],
						  ["Brake levers", 9989, 6789, 2467],
						  ["Chains", 450, 345, 567],
						  ["Mirrors", 345, 456, 234],
						  ["Spokes", 3456, 4567, 8763]];

			// Queue a command to write the sample data to Sheet1   
			sheet2.getRange("A2:D8").values = sheet2values;

			// Queue a command to format the header row of the range
			sheet2.getRange("A2:D2").format.font.bold = true;

			// Sheet3

			// Queue commands to set the title in the sheet and format it
			sheet3.getRange("A1:A1").values = "Quarterly Sales Report - Europe";
			sheet3.getRange("A1:A1").format.font.name = "Century";
			sheet3.getRange("A1:A1").format.font.size = 26;


			// Sample data for Sheet1
			var sheet3values = [["Product", "January", "February", "March"],
						  ["Frames", 3, 3, 3],
						  ["Saddles", 776, 647, 456],
						  ["Brake levers", 4567, 3455, 4567],
						  ["Chains", 454, 786, 768],
						  ["Mirrors", 123, 766, 453],
						  ["Spokes", 6789, 7889, 7673]];

			// Queue a command to write the sample data to Sheet1
			sheet3.getRange("A2:D8").values = sheet3values;

			// Queue a command to format the header row of the range
			sheet3.getRange("A2:D2").format.font.bold = true;

			// Queue a command to activate the Americas sheet
			sheet1.activate();

			// Done adding sample data
			// Next add checkboxes in the task pane for each sheet in the workbook

			// We need to load properties prior to reading them
			// Queue a command to load the name property of each worksheet in the sheets collection
			// We need this for the checkbox names
			sheets.load("name");

			// Run all of the above queued-up commands, and return a promise to indicate task completion
			return ctx.sync().then(function () {

				// For each sheet, add a checkbox in the UI
				for (var i = 0; i < sheets.items.length; i++) {
					// The name property has been loaded already
					var checkboxName = sheets.items[i].name;
					var $input = $('<input type="checkbox" class="sheet" value=' + checkboxName + ' id=' + checkboxName + ' name="check"><label for=' + checkboxName + '>' + checkboxName + '</label><br />');
					$input.appendTo($("#checkbox-set"));
				}
			})
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}

	// Consolidate data from sheets into a new dashboard sheet
	function consolidateSheetsData() {

		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {

			// Create a proxy object for the worksheets collection
			var sheets = ctx.workbook.worksheets;

			// Queue a command to add a new sheet called Dashboard
			var dashboardSheet = sheets.add("Dashboard");

			// Start consolidating dat
			var selectedSheetNamesArray = [];
			var verticalLabels = [];
			var horizontalLabels = [];
			var dashboardHorizontalLabelsRange;
			var dataArray = [];


			// Queue commands to set the title in the sheet and format it
			dashboardSheet.getRange("A1:A1").values = "Consolidated Quarterly Sales Report";
			dashboardSheet.getRange("A1:A1").format.font.name = "Century";
			dashboardSheet.getRange("A1:A1").format.font.size = 26;

			// Get the names of the sheets the user chose in the task pane
			selectedSheetNamesArray = $('input:checkbox:checked.sheet').map(function () {
				return this.value;
			}).get();

			// Queue commands to load the values and address of the first sheet's used range
			var firstUsedRange = sheets.getItem(selectedSheetNamesArray[0]).getUsedRange();
			firstUsedRange.load("address, values");

			// Run the queued-up commands
			return ctx.sync()
			   .then(function () {
				   // Start copying data from the selected sheets into the Dashboard sheet

				   // First queue commands to copy over the vertical labels:
				   for (var row = 1; row < firstUsedRange.values.length; row++) {
					   verticalLabels.push([firstUsedRange.values[row][0]]);
				   }
				   var dashboardTopLeftCell = dashboardSheet.getRange("A2");
				   var dashboardVerticalLabelsRange = dashboardTopLeftCell.getBoundingRect(
					   dashboardTopLeftCell.getOffsetRange(verticalLabels.length - 1, 0));
				   // Here's a useful debugging/visual trick to make sure you've the right range
				   // dashboardVerticalLabelsRange.format.fill.color = "red";
				   dashboardVerticalLabelsRange.values = verticalLabels;


				   // Queue commands to copy over horizontal labels
				   for (var column = 1; column < firstUsedRange.values[1].length; column++) {
					   horizontalLabels.push(firstUsedRange.values[1][column]);
				   }
				   var dashboardTopLeftCell = dashboardSheet.getRange("B2");
				   dashboardHorizontalLabelsRange = dashboardTopLeftCell.getBoundingRect(
				   dashboardTopLeftCell.getOffsetRange(0, horizontalLabels.length - 1));
				   dashboardHorizontalLabelsRange.values = [horizontalLabels];

				   // Queue commands to copy data from all the sheets using formula
				   var usedRangeFullAddress = firstUsedRange.address;
				   var toRemove = selectedSheetNamesArray[0] + "!";
				   var usedRangeAddress = usedRangeFullAddress.replace(toRemove, '');
				   var dashboardEntireRange = dashboardSheet.getRange(usedRangeAddress);
				   var dashboardDataRange = dashboardEntireRange.getCell(2, 1).getBoundingRect(dashboardEntireRange.getLastCell());
				   for (var row = 1; row < firstUsedRange.values.length - 1 ; row++) {
					   var temp = [];
					   for (var col = 0; col < firstUsedRange.values[0].length - 1; col++) {
						   var components = [];
						   selectedSheetNamesArray.forEach(function (item) {
							   var string = "'" + item + "'!" + columnName(col + 1) + (row + 2);
							   components.push(string);
						   });
						   temp.push("=" + components.join("+"));
					   }
					   dataArray.push(temp);
				   }
				   dashboardDataRange.values = dataArray;

				   // Queue a command to activate the Dashboard sheet
				   dashboardSheet.activate();
			   })
				.then(ctx.sync)
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}

	/**
	  * Returns the column name based on a zero-based column index.
	  * For example, columnName(4) = 5th column = "E". Meanwhile, columnName(1000) = 1001st column = "ALM".
	  * @param index Zero-based column index.
	  * @returns {String} Locale-independent column name (e.g., a string comprised of one or more letters in the range "A:Z").
	  */
	function columnName(index) {
		if (typeof index !== 'number' || isNaN(index) || index < 0) {
			throw new OfficeExtension.RuntimeError("InvalidArgument", "The parameter for Excel.columnName(x) must be positive and numeric.", [], { errorLocation: "Excel.Util.columnName" });
		}
		var letters = [];
		while (index >= 0) {
			letters.push(getSingleLetter(index % 26));
			index = Math.floor(index / 26) - 1;
		}
		return letters.reverse().join('');
		function getSingleLetter(zeroThrough25Index) {
			return String.fromCharCode(zeroThrough25Index + 65); // ASCII code for "A" is 65
		}
	}

})();
/*
Excel-Add-in-JS-ConsolidatedSalesReport, https://github.com/OfficeDev/Excel-Add-in-JS-ConsolidatedSalesReport

Copyright (c) Microsoft Corporation

All rights reserved. 

MIT License:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and 
associated documentation files (the "Software"), to deal in the Software without restriction, including 
without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the 
following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial 
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT 
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT 
SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN 
ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE 
USE OR OTHER DEALINGS IN THE SOFTWARE. 
*/
