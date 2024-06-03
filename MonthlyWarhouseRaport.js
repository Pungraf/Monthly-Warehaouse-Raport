function MonthlyRaport() 
{
	//Initialize sheets and ranges
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var raportSheet = ss.getSheetByName("Raport");
	var raportRange = raportSheet.getRange("A1:AF110");
	var ssUsedPowder = SpreadsheetApp.openByUrl('Link to external sheet');
	var usedPowderSheet = ssUsedPowder.getSheetByName('Aktualne_zużycie i odzysk');
	var usedPowderRange = usedPowderSheet.getRange("A2:K900");
	var rngPowderUsed = usedPowderRange.getValues();
	var rngRaport = raportRange.getValues();
	var allsheets = ss.getSheets();

	//------------------------------------
	//Initialize Raport Data
	var month = 0;
	var prMade = 0;
	var pr11Made = 0;
	var prUsedToPrint = 0;
	var prAsStarter = 0;
	var spMade = 0;
	var spPacked = 0;
	var frPa12UsedToRefresh = 0;
	var frPa12Packed = 0;
	var frPa11Packed = 0;

	var industrialMade = 0;
	var iundustrialUsedToPrint = 0;
	var frIndustrialUsedToRefresh = 0;
	var frIndustrialPacked = 0;

	var pbtMade = 0;
	var pbtUsedToPrint = 0;
	var frPBTUsedToRefresh = 0;
	var frPBTPacked = 0;

	var frPa12_2kg = 0;
	var frPa12_6kg = 0;
	var frIndustrial_6kg = 0; 
	var frPBT_6kg = 0;
	var spPa12_2kg = 0;
	var spPa12_4kg = 0;
	var spPa12_6kg = 0;
	var frPa11_2kg = 0;
	var frPa11_6kg = 0;
	var spPa11_2kg = 0;
	var spPa11_4kg = 0;
	var spPa11_6kg = 0;
	var prPa11UsedToPrint = 0;
	var pa11ESD_2kg = 0;
	var pa11CF_6kg = 0;
	var pp_6kg = 0;
	var flexaGrey_2kg = 0;
	var flexaGrey_6kg = 0;
	var flexaBlac_2kg = 0;
	var flexaBlack_6kg = 0;
	var flexaBright_2kg = 0;
	var flexaPerformance_6kg = 0;
	var tpe_2kg = 0;
	var flexaSoft = 0;
	var sealer = 0;

	var frPa12Made = 0;
	var frPBTMade = 0;
	var frIndustrialMade = 0;
	var frPa11Made = 0;
	var flexaGreyMade = 0;
	var flexaBlackMade = 0;
	var flexaBrightMade = 0;
	var flexaPerformanceMade = 0;
	var ppMade = 0;
	var pa11ESDMade = 0;
	var pa11CFMade = 0;
	//------------------------------------
	//Variables
	var dateColumn = 0;
	//Powder Type
	var pa12Column = 3;
	var industrialColumn = 4;
	var pa11Column = 5;
	var pa11ESDColumn = 6;
	var pa11CFColumn = 7;
	var ppColumn = 8;
	var flexaGreyColumn = 9;
	var flexaBlackColumn = 10;
	var flexaPerformanceColumn = 11;
	var flexaSoftColumn = 12;
	var flexaBrightColumn = 13;
	var tpeColumn = 14;
	var pbtColumn = 15;
	var sealerColumn = 16;

	//Product Type
	var fresh_Offset = 0;
	var prSum_Offset = 2;
	var prFresh_Offset = 3;
	var prUsed_Offset = 4;
	var spSum_Offset = 5;
	var spFresh_Offset = 6;
	var spUsed_Offset = 7;
	var fr2kg_Offset = 11;
	var fr6kg_Offset = 12;
	var sp2kg_Offset = 13;
	var sp4kg_Offset = 14;
	var sp6kg_Offset = 15;
	var bottle_Offset = 16;
	var sampleRefresh_Offset =18;

	//Output Cells
	var month_Output = 1;
	var prMade_Output = 2;
	var pr11Made_Output = 3;
	var prUsedToPrint_Output = 4;
	var spMade_Output = 5;
	var spPacked_Output = 6;
	var frPa12UsedToRefresh_Output = 7;
	var frPa12Packed_Output = 8;
	var frPa11Packed_Output = 9;

	var industrialMade_Output = 10;
	var iundustrialUsedToPrint_Output = 11
	var frIndustrialUsedToRefresh_Output = 12;
	var frIndustrialPacked_Output = 13;

	var pbtMade_Output = 14;
	var pbtUsedToPrint_Output = 15
	var frPBTUsedToRefresh_Output = 16;
	var frPBTacked_Output = 17;

	var frPa12_2kg_Output = 18;
	var frPa12_6kg_Output = 19;
	var spPa12_2kg_Output = 20;
	var spPa12_4kg_Output = 21;
	var spPa12_6kg_Output = 22;
	var frIndustrial_6kg_Output = 23;
	var frPBT_6kg_Output = 24;
	var frPa11_2kg_Output = 25;
	var frPa11_6kg_Output = 26;
	var spPa11_2kg_Output = 27;
	var spPa11_4kg_Output = 28;
	var spPa11_6kg_Output = 29;
	var prPa11UsedToPrint_Output = 30;
	var pa11ESD_2kg_Output = 31;
	var pa11CF_6kg_Output = 32;
	var pp_6kg_Output = 33;
	var flexaGrey_2kg_Output = 34;
	var flexaGrey_6kg_Output = 35;
	var flexaBlac_2kg_Output = 36;
	var flexaBlack_6kg_Output = 37;
	var flexaBright_2kg_Output = 38;
	var flexaPerformance_6kg_Output = 39;
	var tpe_2kg_Output = 40;
	var flexaSoft_Output = 41;
	var sealer_Output = 42;

	var frPa12Made_Output = 44;
	var frIndustrialMade_Output = 45;
	var frPBTMade_Output = 46;
	var frPa11Made_Output = 47;
	var flexaGreyMade_Output = 48;
	var flexaBlackMade_Output = 49;
	var flexaPerformanceMade_Output = 50;
	var ppMade_Output = 51;
	var pa11ESDMade_Output = 52;
	var pa11CFMade_Output = 53;
	//------------------------------------  
	//Check Dates
	var weekDate;
	var currentDate = new Date();
	var lastDayOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth()+1, 0);

	//Execute raport only in Friday or Last day of month
	if((currentDate.getDay() != 5) && (currentDate.getDate() != lastDayOfMonth.getDate()))
	{
	//return;
	}

	var actualMonth = currentDate.getUTCMonth();
	//Manual Generator, uncooment and set month number to generate
	//var actualMonth = 5;
	
	//Main Loop summing up materials from actual month, from propoer sheet and one external sheet
	for (var s in allsheets){
		var sheet = allsheets[s];
		var sheetRange = sheet.getRange("A1:Q110");
		var rngSheet = sheetRange.getValues();

		//Check if sheet is valid
		var sheetNameValidation = sheet.getName().substring(0, 4);
		if(sheetNameValidation != "tydz")
		{
		  continue;
		}
		//Check every week in sheet (1 week == 21 rows, so we check every 21 rows)
		for(var i = 5; i < 90; i += 21)
		{
		  weekDate = new Date(rngSheet[i][dateColumn]);
		  //Check if it's current month
		  if(weekDate.getUTCMonth() == actualMonth && weekDate.getUTCFullYear() == currentDate.getUTCFullYear())
		  {
			prMade += +rngSheet[i + prSum_Offset][pa12Column];
			pr11Made += +rngSheet[i + prSum_Offset][pa11Column];
			industrialMade += +rngSheet[i + prSum_Offset][industrialColumn];
			pbtMade += +rngSheet[i + prSum_Offset][pbtColumn];
			spMade += +rngSheet[i + spSum_Offset][pa12Column];
			frPa12UsedToRefresh += +rngSheet[i + prFresh_Offset][pa12Column];
			frIndustrialUsedToRefresh += +rngSheet[i + prFresh_Offset][industrialColumn];
			frPBTUsedToRefresh += +rngSheet[i + prFresh_Offset][pbtColumn];

			frPa12_2kg += +rngSheet[i + fr2kg_Offset][pa12Column];
			frPa12_6kg += +rngSheet[i + fr6kg_Offset][pa12Column];
			frIndustrial_6kg += +rngSheet[i + fr6kg_Offset][industrialColumn];
			frPBT_6kg += +rngSheet[i + fr6kg_Offset][pbtColumn];
			spPa12_2kg += +rngSheet[i + sp2kg_Offset][pa12Column];
			spPa12_4kg += +rngSheet[i + sp4kg_Offset][pa12Column];
			spPa12_6kg += +rngSheet[i + sp6kg_Offset][pa12Column];
			frPa11_2kg += +rngSheet[i + fr2kg_Offset][pa11Column];
			frPa11_6kg += +rngSheet[i + fr6kg_Offset][pa11Column];
			spPa11_2kg += +rngSheet[i + sp2kg_Offset][pa11Column];
			spPa11_4kg += +rngSheet[i + sp4kg_Offset][pa11Column];
			spPa11_6kg += +rngSheet[i + sp6kg_Offset][pa11Column];
			pa11ESD_2kg += +rngSheet[i + fr2kg_Offset][pa11ESDColumn];
			pa11CF_6kg += +rngSheet[i + fr6kg_Offset][pa11CFColumn];
			pp_6kg += +rngSheet[i + fr6kg_Offset][ppColumn];
			flexaGrey_2kg += +rngSheet[i + fr2kg_Offset][flexaGreyColumn];
			flexaGrey_6kg += +rngSheet[i + fr6kg_Offset][flexaGreyColumn];
			flexaBlac_2kg += +rngSheet[i + fr2kg_Offset][flexaBlackColumn];
			flexaBlack_6kg += +rngSheet[i + fr6kg_Offset][flexaBlackColumn];
			flexaBright_2kg += +rngSheet[i + fr2kg_Offset][flexaBrightColumn];
			flexaPerformance_6kg += +rngSheet[i + fr6kg_Offset][flexaPerformanceColumn];
			tpe_2kg += +rngSheet[i + fr2kg_Offset][tpeColumn];
			flexaSoft += +rngSheet[i + fr2kg_Offset][flexaSoftColumn];
			sealer += +rngSheet[i + bottle_Offset][sealerColumn];
			frPa12Made += +rngSheet[i + fresh_Offset][pa12Column];
			frIndustrialMade += +rngSheet[i + fresh_Offset][industrialColumn];
			frPBTMade += +rngSheet[i + fresh_Offset][pbtColumn];
			frPa11Made += +rngSheet[i + fresh_Offset][pa11Column];
			flexaGreyMade += +rngSheet[i + fresh_Offset][flexaGreyColumn];
			flexaBlackMade += +rngSheet[i + fresh_Offset][flexaBlackColumn];
			flexaPerformanceMade += +rngSheet[i + fresh_Offset][flexaPerformanceColumn];
			ppMade += +rngSheet[i + fresh_Offset][ppColumn];
			pa11ESDMade += +rngSheet[i + fresh_Offset][pa11ESDColumn];
			pa11CFMade += +rngSheet[i + fresh_Offset][pa11CFColumn];
			
		  }
		}
	}
	spPacked += (spPa12_2kg * 2) + (spPa12_4kg * 4) + (spPa12_6kg * 6);
	frPa12Packed += (frPa12_2kg * 2) + (frPa12_6kg * 6);
	frPa11Packed += (frPa11_2kg * 2) + (frPa11_6kg * 6);
	frIndustrialPacked += (frIndustrial_6kg * 6);
	frPBTPacked += (frPBT_6kg * 6);


	//Loop counting used powder from diffrent spreadsheet
	for(var i = 0; i < 879; i++)
	{
		var monthDate = new Date(rngPowderUsed[i][1]);
		if(monthDate.getUTCMonth() == actualMonth && monthDate.getUTCFullYear() == currentDate.getUTCFullYear())
		{
			var prConvertToNumber = rngPowderUsed[i][2].toString().replace(".",",");
			var industrialConvertToNumber = rngPowderUsed[i][6].toString().replace(".",",");
			var prToInt = parseFloat(rngPowderUsed[i][2]);
			var industrialToInt = parseFloat(rngPowderUsed[i][6]);
			var prPa11ToInt = parseFloat(rngPowderUsed[i][10])

			if(!isNaN(prToInt))
			{
			  prUsedToPrint += prToInt;
			}
			if(!isNaN(industrialToInt))
			{
			  iundustrialUsedToPrint += industrialToInt;
			}
			if(!isNaN(prPa11ToInt))
			{
			  prPa11UsedToPrint += prPa11ToInt;
			}
		}
	}

	//Fill raport with data colected in main loop
	if(raportRange.getCell(month_Output, raportSheet.getLastColumn()).getValue() == actualMonth + 1)
	{
		raportSheet.getRange(1, raportSheet.getLastColumn(), 55, 1).clearContent();
	}
	var rowToAdd = raportSheet.getLastColumn() + 1;
	raportRange.getCell(month_Output,rowToAdd).setValue(actualMonth + 1);

	raportRange.getCell(industrialMade_Output,rowToAdd).setValue(industrialMade);
	raportRange.getCell(iundustrialUsedToPrint_Output,rowToAdd).setValue(iundustrialUsedToPrint);
	raportRange.getCell(frIndustrialUsedToRefresh_Output,rowToAdd).setValue(frIndustrialUsedToRefresh);
	raportRange.getCell(frIndustrialPacked_Output,rowToAdd).setValue(frIndustrialPacked);
	raportRange.getCell(frIndustrial_6kg_Output,rowToAdd).setValue(frIndustrial_6kg);
	raportRange.getCell(frIndustrialMade_Output,rowToAdd).setValue(frIndustrialMade);

	raportRange.getCell(pbtMade_Output,rowToAdd).setValue(pbtMade);
	raportRange.getCell(pbtUsedToPrint_Output,rowToAdd).setValue(pbtUsedToPrint);
	raportRange.getCell(frPBTUsedToRefresh_Output,rowToAdd).setValue(frPBTUsedToRefresh);
	raportRange.getCell(frPBTacked_Output,rowToAdd).setValue(frPBTPacked);
	raportRange.getCell(frPBT_6kg_Output,rowToAdd).setValue(frPBT_6kg);
	raportRange.getCell(frPBTMade_Output,rowToAdd).setValue(frPBTMade);

	raportRange.getCell(prMade_Output,rowToAdd).setValue(prMade);
	raportRange.getCell(pr11Made_Output,rowToAdd).setValue(pr11Made);
	raportRange.getCell(prUsedToPrint_Output,rowToAdd).setValue(prUsedToPrint);
	raportRange.getCell(spMade_Output,rowToAdd).setValue(spMade);
	raportRange.getCell(spPacked_Output,rowToAdd).setValue(spPacked);
	raportRange.getCell(frPa12UsedToRefresh_Output,rowToAdd).setValue(frPa12UsedToRefresh);
	raportRange.getCell(frPa12Packed_Output,rowToAdd).setValue(frPa12Packed);
	raportRange.getCell(frPa11Packed_Output,rowToAdd).setValue(frPa11Packed);
	raportRange.getCell(frPa12_2kg_Output,rowToAdd).setValue(frPa12_2kg);
	raportRange.getCell(frPa12_6kg_Output,rowToAdd).setValue(frPa12_6kg);
	raportRange.getCell(spPa12_2kg_Output,rowToAdd).setValue(spPa12_2kg);
	raportRange.getCell(spPa12_4kg_Output,rowToAdd).setValue(spPa12_4kg);
	raportRange.getCell(spPa12_6kg_Output,rowToAdd).setValue(spPa12_6kg);
	raportRange.getCell(frPa11_2kg_Output,rowToAdd).setValue(frPa11_2kg);
	raportRange.getCell(frPa11_6kg_Output,rowToAdd).setValue(frPa11_6kg);
	raportRange.getCell(spPa11_2kg_Output,rowToAdd).setValue(spPa11_2kg);
	raportRange.getCell(spPa11_4kg_Output,rowToAdd).setValue(spPa11_4kg);
	raportRange.getCell(spPa11_6kg_Output,rowToAdd).setValue(spPa11_6kg);
	raportRange.getCell(prPa11UsedToPrint_Output,rowToAdd).setValue(prPa11UsedToPrint);
	raportRange.getCell(pa11ESD_2kg_Output,rowToAdd).setValue(pa11ESD_2kg);
	raportRange.getCell(pa11CF_6kg_Output,rowToAdd).setValue(pa11CF_6kg);
	raportRange.getCell(pp_6kg_Output,rowToAdd).setValue(pp_6kg);
	raportRange.getCell(flexaGrey_2kg_Output,rowToAdd).setValue(flexaGrey_2kg);
	raportRange.getCell(flexaGrey_6kg_Output,rowToAdd).setValue(flexaGrey_6kg);
	raportRange.getCell(flexaBlac_2kg_Output,rowToAdd).setValue(flexaBlac_2kg);
	raportRange.getCell(flexaBlack_6kg_Output,rowToAdd).setValue(flexaBlack_6kg);
	raportRange.getCell(flexaBright_2kg_Output,rowToAdd).setValue(flexaBright_2kg);
	raportRange.getCell(flexaPerformance_6kg_Output,rowToAdd).setValue(flexaPerformance_6kg);
	raportRange.getCell(tpe_2kg_Output,rowToAdd).setValue(tpe_2kg);
	raportRange.getCell(flexaSoft_Output,rowToAdd).setValue(flexaSoft);
	raportRange.getCell(sealer_Output,rowToAdd).setValue(sealer);
	raportRange.getCell(frPa12Made_Output,rowToAdd).setValue(frPa12Made);
	raportRange.getCell(frPa11Made_Output,rowToAdd).setValue(frPa11Made);
	raportRange.getCell(flexaGreyMade_Output,rowToAdd).setValue(flexaGreyMade);
	raportRange.getCell(flexaBlackMade_Output,rowToAdd).setValue(flexaBlackMade);
	raportRange.getCell(flexaPerformanceMade_Output,rowToAdd).setValue(flexaPerformanceMade);
	raportRange.getCell(ppMade_Output,rowToAdd).setValue(ppMade);
	raportRange.getCell(pa11ESDMade_Output,rowToAdd).setValue(pa11ESDMade);
	raportRange.getCell(pa11CFMade_Output,rowToAdd).setValue(pa11CFMade);
}