var Excel = require('exceljs'),
workbook1 = new Excel.Workbook(),
sheet = workbook1.addWorksheet('Company Ranking'),
XLSX = require('xlsx'),
fs=require('fs'),
readlineSync = require('readline-sync'),
scraper = require('table-scraper'),
db=require('./Database.js');
var connection=db.mysql_connection();
var workbook;

// Deleting excel file if existed
try
{
	fs.unlinkSync('Company Ranking.xlsx');
}
catch (e){}
sheet.columns = [
{ header: 'Industry Name', key: 'Industry Name', width: 15 },
{ header: 'permID', key: 'permID', width: 12 },
{ header: 'TRBC 2012 Hierarchical ID', key: 'hID', width: 10 },
{ header: 'Link', key: 'hLink', width: 15 },
{ header: 'Ticker', key: 'Ticker', width: 10 },
{ header: 'Name', key: 'Name', width: 15 },
{ header: 'MarketCapitalizatio', key: 'MarketCapitalization', width: 10 },
{ header: 'TTM Sales $', key: 'TTM Sales $', width: 10 },
{ header: 'Employees', key: 'Employees', width: 5 },
{ header: 'Employees Link', key: 'eLink', width: 32 }
];

// preparing excel file to read Industry Names
try
{
	workbook= XLSX.readFile("BusinessClassification.xlsx",{type:"array"});
}

catch (e)
{
	console.log("BusinessClassification.xlsx not found!"); process.exit(0);
}
var worksheet = workbook.Sheets[workbook.SheetNames[0]];  //Selecting sheet
var count=0,
countIndustries=0;
jsonIndex=1,
excelRowIndex=1,
till=-1;
MAX_EXCEL_ROWS_TO_READ=300;

console.log('1 to extract all companies for all industries(around 10,000 companies existed).\n2 to specify number of industries to extact companies')
var choice = readlineSync.question();
if(choice==2)
{
	console.log('Enter number of industries you wish to extract!\n');
	till = readlineSync.question();
}
if(till==0 && choice==2)
	process.exit(0);


/* Takes industries one by one from the excel(given by blueOptima with problem statement). For each industries, scrapes all companies from thomsonreuters.
For each industries in the excel, it searches the database (created during 1_CrawlIndustyCodeToDatabase) for industry code.
Using industry code and url(https://www.reuters.com/sectors/industries/rankings?industryCode=) , it goes to website and scrapes companies */

function nextIndustry()
{
	var cell_industry=worksheet['D'+excelRowIndex]
	var cell_permID=worksheet['E'+excelRowIndex]
	var cell_hirarchicalID=worksheet['F'+excelRowIndex]
	var val_industry = (cell_industry ? cell_industry.v :undefined );
	var val_permID = (cell_permID ? cell_permID.v :undefined );
	var val_hirarchicalID = (cell_hirarchicalID ? cell_hirarchicalID.v :undefined );
	if(val_industry==undefined && val_permID==undefined && val_hirarchicalID==undefined)
	{ 

		if(excelRowIndex<MAX_EXCEL_ROWS_TO_READ && ((till==-1) || (till!=-1 && till>countIndustries)))
		{
			excelRowIndex++;
			nextIndustry();
		}
		else
		{
			workbook1.xlsx.writeFile('Company Ranking.xlsx').then(function()
			{
				console.log("Excel Exported: Company Ranking.xlsx");
				process.exit(0)
			});
		}
	}
	else
	{
		var sql='select indID from industry_code where indy_name="'+val_industry+'"';
		try
		{
			connection.query(sql,function(err,result)
			{
				if(err) throw err;
				if(result[0]!=undefined)
				{
					countIndustries++;
					var URL = "https://www.reuters.com/sectors/industries/rankings?industryCode="+result[0].indID+"&view=size&page=-1&sortby=mktcap&sortdir=DESC";
					scraper.get(URL).then(function(tableData)
					{
	                		// tableData[0] will contain ranking table from thomsonreuters in JSON format
	                		if(tableData[0]!=undefined)
	                		{
	                			sheet.addRow({ 'Industry Name':val_industry,'permID':val_permID,'hID':val_hirarchicalID,'hLink':URL});
	                			for(var jsonCompanyIndex=0;jsonCompanyIndex<tableData[0].length;jsonCompanyIndex++)
	                			{
		                   			 //adding values to json
		                   			 tableData[0][jsonCompanyIndex].indId=result[0].indID;
		                   			 tableData[0][jsonCompanyIndex].indName=val_industry;
		                   			 tableData[0][jsonCompanyIndex].eLink="https://www.reuters.com/finance/stocks/company-officers/"+(tableData[0][jsonCompanyIndex].Ticker);

				                    // adding row to excel object
				                    sheet.addRow(tableData[0][jsonCompanyIndex]);

				                    // renaming JSON key name
				                    var temp=tableData[0][jsonCompanyIndex]["TTM Sales $"];
				                    delete tableData[0][jsonCompanyIndex]["TTM Sales $"];
				                    tableData[0][jsonCompanyIndex].TTM_Sales_$=temp;
				                }
				                var sql = "INSERT INTO company_ranking SET ?";
				                jsonIndex=1;
				                count=0;

		                  		// inserts ranking table data into database row by row
		                  		function insertRow()
		                  		{
		                  			if(jsonIndex<tableData[0].length)
		                  			{
		                  				try
		                  				{
		                  					connection.query(sql,tableData[0][jsonIndex],function(err,rows)
		                  					{
		                  						if(err) throw err;
		                  						count++;
		                  						jsonIndex++;
		                  						insertRow();
		                  					});
		                  				}
		                  				catch(e)
		                  				{
		                  					console.log("Cant establish connection")
		                  				}
		                  			}
		                  			else
		                  			{
		                  				console.log("Industry code: "+result[0].indID+"  count: "+count);
		                  				if(excelRowIndex<MAX_EXCEL_ROWS_TO_READ && ((till==-1) || (till!=-1 && till>countIndustries)))
		                  				{
		                  					excelRowIndex++;
		                  					nextIndustry();
		                  				}
		                  				else
		                  				{
		                  					workbook1.xlsx.writeFile('Company Ranking.xlsx').then(function()
		                  					{
		                  						console.log("Excel Exported: Company Ranking.xlsx");
		                  						process.exit(0)
		                  					});
		                  				}
		                  			}
		                  		}
		                  		insertRow();
		                  	}
		                  	else
		                  	{
		                  		if(excelRowIndex<MAX_EXCEL_ROWS_TO_READ && ((till==-1) || (till!=-1 && till>countIndustries)))
		                  		{
		                  			excelRowIndex++;
		                  			nextIndustry();
		                  		}
		                  		else
		                  		{
		                  			workbook1.xlsx.writeFile('Company Ranking.xlsx').then(function()
		                  			{
		                  				console.log("Excel Exported: Company Ranking.xlsx");
		                  				process.exit(0)
		                  			});
		                  		}
		                  	}
	                });
	        	} // undefined if check close
	        	else
	        	{
	        		if(excelRowIndex<MAX_EXCEL_ROWS_TO_READ && ((till==-1) || (till!=-1 && till>countIndustries)))
	        		{
		                /* Calling funtion for next industry after complete processing of an industry ranking,
		                will provide synchronization */
		                excelRowIndex++;
		                nextIndustry()
	            	}
		            else
		            {
		            	workbook1.xlsx.writeFile('Company Ranking.xlsx').then(function()
		            	{
		            		console.log("Excel Exported: Company Ranking.xlsx");
		            		process.exit(0)
		            	});
		            }
		        }

	      	})	//End of database connection
		}
		catch(e)
		{
			console.log("Cant stablish database connection")
		}
	}
}

// clearing previous data in database, before starting
try
{
	connection.query("delete from company_ranking",function(errDelete,resultDelete)
	{
		if(errDelete) throw errDelete;
		console.log("Database cleared");
		nextIndustry();
	})
}
catch(e)
{
	console.log("Cant establish connection")
}
