var XLSX = require('xlsx'),
json2xls = require('json2xls'),
readlineSync = require('readline-sync'),
Excel = require('exceljs');
var workbook1 = new Excel.Workbook();
var sheet = workbook1.addWorksheet('Employees');
sheet.columns =[
{ header: 'Industry', key: 'Industry', width:  15},
{ header: 'Company Name', key: 'Company Name', width: 15},
{ header: 'Link', key: 'hLink', width: 20 },
{ header: 'Name', key: 'Name', width: 20 },
{ header: 'Age', key: 'Age', width: 4 },
{ header: 'Since', key: 'Since', width: 5 },
{ header: 'Current Position', key: 'Current Position', width: 20 },
{ header: 'Description', key: 'Description', width: 500 },
];

var fs=require('fs'),
scraper = require('table-scraper'),
mysql = require('mysql'),
sql='select ticker,name,indName from company_ranking',
db=require('./Database.js'),
connection=db.mysql_connection();
try { fs.unlinkSync('Employees.xlsx'); }
catch (e) {}

var companyIndex=1,till=0;
console.log('1 to extract all employees of all companies(around 1,00,000 employees).\n2 to provide input number of companies');
var choice = readlineSync.question();
if(choice==2)
{

	console.log('Enter number of companies emplyoees info you wish to extract.');
	till = readlineSync.question();
}
if(till==0 && choice==2)
	process.exit(0);

try
{
	connection.query("delete from company_officers",function(errDelete,resultDelete)
	{
		if(errDelete) throw errDelete;
		console.log("Database cleared");
		connection.query(sql,function(err,DB_Result)
		{
			if(err) throw err;
			if(DB_Result!=undefined)
			{
				if(choice==1)
					till=DB_Result.length
				/* scrape employees for each company in the database(created in 2_CrawlCompanyRankingToDatabase.js) */
				function loopTicker()
				{
					var URL = "https://www.reuters.com/finance/stocks/company-officers/"+DB_Result[companyIndex].ticker;
					scraper.get(URL).then(function(tableData)
					{
						if(tableData[0]!=undefined && tableData[0]!=null && tableData!=undefined && tableData[0][0].Age!=undefined )
						{
							sheet.addRow({'Industry':DB_Result[companyIndex].indName, 'Company Name':DB_Result[companyIndex].name,'hLink':URL });
							for(var i=0;i<tableData[0].length;i++)
							{
		                       //adding description from 2nd table's JSON object of webpage, to 1st table JSON object
		                       tableData[0][i].Description = tableData[1][i].Description;
		                       sheet.addRow(tableData[0][i])

		                        // renaming key
		                        var temp=tableData[0][i]["Current Position"];
		                        delete tableData[0][i]["Current Position"];
		                        tableData[0][i].Current_Position=temp;
		                    }
		                    var sql = "INSERT INTO company_officers SET ?";
		                    var employeeIndex=1;
		                    var countNumberOfEmpInCompany=0;

	                      // insert's emplyoee info to database row by row
	                      function insertEmployeeRow()
	                      {
	                      	if(employeeIndex<tableData[0].length)
	                      	{
	                      		connection.query(sql,tableData[0][employeeIndex],function(err,rows)
	                      		{
	                      			if(err) throw err;
	                      			countNumberOfEmpInCompany++;
	                      			employeeIndex++;
	                      			insertEmployeeRow();
	                      		});
	                      	}
	                      	else
	                      	{

	                      		console.log(companyIndex+"\t"+"Ticker: "+DB_Result[companyIndex].ticker+"\tCount: "+countNumberOfEmpInCompany )
	                      		if(companyIndex<till-1)
	                      		{
	                      			companyIndex++
	                      			loopTicker()
	                      		}
	                      		else
	                      		{
	                      			workbook1.xlsx.writeFile('Employees.xlsx').then(function()
	                      			{
	                      				console.log("Excel exported: Employees.xlsx");
	                      				process.exit(0)
	                      			});
	                      		}
	                      	}
	                      }
	                      insertEmployeeRow()
	                    }
	                    else
	                    {
		                  	if(companyIndex<till-1)
		                  	{
		                  		companyIndex++;
		                  		loopTicker();
		                  	}
		                  	else
		                  	{
		                  		workbook1.xlsx.writeFile('Employees.xlsx').then(function()
		                  		{
		                  			console.log("Excel exported: Employees.xlsx");
		                  			process.exit(0)
		                  		});
		                  	}
	                    }
	             	});
      			}  // end of  loopTicker()
     			loopTicker()
   			}   //end of if(DB_Result[0]!=undefined)
  		}) // End of DB

  	}) //End of clear db
}
catch (e)
{
	console.log("Can't establish database connection.")
}
