var db=require('./Database.js'),
connection=db.mysql_connection(),
fs = require('fs'),
request = require('request'),
cheerio = require('cheerio'),
industryCode=0,
noOfEntries=0,
MAX_INDUSTRY_CODE=300;


/* Scrapes Ranking page one by one by iterating industry code */
function scrapePage()
{
    	var URL = "https://www.reuters.com/sectors/industries/rankings?view=size&industryCode="+industryCode;
    	/* makes an HTTP request for the page to be scraped */
      	request(URL, function(error, response, responseHtml)
		{
					if(!error && response.statusCode === 200)
					{
						var $ = cheerio.load(responseHtml),
						title = $('#sectionTitle h1').html();
						if(title!=null)
						{
					   		noOfEntries++;
					   		var industryName=(title.substring(0,title.indexOf('<'))).replace('&amp;','&').trim();
							var insert = {indId:industryCode,indy_name:industryName};
							var sql = "INSERT INTO industry_code SET ?";
			                try
			                {
			                  	connection.query(sql,insert,function(err,rows)
			                  	{
			                    	if(err) throw err;
					                console.log("Inserted:"+industryName)
			                  	});
			                }
			                catch (e)
			                {
			                  	console.log("Data not inserted! Something went wrong with database.")
			                }

						}
					}
					industryCode++;
					if(industryCode<MAX_INDUSTRY_CODE)
						scrapePage();
					else
					{
						console.log(noOfEntries);
						process.exit(0)
					}
    	});
}

try
{
  connection.query("delete from industry_code",function(errDelete,resultDelete)
  {
    if(errDelete) throw errDelete;
    console.log("Database cleared");
    // scrapePage excution begins here
    scrapePage();
  })
}
catch(e)
{
  console.log("Something went wrong with database")
}
