var mysql = require('mysql');

exports.mysql_connection = function()
{
  var con = mysql.createConnection
  ({
    host: "localhost",
    database:"blueoptima_db",
    user: "YOUR USERNAME HERE",
    password: "YOUR PASSWORD HERE"
  });
return con;
}
