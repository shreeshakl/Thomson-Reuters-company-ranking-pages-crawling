var mysql = require('mysql');

exports.mysql_connection = function()
{
  var con = mysql.createConnection
  ({
    host: "localhost",
    database:"blueoptima_db",
    user: "root",
    password: "greatgrass"
  });
return con;
}
