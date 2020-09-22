'use strict'

// database config
let model = require("./lib/database.js");

// required modules
let express = require("express");
let app = express();
let sql = require('mssql');
let hb = require('handlebars');
let nodeSSPI = require('node-sspi');
let nodemailer = require("nodemailer");
let csv = require('csv');

// app running on port 3100
app.set('port', process.env.PORT || 3100);
app.use(express.static(__dirname + '/views')); // allows direct navigation to static files
app.use(require("body-parser").urlencoded({extended: true})); // parse form submissions

//Express-Handlebars module & setting up templating engine
let handlebars =  require("express-handlebars")
.create({ defaultLayout: "main"});
app.engine("handlebars", handlebars.engine);
app.set("view engine", "handlebars");      

// for Handlebars Index value: this starts index at 1 as opposed to 0
hb.registerHelper("inc", function(value, options){
    return parseInt(value) + 1;
});

// send email function with 5 params to customize each email notification
async function sendEmail(user, err, pathinfo, message, email) {

    process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
    
  // create reusable transporter object using the default SMTP transport
  let transporter = nodemailer.createTransport({
    host: "rec-o365-01.resers.com",
    port: 25,
    secure: false, // true for 465, false for other ports
      tls: { secureProtocol: "TLSv1_method" }
  });

let uname = email + "," + user;
  // send mail with defined transport object
  let info = await transporter.sendMail({
    from: '"ICLMS" <noreply@resers.com>', // sender address
    to: uname, // list of receivers
    subject: "ICLMS: Notification", // Subject line
    html: message + "<br> <br> <hr> <br> Username: " + user + " <br> Date: " + new Date().toJSON().slice(0,10) + "<br> Error: <b style='color:red'>"  + err + "</b> <br> URL where notification came from: " + pathinfo + "<br><br> <i>This is an automated message. Please do not reply.</i>" // body
  });
}

// return all plant names and numbers
new sql.ConnectionPool(model.config2).connect().then(pool => {
    
let querycode = 'select plant, plantname from tblplant';
     
return pool.query(querycode);
}).then(result => {   
 let plantNums = "";
        try{
global.plantNums = result.recordset;
        } catch {
global.plantNums = "";
}
}); 
		
// return all from LX MD_Item table
new sql.ConnectionPool(model.config3).connect().then(pool => {
    
let querycode = "select distinct CaseUPC, itemdescription from vw_Item where CaseUPC is not null";
     
return pool.query(querycode);
}).then(result => {   
	
global.lxInfo = result.recordset;

}).catch(console.error); 

// windows authentication with node-sspi module
app.use(function (req, res, next) {
  var nodeSSPIObj = new nodeSSPI({
    retrieveGroups: true
  })
  nodeSSPIObj.authenticate(req, res, function(err){
    res.finished || next()
  })
})

// global variables
app.locals.currentYear = new Date().getFullYear(); //current year variable

// search page
app.get('/', function(req, res) { 

app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

// pass params and return records from Products table
new sql.ConnectionPool(model.config).connect().then(pool => {
    
let querycode = 'SELECT * from Products_NEW where (GTIN like ' + "'" + req.query.gtin + "'" + ' or FG like ' +  "'" + req.query.fg + "'" + ' or LxDescription like ' +  "'" + req.query.lx + "'" + ' or Formula like ' +  "'" + req.query.formula + "')";
    
return pool.query(querycode);
}).then(result => {   

res.render('search', {title: "Search", data: result.recordset, lx: req.query.lx, gtin: req.query.gtin, formula: req.query.formula, fg: req.query.fg})

}).catch(err => {
console.log(err);
    
res.render('failure', { title: "Failure", err: err})
}); 
}); 

//create page where new record gets inserted into database table
app.use('/create', function(req, res) { 
    
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
    
if (req.method === 'POST'){
// pass params and insert record into Products_NEW table
    
//get multiple checkbox values via request object and then concatenate into string    
global.facil = ''; 
    
for(var i=0; i<plantNums.length; i++){  
    
if (req.body[plantNums[i].plant] == 'on'){  
    
global.facil = global.facil + plantNums[i].plant + ',';
}   
    }
    
if (req.body.CORP == 'on'){  
global.facil = global.facil + 'CORP,';
}
    
global.facil = (global.facil).substring(0, (global.facil).length - 1);
      
new sql.ConnectionPool(model.config).connect().then(pool => {
//create and insert record into database table 
let querycode = 'INSERT INTO [dbo].[Products_NEW] ([Brand] ,[Country] ,[Facility] ,[PartNum] ,[FG] ,[GTIN] ,[LxDescription] ,[LabelDescription] ,[SubDescription] ,[Formula] ,[Shelf] ,[Guarantee] ,[SafeHandling] ,[NetWeight] ,[WeightUOM] ,[PackSize] ,[NetContent] ,[NetWeightUOM] ,[DateCode] ,[CADDateCode] ,[DistrStatement] ,[Logo] ,[Ingredients] ,[InspectionLegend] ,[InspectionCode], [LastUpdatedBy], [IsActive], [NutritionImage],[RepackDescription],[RepackNutrition],[RepackIngredients], [UnitUPC]) VALUES (' + "'" + req.body.Brand.replace(/'/g, "''") + "'" + ','+ "'" + req.body.country + "'" +','+ "'" + global.facil + "'" +','+ "'" + req.body.PartNum + "'" +','+ "'" + req.body.fg + "'" +','+ "'" + req.body.gtin + "'" +','+ "'" + req.body.lx.replace(/'/g, "''") + "'" +','+ "'" + req.body.LabelDescription.replace(/'/g, "''") + "'" +','+ "'" + req.body.SubDescription.replace(/'/g, "''") + "'" +','+ "'" + req.body.formula.replace(/'/g, "''") + "'" +','+ "'" + req.body.Shelf + "'" +','+ "'" + req.body.Guarantee + "'" +','+ "'" + req.body.safe + "'" +','+ "'" + req.body.NetWeight + "'" +','+ "'" + req.body.WeightUOM + "'" +','+ "'" + req.body.PackSize + "'" +','+ "'" + req.body.NetContent + "'" +','+ "'" + req.body.NetWeightUOM + "'" +','+ "'" + req.body.DateCode + "'" +','+ "'" + req.body.CADDateCode + "'" +','+ "'" + req.body.DistrStatement.replace(/'/g, "''") + "'" +','+ "'" + req.body.Logo.replace(/'/g, "''") + "'" +','+ "'" + req.body.Ingredients.replace(/'/g, "''") + "'" +','+ "'" + req.body.InspectionLegend.replace(/'/g, "''") + "'" +',' + "'" + req.body.InspectionCode + "'" + ',' + "'" + app.locals.currentUser + "'" + ', 1, ' + "'" + req.body.NutritionImage.replace(/'/g, "''") + "'" + ',' + "'" + req.body.RepackDescription.replace(/'/g, "''") + "'" + ',' + "'" + req.body.RepackNutrition.replace(/'/g, "''") + "'"  + ',' + "'" + req.body.RepackIngredients.replace(/'/g, "''") + "'" + ',' + "'" + req.body.UnitUPC + "'" + ')';
 
global.fg = req.body.fg;
    
return pool.query(querycode);
}).then(result => {  
    
res.render('success', { title: "Success", upc: req.query.upc})
    
sendEmail(app.locals.currentUser, "There are no errors.", (req.protocol + '://' + req.get('host') + req.originalUrl), "A record has been successfully created by the user below. If you are " + app.locals.currentUser + ", then please make a Samanage ticket <a href='https://app.samanage.com/'>at this location</a> for record with UPC# " + global.fg, "ICLMSNotifications@resers.com").catch(console.error);
    
}).catch(err => {
res.render('failure', { title: "Failure", err: err})
    
console.log(err);
    
sendEmail(app.locals.currentUser, err, (req.protocol + '://' + req.get('host') + req.originalUrl), "Your record was not successfully created and resulted in error. Please error see details below.", "ICLMSAdmins@resers.com").catch(console.error);
}); 
} else {

if(req.query.upc){
new sql.ConnectionPool(model.config3).connect().then(pool => {
    
let querycode = "select * from vw_Item where CaseUPC = " + "'00" + req.query.upc.replace(/\./g,'') + "'";
     		
return pool.query(querycode);
}).then(result => {   
	
res.render('create', {title: "Create New Record", plantNums: plantNums, lxInfo: lxInfo, lxResult: result.recordset[0], queryString: req.query.upc}) 

}).catch(console.error); 
}   else { 
res.render('create', {title: "Create New Record", plantNums: plantNums, lxInfo: lxInfo, queryString: req.query.upc}) }
}}); 

// historylist page with a clickable list of all changes to a record
app.get('/historylist', function(req, res) { 
    
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

// pass param and return records from Products_NEW table
new sql.ConnectionPool(model.config).connect().then(pool => {
    
let querycode = 'SELECT * from Products_Changes WHERE ProductID = ' + "'" + req.query.ProductID + "'" + ' order by ID desc';
     
return pool.query(querycode);
}).then(result => {   
    try {
res.render('historylist', {title: "History", upc: result.recordset[0].FG, data: result.recordset})
    } catch {
res.render('historylist', {title: "History"})      
    }

}).catch(err => {
    res.render('failure', { title: "Failure", err: err})
    
    sendEmail(app.locals.currentUser, err, (req.protocol + '://' + req.get('host') + req.originalUrl), "Your request was unsuccessful. Please see error details below.", "ICLMSAdmins@resers.com").catch(console.error);
    
    console.log(err);
}); 
}); 

// details page where each field of a record is rendered
app.get('/details', function(req, res) { 

app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

// pass params and return records from Products_NEW table
new sql.ConnectionPool(model.config).connect().then(pool => {
    
let querycode = 'SELECT * from Products_NEW WHERE ProductID = ' + "'" + req.query.ProductID + "'";
     
return pool.query(querycode);
}).then(result => {   

res.render('details', {title: "Details", ProductID: req.query.ProductID, data: result.recordset[0]})

}).catch(err => {
    res.render('failure', { title: "Failure", err: err})
    
    console.log(err);
    
    sendEmail(app.locals.currentUser, err, (req.protocol + '://' + req.get('host') + req.originalUrl), "There appears to be an issue with your request. Please see error details below.", "ICLMSAdmins@resers.com").catch(console.error);
}); 
}); 

// Update page
app.use('/update', function(req,res) {
    
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
    
     if (req.method === 'POST') {    
new sql.ConnectionPool(model.config).connect().then(pool => {   
    
    //get multiple checkbox values via request object and then concatenate into string    
global.facil = ''; 
    
for(var i=0; i<plantNums.length; i++){  
    
if (req.body[plantNums[i].plant] == 'on'){  
    
global.facil = global.facil + plantNums[i].plant + ',';
}   
    }
if (req.body.CORP == 'on'){  
global.facil = global.facil + 'CORP,';
}
    
global.facil = (global.facil).substring(0, (global.facil).length - 1);
    
//set global variable to blank string    
global.qc = '';
    
// insert old record into Product_Changes table first, and then update original record in Products_NEW
global.qc = 'SELECT * FROM Products_NEW WHERE ProductID = ' + "'" + req.query.ProductID + "'" + '; INSERT INTO [Products_Changes] ([UpdatedOn],[Brand],[Country],[Facility],[PartNum],[FG],[GTIN],[LxDescription],[LabelDescription],[SubDescription],[Formula],[Shelf],[Guarantee],[SafeHandling],[NetWeight],[WeightUOM],[PackSize],[NetContent],[NetWeightUOM],[DateCode],[CADDateCode],[DistrStatement],[Logo],[Ingredients],[InspectionLegend],[InspectionCode], [Username], [NutritionImage],[ProductID],[RepackDescription],[RepackNutrition],[RepackIngredients], [UnitUPC]) (select DATEADD(hh,-7,GETDATE()), [Brand],[Country],[Facility],[PartNum],[FG],[GTIN],[LxDescription],[LabelDescription],[SubDescription],[Formula],[Shelf],[Guarantee],[SafeHandling],[NetWeight],[WeightUOM],[PackSize],[Netcontent],[NetWeightUOM],[DateCode],[CADDateCode],[DistrStatement],[Logo],[Ingredients],[InspectionLegend],[InspectionCode], [LastUpdatedBy], [NutritionImage],[ProductID],[RepackDescription],[RepackNutrition],[RepackIngredients], [UnitUPC] from Products_NEW WHERE ProductID = ' + "'" + req.query.ProductID + "'" + '); UPDATE [Products_NEW] SET [Brand] =' + "'" + req.body.Brand.replace(/'/g, "''") + "'" + ',[Country] ='+ "'" + req.body.country + "'" +',[Facility]='+ "'" + global.facil + "'" +',[PartNum]='+ "'" + req.body.PartNum + "'" +',[FG]='+ "'" + req.body.fg + "'" +',[GTIN]='+ "'" + req.body.gtin + "'" +',[LxDescription]='+ "'" + req.body.lx.replace(/'/g, "''") + "'" +',[LabelDescription]='+ "'" + req.body.LabelDescription.replace(/'/g, "''") + "'" +',[SubDescription]='+ "'" + req.body.SubDescription.replace(/'/g, "''") + "'" +',[Formula]='+ "'" + req.body.formula.replace(/'/g, "''") + "'" +',[Shelf]='+ "'" + req.body.Shelf + "'" +',[Guarantee]='+ "'" + req.body.Guarantee + "'" +',[SafeHandling]='+ "'" + req.body.safe + "'" +',[NetWeight]='+ "'" + req.body.NetWeight + "'" +',[WeightUOM]='+ "'" + req.body.WeightUOM + "'" +',[PackSize]='+ "'" + req.body.PackSize + "'" +',[NetContent]='+ "'" + req.body.NetContent + "'" +',[NetWeightUOM]='+ "'" + req.body.NetWeightUOM + "'" +',[DateCode]='+ "'" + req.body.DateCode + "'" +',[CADDateCode]='+ "'" + req.body.CADDateCode + "'" +',[DistrStatement]='+ "'" + req.body.DistrStatement.replace(/'/g, "''") + "'" +',[Logo]='+ "'" + req.body.Logo + "'" +',[Ingredients]='+ "'" + req.body.Ingredients.replace(/'/g, "''") + "'" +',[InspectionLegend]='+ "'" + req.body.InspectionLegend.replace(/'/g, "''") + "'" +',[InspectionCode] = ' + "'" + req.body.InspectionCode + "'" + ',[LastUpdatedBy] = ' + "'" + app.locals.currentUser  + "'" + ',[NutritionImage] = ' + "'" + req.body.NutritionImage  + "'" + ',[RepackDescription] = ' + "'" + req.body.RepackDescription.replace(/'/g, "''")  + "'" + ',[RepackNutrition] = ' + "'" + req.body.RepackNutrition.replace(/'/g, "''")  + "'" + ',[RepackIngredients] = ' + "'" + req.body.RepackIngredients.replace(/'/g, "''")  + "'" + ',[UnitUPC] = ' + "'" + req.body.UnitUPC  + "'" + ' WHERE ProductID = ' + "'" + req.query.ProductID + "'";

return pool.query(qc);
}).then(result => {
    //if changes happen to any of the 3 critical fields, then send out email to user and admins
    if(req.body.PartNum != result.recordset[0].PartNum || req.body.fg != result.recordset[0].FG || req.body.Ingredients != result.recordset[0].Ingredients){
        
    res.redirect('email?ProductID=' + req.query.ProductID + '&PartNum=' + req.body.PartNum + '&fg=' + req.body.fg + '&Ingredients=' + req.body.Ingredients); 
    } else {
    res.render('success', { title: "Success", ProductID: req.query.ProductID})
    
    }
}).catch(err => {
    res.render('failure', { title: "Failure", err: err})
    
    console.log(err);
    
    sendEmail(app.locals.currentUser, err, (req.protocol + '://' + req.get('host') + req.originalUrl), "Your record was not successfully updated and resulted in error. Please see details below.", "ICLMSAdmins@resers.com").catch(console.error);
}); 
    } else {
    
 new sql.ConnectionPool(model.config).connect().then(pool => {

let querycode = 'SELECT * from Products_NEW WHERE ProductID = ' + "'" + req.query.ProductID + "'";
     
return pool.query(querycode);
}).then(result => {

  
res.render('Update', {title: "Update", ProductID: req.query.ProductID, data: result.recordset[0], plantNums: plantNums, recordFacs:result.recordset[0].Facility.split(",")})

}).catch(err => {
     res.render('failure', { title: "Failure", err: err})
     
    console.log(err); 
     
     sendEmail(app.locals.currentUser, err, (req.protocol + '://' + req.get('host') + req.originalUrl), "Your request was not successful. Please see error details below.", "ICLMSAdmins@resers.com").catch(console.error);
}); }
}); 

// history page which shows an entire snapshot of a record at a certain moment
app.get('/history', function(req, res) { 
    
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

// pass params and return records from Products table
new sql.ConnectionPool(model.config).connect().then(pool => {
    
let querycode = 'SELECT * from Products_Changes where ID = ' + "'" + req.query.id + "'";

return pool.query(querycode);
}).then(result => {   
    
res.render('history', { title: "History", data: result.recordset[0]})
    
}).catch(err => {
    res.render('failure', { title: "Failure", err: err})
    
    console.log(err);
    
    sendEmail(app.locals.currentUser, err, (req.protocol + '://' + req.get('host') + req.originalUrl), "Your request was not successful. Please see error details below.", "ICLMSAdmins@resers.com").catch(console.error);
}); 
}); 

//IsActive page; sets IsActive column to 0 or 1 based on a button click
app.get('/ishidden', function(req, res) { 
    if (req.query.bit) {

new sql.ConnectionPool(model.config).connect().then(pool => {
    
let querycode = '';
    
if (req.query.bit == '0') {
querycode = 'UPDATE Products_NEW set IsActive = 0 WHERE ProductID = ' + "'" + req.query.ProductID + "'";
}
if (req.query.bit == '1') {    
querycode = 'UPDATE Products_NEW set IsActive = 1 WHERE ProductID = ' + "'" + req.query.ProductID + "'";
}
return pool.query(querycode);
}).then(result => { 
    
res.redirect('details?ProductID=' + req.query.ProductID)
    
}).catch(err => {
console.log(err);   
});} 
});

//email page; sends email notification if change happens to any of the critical fields
app.get('/email', function(req, res) { 
    
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

sendEmail(app.locals.currentUser, "No Errors Found.", (req.protocol + '://' + req.get('host') + "/update?ProductID=" + req.query.ProductID), "A change has been made to either the Part Number, and/or the Finished Goods Number, and/or the Ingredients for Product ID #" + req.query.ProductID, "ICLMSNotifications@resers.com").catch(console.error);

res.render('success', { title: "Success", ProductID: req.query.ProductID}) 
});

    // pulls sql records for weekly plant report and puts them into csv format
app.get('/csv', function(req,res) {
    let connection = new sql.ConnectionPool(model.config, function (err) {
		if (err) {
			console.log(err.message);
		} else {

		//You can pipe mssql request as per docs
		var request = new sql.Request(connection);
		request.stream = true;
		request.query('SELECT [Brand],[Country],[Facility],[PartNum],' + "('''' + convert(varchar(20),[FG]) + '''') as FG" +',' + "('''' + [GTIN] + '''') as GTIN" +',[LxDescription],[LabelDescription],[SubDescription],[Formula],[Shelf],[Guarantee],[SafeHandling],[NetWeight],[WeightUOM],[PackSize],[NetContent],[NetWeightUOM],[DateCode],[CADDateCode],[DistrStatement],[Logo],[Ingredients],[InspectionLegend],[InspectionCode],[LastUpdatedBy],[IsActive],[NutritionImage], [UnitUPC] from Products_NEW where (GTIN like ' + "'" + req.query.gtin + "'" + ' or FG like ' +  "'" + req.query.fg + "'" + ' or LxDescription like ' +  "'" + req.query.lx + "'" + ' or Formula like ' +  "'" + req.query.formula + "')");

		var stringifier = csv.stringify({header: true});
		//Then simply call attachment and pipe it to the response
        let filename = 'ICLMS_Search_Results_Generated_on_' + (new Date().getFullYear()+'-'+(new Date().getMonth()+1)+'-'+new Date().getDate()) + '.csv';
		res.attachment(filename);
		request.pipe(stringifier).pipe(res);
        }
	});
})  
    
// success page
app.get('/success', function(req, res) { 
    
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

res.render('success', { title: "Success", ProductID: req.query.ProductID })
}); 

// failure page
app.get('/failure', function(req, res) { 
    
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

res.render('failure', { title: "Failure"})
}); 

// 404 handler
app.use(function(req,res) {
    
app.locals.currentUser = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
    
  res.render('404', {title: "404 - Page Not Found"} ); 
});

app.listen(app.get('port'), function() {
 console.log('App has started');
});
