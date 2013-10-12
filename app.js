/**
 * Module dependencies.
 */
// var express = require('express'),
//     passport = require('passport'),
//     TwitterStrategy = require('passport-twitter').Strategy,
//     api = require('./routes/api'),
//     routesIndex = require('./routes/index'),
//     dataModel = require('./mongoDataModel'),
//     config = require('./config');

var express = require('express'),
    routesIndex = require('./routes/index');
    
var app = module.exports = express();

// Configuration
app.configure(function(){
  app.set('views', __dirname + '/views');
  app.engine('.html', require('ejs').renderFile);
  app.set('view engine', 'html');
  app.set('view options', {
    layout: false
  });
  app.use(express.bodyParser());
  app.use(express.methodOverride());
  app.use(express.cookieParser());
  app.use(express.session({ secret: 'omgnodeworks' }));
  app.use(app.router);
  app.use(express.static(__dirname + '/public'));
});

app.configure('development', function(){
  app.use(express.errorHandler({ dumpExceptions: true, showStack: true }));
});

app.configure('production', function(){
  app.use(express.errorHandler());
});

routesIndex.init(app, dbConnection);


appServer = app.listen(config.express.port, function(){
  console.log("Express server listening on port %d in %s mode", appServer.address().port, app.settings.env);
});
