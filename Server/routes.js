/*
 * GET home page.
 */

var routesIndex = {

 init: function (app) {

    var CreateSerial = function (clientId) {
      return clientId;
    }
    
    // Routes
    app.get('/', function(req, res) {
      res.render('index', { 
        login: "none",
        user: {}
      });
    });

    app.get('/install', function(req, res){

      var clientId = req.params.clientId;
      var todayDate = new Date();
      var endDate = new Date();
      endDate.setDate(todayDate.getDate()+30);

      var serialNumber = CreateSerial(clientId);

      res.json({
          serial: serialNumber,
          endDate: endDate
        });  
      
    });

  }

}

module.exports = routesIndex;
