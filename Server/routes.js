/*
 * GET home page.
 */

var routesIndex = {

 init: function (app) {
    // Routes
    app.get('/', function(req, res) {
      res.render('index', { 
        login: "none",
        user: {}
      });
    });

    app.get('/login', function(req, res){
      if ( typeof(req.user) != "undefined" ) {

        db.getUser(req.user.id, function (err, user) {
          res.render('index', {
            login: "success",
            user: {
              id: user._id,
              uid: req.user.id,
              name: req.user.displayName,
              image: req.user.photos[0].value,
              profileURL: "http://twitter.com/" + req.user.username
            }
          });  
        });

        
      } else {
        res.render('error', {});
      }
    });

    app.get('/loginfail', function(req, res){
      res.render('index', {
        login: "failure",
        user: {}
      });
    });

    app.get('/partials/:name', function (req, res) {
      var name = req.params.name;
      res.render('partials/' + name);
    });

    // Redirect the user to Twitter for authentication.  When complete, Twitter
    // will redirect the user back to the application at
    //   /auth/twitter/callback
    app.get('/auth/twitter', passport.authenticate('twitter'));

    // Twitter will redirect the user to this URL after approval.  Finish the
    // authentication process by attempting to obtain an access token.  If
    // access was granted, the user will be logged in.  Otherwise,
    // authentication has failed.
    app.get('/auth/twitter/callback', 
      passport.authenticate('twitter', { successRedirect: '/login',
                                         failureRedirect: '/loginfail' }));

    var business = api(db);

    app.get('/trades/open',business.tradeOpen);
    app.get('/trades/add',business.tradeAdd);
    app.get('/products/all',business.productsAll);
  }

}

module.exports = routesIndex;
