// require in our request library.
var request = require("request");

// exports.handler must match the name of the handler given to AWS in the Handler section below Function Code
exports.handler = function(event, context, callback) {

  // create our URL, you could make this dynamic with template literals and use the info you provide it via a POST (we'll set
  // up our HTTP methods in part 2), however for this simple example we're hard coding just to show you how it all works.
  
  var url = 'https://swapi.co/api/people/1'; 
	
  // execute our request to the Star Wars API.
  request({
    url: url,
    json: true,
  }, function (error, response, body) {
    
  // ensure there are no errors.
    if (!error && response.statusCode === 200) {
        
      //log the response for debug purposes feel free to delete if working correctly.
      console.log(body)

      // create our response with appropriate headers, status code and the data received from the Star Wars API.
      const apiResponse = {
        statusCode: 200,
        headers: {
          "Access-Control-Allow-Origin" : "*", // Required for CORS support to work.
          "Access-Control-Allow-Credentials" : true // Required for cookies, authorization headers with HTTPS.
        },
        body: JSON.stringify(body)
      };
        
      // log our apiResponse object for debug purposes feel free to delete if working correctly.
      console.log('apiResponse', apiResponse)
        
      // return the apiResponse object back to our API call.
      callback(null, apiResponse);
    }   
  });
};